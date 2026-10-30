"""Prepare revenue batches for PAD and record confirmed web saves."""

from __future__ import annotations

import json
import os
from datetime import datetime
from pathlib import Path
from typing import Iterable

from .revenue import RevenueStore

OPERATION = "NHAP_DOANH_THU"
MAX_SAFE_INTEGER = 9_007_199_254_740_991


class RevenuePadError(ValueError):
    """A revenue batch cannot be sent or confirmed safely."""


def build_selection(store: RevenueStore, selected_ids: Iterable[int]) -> dict:
    """Selecting one row means importing the full total for its SQT."""
    rows = store.list_rows()
    by_id = {row["id"]: row for row in rows}
    selected = {int(row_id) for row_id in selected_ids}
    if not selected:
        raise RevenuePadError("Hãy chọn ít nhất một dòng doanh thu.")
    if selected - by_id.keys():
        raise RevenuePadError("Danh sách doanh thu đã thay đổi. Hãy mở lại cửa sổ.")
    for row_id in selected:
        row = by_id[row_id]
        if row["review_status"] != "OK" or row["import_status"] != "CHUA_NHAP":
            raise RevenuePadError(f"Dòng {row_id} chưa đủ điều kiện nhập.")

    items = []
    for sqt in sorted({by_id[row_id]["sqt"] for row_id in selected}, key=lambda value: int(value)):
        group = [row for row in rows if row["sqt"] == sqt]
        if any(row["review_status"] != "OK" or not row["amount_vnd"] for row in group):
            raise RevenuePadError(f"SQT {sqt} còn dòng cần kiểm tra; chưa thể tính tổng.")
        containers = {row["container"] for row in group}
        amount = sum(row["amount_vnd"] for row in group)
        if amount > MAX_SAFE_INTEGER:
            raise RevenuePadError(f"Tổng doanh thu SQT {sqt} vượt giới hạn số nguyên an toàn.")
        items.append({
            "sqt": sqt,
            "amount_vnd": amount,
            "containers_key": "|" + "|".join(sorted(containers)) + "|",
            "rows": [
                {
                    "id": row["id"], "amount_vnd": row["amount_vnd"],
                    "container": row["container"], "import_status": row["import_status"],
                }
                for row in group
            ],
        })
    return {
        "operation": OPERATION,
        "database_path": str(Path(store.db_path).resolve()),
        "created_at": datetime.now().isoformat(timespec="seconds"),
        "items": items,
    }


def read_selection(path: str | os.PathLike[str]) -> dict:
    try:
        payload = json.loads(Path(path).read_text(encoding="utf-8-sig"))
    except (OSError, json.JSONDecodeError) as exc:
        raise RevenuePadError(f"Không đọc được file chọn doanh thu: {exc}") from exc
    if payload.get("operation") != OPERATION or not isinstance(payload.get("items"), list):
        raise RevenuePadError("File chọn không thuộc luồng doanh thu.")
    return payload


def validate_selection(payload: dict, db_path: str) -> None:
    """Reject an old batch if its total or any source row has changed."""
    if Path(payload.get("database_path", "")).resolve() != Path(db_path).resolve():
        raise RevenuePadError("File lượt PAD thuộc cơ sở dữ liệu khác.")
    store = RevenueStore(db_path)
    with store._connect() as conn:
        for item in payload["items"]:
            expected_rows = item.get("rows")
            if not isinstance(expected_rows, list) or not expected_rows:
                raise RevenuePadError("File lượt PAD thiếu các dòng doanh thu.")
            expected_ids = {row["id"] for row in expected_rows}
            current_ids = {
                row[0] for row in conn.execute(
                    "SELECT id FROM revenue_rows WHERE sqt=?", (item["sqt"],)
                ).fetchall()
            }
            if current_ids != expected_ids:
                raise RevenuePadError(f"SQT {item['sqt']} đã thêm hoặc bớt dòng doanh thu.")
            if sum(row["amount_vnd"] for row in expected_rows) != item.get("amount_vnd"):
                raise RevenuePadError(f"Tổng tiền SQT {item.get('sqt')} không khớp file lượt PAD.")
            for expected in expected_rows:
                row = conn.execute(
                    "SELECT sqt,container,amount_vnd,review_status,import_status "
                    "FROM revenue_rows WHERE id=?", (expected["id"],)
                ).fetchone()
                if row is None or (
                    row["sqt"] != item["sqt"]
                    or row["container"] != expected["container"]
                    or row["amount_vnd"] != expected["amount_vnd"]
                    or row["review_status"] != "OK"
                    or (
                        expected["import_status"] == "DA_NHAP"
                        and row["import_status"] != "DA_NHAP"
                    )
                ):
                    raise RevenuePadError(
                        f"Dữ liệu SQT {item['sqt']} đã thay đổi; không thể chạy tiếp lượt cũ."
                    )


def write_selection(path: str | os.PathLike[str], payload: dict) -> str:
    target = Path(path)
    if target.exists():
        raise RevenuePadError(
            f"Còn lượt PAD doanh thu chưa kết thúc: {target}. "
            "Hãy kiểm tra lượt trước trước khi chạy lượt mới."
        )
    target.parent.mkdir(parents=True, exist_ok=True)
    temporary = target.with_name(target.name + ".tmp")
    try:
        temporary.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
        os.replace(temporary, target)
    finally:
        temporary.unlink(missing_ok=True)
    return str(target)


def mark_sqt_imported(selection_path: str | os.PathLike[str], sqt: str) -> int:
    """Mark only unchanged rows after PAD confirms the website's value."""
    payload = read_selection(selection_path)
    matches = [item for item in payload["items"] if item.get("sqt") == sqt]
    if len(matches) != 1:
        raise RevenuePadError(f"SQT {sqt} không có duy nhất trong lượt PAD.")
    item = matches[0]
    store = RevenueStore(payload["database_path"])
    changes = []
    with store._connect() as conn:
        conn.execute("BEGIN IMMEDIATE")
        current_ids = {
            row[0] for row in conn.execute(
                "SELECT id FROM revenue_rows WHERE sqt=?", (sqt,)
            ).fetchall()
        }
        if current_ids != {expected["id"] for expected in item["rows"]}:
            raise RevenuePadError(f"SQT {sqt} đã thêm hoặc bớt dòng doanh thu.")
        for expected in item["rows"]:
            row = conn.execute(
                "SELECT * FROM revenue_rows WHERE id=?", (expected["id"],)
            ).fetchone()
            if row is None or (
                row["sqt"] != sqt
                or row["amount_vnd"] != expected["amount_vnd"]
                or row["container"] != expected["container"]
                or row["review_status"] != "OK"
                or (
                    expected["import_status"] == "DA_NHAP"
                    and row["import_status"] != "DA_NHAP"
                )
            ):
                raise RevenuePadError(f"Dữ liệu SQT {sqt} đã thay đổi trong lúc PAD chạy.")
            if row["import_status"] == "DA_NHAP":
                continue
            before = {"import_status": "CHUA_NHAP"}
            after = {"import_status": "DA_NHAP"}
            now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            conn.execute(
                "UPDATE revenue_rows SET import_status='DA_NHAP',updated_at=? WHERE id=?",
                (now, row["id"]),
            )
            conn.execute(
                "INSERT INTO revenue_row_events "
                "(revenue_row_id,before_json,after_json,changed_at) VALUES (?,?,?,?)",
                (row["id"], json.dumps(before), json.dumps(after), now),
            )
            changes.append(row["id"])
    return len(changes)


def finish_selection(selection_path: str | os.PathLike[str]) -> None:
    """Remove the batch only after all its rows have been marked imported."""
    payload = read_selection(selection_path)
    store = RevenueStore(payload["database_path"])
    with store._connect() as conn:
        for item in payload["items"]:
            for expected in item["rows"]:
                row = conn.execute(
                    "SELECT import_status FROM revenue_rows WHERE id=?", (expected["id"],)
                ).fetchone()
                if row is None or row["import_status"] != "DA_NHAP":
                    raise RevenuePadError("Lượt PAD chưa nhập xong tất cả SQT.")
    Path(selection_path).unlink()

