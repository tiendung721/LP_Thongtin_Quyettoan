"""Persist reviewed revenue rows independently of the daily Excel workflow."""

from __future__ import annotations

import hashlib
import json
import re
import sqlite3
import unicodedata
from collections import defaultdict
from contextlib import contextmanager
from datetime import datetime
from decimal import Decimal, InvalidOperation
from pathlib import Path
from typing import Any, Iterator


def _now() -> str:
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def _plain(value: Any) -> str:
    text = unicodedata.normalize("NFD", str(value or "").casefold())
    return "".join(char for char in text if unicodedata.category(char) != "Mn").replace("đ", "d")


def is_revenue_row(row: dict[str, Any]) -> bool:
    return _plain(row.get("loai_chung_tu")).strip() == "doanh thu"


def revenue_rows(payload: dict[str, Any]) -> list[dict[str, Any]]:
    return [row for row in payload.get("du_lieu_boc_tach", []) if isinstance(row, dict) and is_revenue_row(row)]


def parse_amount_vnd(value: Any) -> int | None:
    """Read integer VND without letting SQLite silently store a float."""
    if value in (None, "") or isinstance(value, bool):
        return None
    if isinstance(value, (int, float, Decimal)):
        number = Decimal(str(value))
    else:
        raw = re.sub(r"\s|[₫đ]", "", str(value))
        if not raw:
            return None
        if re.fullmatch(r"\d{1,3}(?:[.,]\d{3})+(?:[.,]00?)?", raw):
            parts = re.split(r"[.,]", raw)
            if parts[-1] in ("0", "00"):
                parts.pop()
            raw = "".join(parts)
        elif re.fullmatch(r"\d+[.,]0+", raw):
            raw = raw.split(".")[0].split(",")[0]
        try:
            number = Decimal(raw)
        except InvalidOperation:
            return None
    if not number.is_finite() or number != number.to_integral_value():
        return None
    result = int(number)
    return result if 0 <= result <= 9_223_372_036_854_775_807 else None


def normalize_sqt(value: Any) -> str:
    text = str(value or "").strip()
    return text[:-2] if re.fullmatch(r"\d+\.0", text) else text


def normalize_container(value: Any) -> str:
    return re.sub(r"[\s\-]", "", str(value or "")).upper()


def review_errors(partner: str, sqt: str, container: str, amount_vnd: int | None) -> list[str]:
    errors = []
    if not partner.strip():
        errors.append("Thiếu đối tác")
    if not sqt.strip() or not sqt.isdigit() or int(sqt) <= 0:
        errors.append("SQT phải là số dương")
    if not re.fullmatch(r"[A-Z]{4}\d{7}", container):
        errors.append("Container phải gồm 4 chữ và 7 số")
    if amount_vnd is None or amount_vnd <= 0:
        errors.append("Thiếu thành tiền gồm VAT hợp lệ")
    return errors


class RevenueStore:
    """Revenue tables live in app_state.db but never touch the daily Excel tables."""

    def __init__(self, db_path: str):
        self.db_path = db_path

    @contextmanager
    def _connect(self) -> Iterator[sqlite3.Connection]:
        conn = sqlite3.connect(self.db_path)
        conn.row_factory = sqlite3.Row
        conn.execute("PRAGMA foreign_keys = ON")
        try:
            yield conn
            conn.commit()
        finally:
            conn.close()

    def init_db(self) -> None:
        with self._connect() as conn:
            conn.executescript("""
                CREATE TABLE IF NOT EXISTS revenue_sources (
                    id INTEGER PRIMARY KEY,
                    source_key TEXT NOT NULL UNIQUE,
                    source_name TEXT NOT NULL,
                    source_md5 TEXT,
                    period TEXT,
                    first_seen_at TEXT NOT NULL
                );
                CREATE TABLE IF NOT EXISTS revenue_rows (
                    id INTEGER PRIMARY KEY,
                    source_id INTEGER NOT NULL REFERENCES revenue_sources(id),
                    source_ref TEXT NOT NULL,
                    source_sheet TEXT,
                    source_page TEXT,
                    source_line TEXT,
                    partner TEXT NOT NULL DEFAULT '',
                    sqt TEXT NOT NULL DEFAULT '',
                    container TEXT NOT NULL DEFAULT '',
                    amount_vnd INTEGER,
                    raw_json TEXT NOT NULL,
                    warning_json TEXT NOT NULL DEFAULT '[]',
                    warning_resolved INTEGER NOT NULL DEFAULT 0,
                    review_status TEXT NOT NULL CHECK(review_status IN ('OK','CAN_KIEM_TRA')),
                    import_status TEXT NOT NULL DEFAULT 'CHUA_NHAP'
                        CHECK(import_status IN ('CHUA_NHAP','DA_NHAP')),
                    created_at TEXT NOT NULL,
                    updated_at TEXT NOT NULL,
                    UNIQUE(source_id, source_ref)
                );
                CREATE TABLE IF NOT EXISTS revenue_row_events (
                    id INTEGER PRIMARY KEY,
                    revenue_row_id INTEGER NOT NULL REFERENCES revenue_rows(id),
                    before_json TEXT NOT NULL,
                    after_json TEXT NOT NULL,
                    changed_at TEXT NOT NULL
                );
                CREATE INDEX IF NOT EXISTS idx_revenue_group
                    ON revenue_rows(partner, sqt);
                CREATE INDEX IF NOT EXISTS idx_revenue_state
                    ON revenue_rows(review_status, import_status);
            """)

    def ingest_payload(self, payload: dict[str, Any], output_path: str = "") -> int:
        """Insert new GPT rows, preserving all user edits on repeated imports."""
        rows = revenue_rows(payload)
        if not rows:
            return 0
        occurrences: dict[tuple[str, str], int] = defaultdict(int)
        added = 0
        with self._connect() as conn:
            for row in rows:
                extra = row.get("doanh_thu") if isinstance(row.get("doanh_thu"), dict) else {}
                source_name = str(row.get("file_nguon") or Path(output_path).name or "Không rõ file nguồn")
                md5 = str(row.get("ma_md5_file") or "").strip().lower()
                if re.fullmatch(r"[0-9a-f]{32}", md5):
                    source_key = "md5:" + md5
                else:
                    # The GPT instructions require MD5; this fallback keeps an
                    # unidentifiable source separate and reviewable.
                    source_key = "fallback:" + hashlib.sha256(
                        (source_name + "|" + output_path).encode("utf-8")
                    ).hexdigest()
                sheet = str(extra.get("sheet_nguon") or "").strip()
                page = str(extra.get("trang_nguon") or "").strip()
                line = str(extra.get("dong_nguon") or "").strip()
                partner = str(extra.get("doi_tac") or "").strip()
                sqt = normalize_sqt(extra.get("sqt"))
                container = normalize_container(row.get("so_container"))
                amount = parse_amount_vnd(extra.get("thanh_tien_gom_vat"))
                location = "|".join((sheet, page, line)) if any((sheet, page, line)) else ""
                identity = location or "|".join((partner.casefold(), sqt, container))
                occurrence_key = (source_key, identity)
                occurrences[occurrence_key] += 1
                source_ref = f"{identity}|{occurrences[occurrence_key]}"
                now = _now()
                conn.execute(
                    "INSERT OR IGNORE INTO revenue_sources "
                    "(source_key,source_name,source_md5,period,first_seen_at) VALUES (?,?,?,?,?)",
                    (source_key, source_name, md5 or None, str(extra.get("ky_du_lieu") or ""), now),
                )
                source_id = conn.execute(
                    "SELECT id FROM revenue_sources WHERE source_key=?", (source_key,)
                ).fetchone()[0]
                warnings = row.get("canh_bao") if isinstance(row.get("canh_bao"), list) else []
                warnings = list(warnings)
                duplicate = conn.execute(
                    "SELECT id FROM revenue_rows WHERE source_id != ? AND "
                    "partner = ? AND sqt = ? AND container = ? LIMIT 1",
                    (source_id, partner, sqt, container),
                ).fetchone()
                if duplicate:
                    warnings.append("Nghi trùng doanh thu với file nguồn đã lưu; cần đối chiếu.")
                valid_md5 = bool(re.fullmatch(r"[0-9a-f]{32}", md5))
                if not valid_md5:
                    warnings.append("Không xác minh được MD5 file nguồn.")
                needs_attention = bool(warnings) or _plain(row.get("do_tin_cay")) == "thap" or _plain(row.get("trang_thai")) not in ("", "ok")
                status = "CAN_KIEM_TRA" if review_errors(partner, sqt, container, amount) or needs_attention else "OK"
                cursor = conn.execute(
                    "INSERT OR IGNORE INTO revenue_rows "
                    "(source_id,source_ref,source_sheet,source_page,source_line,partner,sqt,container,"
                    "amount_vnd,raw_json,warning_json,review_status,created_at,updated_at) "
                    "VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?)",
                    (source_id, source_ref, sheet, page, line, partner, sqt, container,
                     amount, json.dumps(row, ensure_ascii=False), json.dumps(warnings, ensure_ascii=False),
                     status, now, now),
                )
                added += cursor.rowcount
        return added

    def list_rows(self) -> list[dict[str, Any]]:
        with self._connect() as conn:
            rows = conn.execute(
                "SELECT r.*,s.source_name,s.period FROM revenue_rows r "
                "JOIN revenue_sources s ON s.id=r.source_id ORDER BY r.id DESC"
            ).fetchall()
        return [dict(row) for row in rows]

    def save_row(
        self, row_id: int, *, partner: str, sqt: str, container: str,
        amount: Any, warning_resolved: bool, import_status: str,
    ) -> dict[str, Any]:
        if import_status not in ("CHUA_NHAP", "DA_NHAP"):
            raise ValueError("Trạng thái nhập không hợp lệ.")
        amount_vnd = parse_amount_vnd(amount)
        if str(amount).strip() and amount_vnd is None:
            raise ValueError("Thành tiền gồm VAT phải là số nguyên VND hợp lệ.")
        partner = partner.strip()
        sqt = normalize_sqt(sqt)
        container = normalize_container(container)
        with self._connect() as conn:
            existing = conn.execute("SELECT * FROM revenue_rows WHERE id=?", (row_id,)).fetchone()
            if existing is None:
                raise ValueError("Không còn dòng doanh thu này.")
            before = dict(existing)
            warning_items = json.loads(before["warning_json"])
            raw = json.loads(before["raw_json"])
            source_alert = bool(warning_items) or _plain(raw.get("do_tin_cay")) == "thap" or _plain(raw.get("trang_thai")) not in ("", "ok")
            status = "CAN_KIEM_TRA" if review_errors(partner, sqt, container, amount_vnd) or (source_alert and not warning_resolved) else "OK"
            after = {**before, "partner": partner, "sqt": sqt, "container": container,
                     "amount_vnd": amount_vnd, "warning_resolved": int(warning_resolved),
                     "review_status": status, "import_status": import_status}
            changed = any(after[key] != before[key] for key in (
                "partner", "sqt", "container", "amount_vnd", "warning_resolved",
                "review_status", "import_status",
            ))
            if changed:
                now = _now()
                conn.execute(
                    "UPDATE revenue_rows SET partner=?,sqt=?,container=?,amount_vnd=?,"
                    "warning_resolved=?,review_status=?,import_status=?,updated_at=? WHERE id=?",
                    (partner, sqt, container, amount_vnd, int(warning_resolved), status,
                     import_status, now, row_id),
                )
                keys = ("partner", "sqt", "container", "amount_vnd", "warning_resolved", "review_status", "import_status")
                conn.execute(
                    "INSERT INTO revenue_row_events (revenue_row_id,before_json,after_json,changed_at) "
                    "VALUES (?,?,?,?)",
                    (row_id, json.dumps({key: before[key] for key in keys}, ensure_ascii=False),
                     json.dumps({key: after[key] for key in keys}, ensure_ascii=False), now),
                )
            return after
