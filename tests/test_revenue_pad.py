from __future__ import annotations

import tempfile
import unittest
from pathlib import Path

from app.revenue import RevenueStore
from app.revenue_pad import (
    RevenuePadError, build_selection, finish_selection, mark_sqt_imported,
    validate_selection, write_selection,
)


def revenue_row(sqt: str, container: str, amount: int, line: str) -> dict:
    return {
        "file_nguon": "revenue.xlsx",
        "ma_md5_file": "a" * 32,
        "loai_chung_tu": "Doanh thu",
        "so_container": container,
        "doanh_thu": {
            "sqt": sqt,
            "thanh_tien_gom_vat": amount,
            "dong_nguon": line,
        },
    }


class RevenuePadTests(unittest.TestCase):
    def setUp(self) -> None:
        self.temp = tempfile.TemporaryDirectory()
        self.root = Path(self.temp.name)
        self.store = RevenueStore(str(self.root / "app_state.db"))
        self.store.init_db()
        self.store.ingest_payload({"du_lieu_boc_tach": [
            revenue_row("877", "GESU3629007", 50_000_000, "1"),
            revenue_row("877", "GESU3629007", 23_635_512, "2"),
            revenue_row("878", "VSUI2620950", 9_000_000, "3"),
        ]}, "revenue.json")
        self.path = self.root / "runtime" / "rpa_revenue_selection.json"

    def tearDown(self) -> None:
        self.temp.cleanup()

    def test_selected_row_sends_full_sqt_total_and_marks_all_rows(self) -> None:
        selected = next(row["id"] for row in self.store.list_rows() if row["sqt"] == "877")
        payload = build_selection(self.store, [selected])
        self.assertEqual(len(payload["items"]), 1)
        self.assertEqual(payload["items"][0]["amount_vnd"], 73_635_512)
        self.assertEqual(len(payload["items"][0]["rows"]), 2)
        write_selection(self.path, payload)
        self.assertEqual(mark_sqt_imported(self.path, "877"), 2)
        self.assertEqual(mark_sqt_imported(self.path, "877"), 0)
        finish_selection(self.path)
        self.assertFalse(self.path.exists())
        self.assertEqual(
            {row["import_status"] for row in self.store.list_rows() if row["sqt"] == "877"},
            {"DA_NHAP"},
        )
        self.assertEqual(
            next(row["import_status"] for row in self.store.list_rows() if row["sqt"] == "878"),
            "CHUA_NHAP",
        )

    def test_changed_row_blocks_status_update(self) -> None:
        selected = next(row["id"] for row in self.store.list_rows() if row["sqt"] == "877")
        payload = build_selection(self.store, [selected])
        write_selection(self.path, payload)
        row = next(row for row in self.store.list_rows() if row["id"] == selected)
        self.store.save_row(
            row["id"], partner=row["partner"], sqt=row["sqt"],
            container=row["container"], amount=99_000_000,
            import_status="CHUA_NHAP",
        )
        with self.assertRaises(RevenuePadError):
            mark_sqt_imported(self.path, "877")
        with self.assertRaises(RevenuePadError):
            validate_selection(payload, self.store.db_path)
        self.assertTrue(all(
            row["import_status"] == "CHUA_NHAP"
            for row in self.store.list_rows() if row["sqt"] == "877"
        ))

    def test_existing_imported_line_remains_in_sqt_total(self) -> None:
        group = [row for row in self.store.list_rows() if row["sqt"] == "877"]
        first = group[0]
        self.store.save_row(
            first["id"], partner=first["partner"], sqt=first["sqt"],
            container=first["container"], amount=first["amount_vnd"],
            import_status="DA_NHAP",
        )
        pending_id = next(row["id"] for row in self.store.list_rows()
                          if row["sqt"] == "877" and row["import_status"] == "CHUA_NHAP")
        payload = build_selection(self.store, [pending_id])
        self.assertEqual(payload["items"][0]["amount_vnd"], 73_635_512)
        validate_selection(payload, self.store.db_path)
        write_selection(self.path, payload)
        self.assertEqual(mark_sqt_imported(self.path, "877"), 1)

    def test_incomplete_batch_is_kept_for_retry(self) -> None:
        selected = next(row["id"] for row in self.store.list_rows() if row["sqt"] == "877")
        write_selection(self.path, build_selection(self.store, [selected]))
        with self.assertRaises(RevenuePadError):
            finish_selection(self.path)
        with self.assertRaises(RevenuePadError):
            write_selection(self.path, build_selection(self.store, [selected]))
        self.assertTrue(self.path.exists())

    def test_same_sqt_with_different_containers_is_summed(self) -> None:
        self.store.ingest_payload({"du_lieu_boc_tach": [
            revenue_row("877", "ABCD1234567", 1_000_000, "4"),
        ]}, "revenue.json")
        selected = next(row["id"] for row in self.store.list_rows() if row["sqt"] == "877")
        item = build_selection(self.store, [selected])["items"][0]
        self.assertEqual(item["amount_vnd"], 74_635_512)
        self.assertIn("|ABCD1234567|", item["containers_key"])


if __name__ == "__main__":
    unittest.main()
