from __future__ import annotations

import os
import json
import logging
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

from PySide6.QtWidgets import QApplication

from app.dialogs.revenue_dialog import RevenueDialog
from app.daily_import_ui import JsonExtractDialog
from app.config import AppConfig, get_default_settings
from app.database import Database
from app.main_window import MainWindow
from app.revenue import RevenueStore, parse_amount_vnd


def _row(sqt: str, container: str, amount, line: str, *, status: str = "OK") -> dict:
    return {
        "file_nguon": "Thang 10.xls",
        "ma_md5_file": "a" * 32,
        "loai_chung_tu": "Doanh thu",
        "trang_thai": status,
        "so_container": container,
        "canh_bao": [],
        "doanh_thu": {
            "sqt": sqt, "doi_tac": "Mía đường TN", "thanh_tien_gom_vat": amount,
            "sheet_nguon": "Mía đường TN", "dong_nguon": line,
        },
    }


class RevenueTests(unittest.TestCase):
    def setUp(self) -> None:
        self.temp = tempfile.TemporaryDirectory()
        self.store = RevenueStore(str(Path(self.temp.name) / "app_state.db"))
        self.store.init_db()

    def tearDown(self) -> None:
        self.temp.cleanup()

    def test_money_parses_vnd_without_float_rounding(self) -> None:
        self.assertEqual(parse_amount_vnd("75.239.928"), 75239928)
        self.assertEqual(parse_amount_vnd("75,239,928"), 75239928)
        self.assertEqual(parse_amount_vnd("75239928.0"), 75239928)
        self.assertIsNone(parse_amount_vnd("75,239.50"))

    def test_mixed_payload_keeps_incomplete_rows_and_reimport_keeps_edits(self) -> None:
        payload = {"du_lieu_boc_tach": [
            {"loai_chung_tu": "Bill", "so_container": "GAOU2133678"},
            _row("920", "VSGU2231272", 75239928, "9"),
            _row("921", "GAOU2133678", None, "10"),
        ]}
        self.assertEqual(self.store.ingest_payload(payload, "output.json"), 2)
        rows = self.store.list_rows()
        self.assertEqual({row["review_status"] for row in rows}, {"OK", "CAN_KIEM_TRA"})
        missing = next(row for row in rows if row["sqt"] == "921")
        self.store.save_row(
            missing["id"], partner="Mía đường TN", sqt="921",
            container="GAOU2133678", amount="77.497.560", warning_resolved=True,
            import_status="DA_NHAP",
        )
        self.assertEqual(self.store.ingest_payload(payload, "output.json"), 0)
        restored = next(row for row in self.store.list_rows() if row["sqt"] == "921")
        self.assertEqual(restored["amount_vnd"], 77497560)
        self.assertEqual(restored["review_status"], "OK")
        self.assertEqual(restored["import_status"], "DA_NHAP")
        with self.store._connect() as conn:
            self.assertEqual(conn.execute("SELECT COUNT(*) FROM revenue_row_events").fetchone()[0], 1)

    def test_amended_source_flags_possible_duplicate(self) -> None:
        row = _row("920", "VSGU2231272", 75239928, "9")
        self.store.ingest_payload({"du_lieu_boc_tach": [row]}, "first.json")
        amended = dict(row, ma_md5_file="b" * 32)
        self.store.ingest_payload({"du_lieu_boc_tach": [amended]}, "second.json")
        rows = self.store.list_rows()
        self.assertEqual(len(rows), 2)
        self.assertEqual(rows[0]["review_status"], "CAN_KIEM_TRA")
        self.assertIn("Nghi trùng", rows[0]["warning_json"])

    def test_dialog_edits_one_container_and_keeps_two_statuses_separate(self) -> None:
        app = QApplication.instance() or QApplication([])
        self.store.ingest_payload({"du_lieu_boc_tach": [
            _row("920", "VSGU2231272", 75239928, "9"),
            _row("920", "GAOU2133678", None, "10"),
        ]}, "output.json")
        dialog = RevenueDialog(self.store)
        self.assertEqual(dialog.group_table.rowCount(), 1)
        self.assertEqual(dialog.line_table.rowCount(), 2)
        dialog.select_visible.click()
        self.assertEqual(len(dialog.selected_ids), 1)
        self.assertEqual(
            next(row for row in dialog.rows if row["id"] in dialog.selected_ids)["container"],
            "VSGU2231272",
        )
        missing_index = next(
            index for index, row in enumerate(dialog.visible_lines)
            if row["container"] == "GAOU2133678"
        )
        dialog.line_table.selectRow(missing_index)
        dialog.amount_edit.setText("77497560")
        dialog.import_combo.setCurrentIndex(dialog.import_combo.findData("DA_NHAP"))
        dialog.save_current()
        rows = self.store.list_rows()
        changed = next(row for row in rows if row["container"] == "GAOU2133678")
        self.assertEqual(changed["review_status"], "OK")
        self.assertEqual(changed["import_status"], "DA_NHAP")
        self.assertIn("Nhập một phần", dialog.group_table.item(0, 6).text())
        dialog.close()
        del app

    def test_old_review_screen_preserves_but_does_not_show_revenue(self) -> None:
        app = QApplication.instance() or QApplication([])
        path = Path(self.temp.name) / "extract.json"
        payload = {"metadata": {}, "du_lieu_boc_tach": [
            {"loai_chung_tu": "Bill", "file_nguon": "bill.pdf", "so_container": "VSGU2231272"},
            _row("920", "VSGU2231272", 75239928, "9"),
        ]}
        path.write_text(json.dumps(payload, ensure_ascii=False), encoding="utf-8")
        dialog = JsonExtractDialog(str(path))
        self.assertEqual(len(dialog.rows), 1)
        self.assertEqual(dialog.summary_values["total"].text(), "1")
        dialog.save()
        saved = json.loads(path.read_text(encoding="utf-8"))
        self.assertEqual(len(saved["du_lieu_boc_tach"]), 2)
        self.assertEqual(saved["du_lieu_boc_tach"][1]["doanh_thu"]["sqt"], "920")
        del app

    def test_new_revenue_output_routes_to_step_four_without_old_review(self) -> None:
        app = QApplication.instance() or QApplication([])
        root = Path(self.temp.name)
        config = AppConfig(get_default_settings(str(root)), str(root / "Config" / "settings.json"))
        config.ensure_folders()
        database = Database(config.database_path)
        database.init_db()
        path = root / "output" / "boc_tach_doanh_thu.json"
        path.write_text(json.dumps({"metadata": {}, "du_lieu_boc_tach": [
            _row("920", "VSGU2231272", 75239928, "9"),
        ]}, ensure_ascii=False), encoding="utf-8")
        record_id = database.insert_processed_file(str(path), "READY_FOR_REVIEW")
        with patch("app.main_window.OutputWatcher.start"), patch("app.main_window.OutputWatcher.stop"), patch.object(
            MainWindow, "_start_excel_lock_check"
        ), patch.object(MainWindow, "_open_extract_editor") as old_review:
            window = MainWindow(config, database, logging.getLogger("test_revenue_ui"))
            window.on_output_ready({
                "id": record_id, "working_path": str(path), "file_name": path.name,
                "status": "READY_FOR_REVIEW",
            })
            self.assertEqual(len(window.revenue_store.list_rows()), 1)
            self.assertFalse(window.btn_view_extract.isEnabled())
            self.assertFalse(window.btn_import_daily.isEnabled())
            self.assertIn("1 dòng doanh thu", window.lbl_input_revenue_count.text())
            old_review.assert_not_called()
            window.close()
        del app


if __name__ == "__main__":
    unittest.main()
