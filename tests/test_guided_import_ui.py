from __future__ import annotations

import json
import logging
import os
import tempfile
import unittest
from datetime import datetime
from pathlib import Path

from openpyxl import Workbook

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

from PySide6.QtWidgets import QApplication, QDialog  # noqa: E402

from app.daily_import import EXPENSE_HEADERS, EXPENSE_SHEET, INFO_HEADERS, DailyImportService  # noqa: E402
from app.database import Database  # noqa: E402
from app.guided_import_ui import GuidedImportResolutionDialog  # noqa: E402


class GuidedImportDialogTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.app = QApplication.instance() or QApplication([])

    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        root = Path(self.temp.name)
        self.output = root / "output.json"
        self.daily = root / "daily.xlsx"
        db = Database(str(root / "state.db"))
        db.init_db()
        self.db = db
        self.service = DailyImportService(db, logging.getLogger("guided-import-test"))

        workbook = Workbook()
        info = workbook.active
        info.title = "Tháng 6"
        info.append(INFO_HEADERS)
        existing = {"SQT PM": 250, "Ngày Đóng": datetime(2026, 6, 15), "Số Container": "IIJJ1234567"}
        info.append([existing.get(name) for name in INFO_HEADERS])
        expense = workbook.create_sheet(EXPENSE_SHEET)
        expense.append(EXPENSE_HEADERS)
        workbook.save(self.daily)
        workbook.close()

        rows = [
            {"file_nguon": "scale.jpg", "ma_md5_file": "scale-a1", "loai_chung_tu": "Phiếu cân",
             "trang_thai": "OK", "ngay_dong": "17/06/2026", "so_container": "MMNN1234567", "so_tan": 29},
            {"file_nguon": "bill-a.pdf", "ma_md5_file": "bill-a", "loai_chung_tu": "Bill",
             "trang_thai": "OK", "so_container": "MMNN1234567", "so_bill": "BL-A", "ten_tau": "Tàu A"},
            {"file_nguon": "bill-b.pdf", "ma_md5_file": "bill-b", "loai_chung_tu": "Bill",
             "trang_thai": "OK", "so_container": "MMNN1234567", "so_bill": "BL-B", "ten_tau": "Tàu B"},
            {"file_nguon": "cost.jpg", "ma_md5_file": "cost-b1", "loai_chung_tu": "Khoản chi",
             "trang_thai": "OK", "ngay_chay": "14/06/2026", "ngay_dong": "15/06/2026",
             "so_container": "IIJJ1234567", "so_hd": "HD-88", "don_gia": 110_000},
        ]
        self.output.write_text(json.dumps({"metadata": {}, "du_lieu_boc_tach": rows}, ensure_ascii=False), encoding="utf-8")
        analysis = self.service.analyze(str(self.output), str(self.daily))
        self.dialog = GuidedImportResolutionDialog(self.service, analysis)

    def tearDown(self):
        self.dialog.accept()
        self.temp.cleanup()

    def test_edit_uses_sail_date_and_unselected_bill_remains_pending(self):
        self.assertEqual(len(self.dialog.analysis.unresolved_issues), 2)
        self.assertEqual(self.dialog.edit_table.item(0, 2).text(), "Ngày chạy")
        self.assertFalse(self.dialog.edit_apply.isEnabled())

        self.dialog.edit_table.item(0, 3).setText("15/06/2026")
        self.assertTrue(self.dialog.edit_apply.isEnabled())
        self.dialog.apply_edits()
        self.assertEqual(self.dialog.sandbox.payload["du_lieu_boc_tach"][3]["ngay_chay"], "15/06/2026")
        self.assertEqual(self.dialog.sandbox.payload["du_lieu_boc_tach"][3]["ngay_dong"], "15/06/2026")
        self.assertEqual(self.dialog.tabs.currentIndex(), 1)

        self.dialog.choice_table.selectRow(0)
        self.dialog.choose_from_preview(1, 0)
        self.assertEqual(self.dialog.choice_preview.item(1, 1).text(), "BL-B")
        self.assertTrue(self.dialog.choice_apply.isEnabled())
        self.dialog.apply_choices()
        self.assertFalse(self.dialog.analysis.unresolved_issues)
        self.assertIn("BL-A", self.dialog.review_attention.text())

        self.dialog.finish()
        self.assertEqual(self.dialog.result(), QDialog.Accepted)
        summary = self.service.commit(self.dialog.final_analysis)
        self.assertEqual(summary.status, "DAILY_IMPORT_PARTIAL")
        self.assertGreater(summary.pending_count, 0)

    def test_unapplied_choice_cannot_be_saved_as_a_different_decision(self):
        self.dialog.choice_table.selectRow(0)
        self.dialog.choose_from_preview(0, 0)
        self.dialog.save_and_close()
        self.assertFalse(self.dialog.saved_only)
        self.assertEqual(self.dialog.tabs.currentIndex(), 1)
        self.assertIn("chưa ghép", self.dialog.feedback.text())

    def test_save_for_later_keeps_workbook_unchanged_and_pending_rows(self):
        before = self.daily.read_bytes()
        self.dialog.choice_table.selectRow(0)
        self.dialog.choose_from_preview(1, 0)
        self.dialog.apply_choices()
        self.dialog.save_and_close()

        self.assertTrue(self.dialog.saved_only)
        self.assertEqual(self.dialog.result(), QDialog.Accepted)
        self.assertEqual(self.daily.read_bytes(), before)
        self.assertGreater(self.db.count_pending_rows(), 0)


if __name__ == "__main__":
    unittest.main()
