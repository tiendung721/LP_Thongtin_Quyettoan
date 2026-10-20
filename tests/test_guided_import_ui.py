from __future__ import annotations

import json
import logging
import os
import tempfile
import unittest
from datetime import datetime
from pathlib import Path

from openpyxl import Workbook
from PySide6.QtCore import QItemSelectionModel

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

from PySide6.QtWidgets import QApplication, QDialog  # noqa: E402

from app.daily_import import (  # noqa: E402
    EXPENSE_HEADERS, EXPENSE_SHEET, INFO_HEADERS, DailyImportService,
    FieldConflict, ImportIssue, InfoChange, ISSUE_FIELD_CONFLICT,
)
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

    def test_choice_tables_show_candidates_and_update_the_import_selection(self):
        self.dialog.tabs.setCurrentIndex(1)
        self.dialog.show()
        self.app.processEvents()
        self.assertTrue(self.dialog.choice_table.isVisible())
        self.assertTrue(self.dialog.choice_preview.isVisible())
        self.assertLessEqual(self.dialog.choice_table.height(), 100)
        self.assertLessEqual(self.dialog.choice_table.rowHeight(0), 44)
        self.assertGreater(self.dialog.choice_preview.height(), self.dialog.choice_table.height())
        self.assertLessEqual(self.dialog.choice_preview.rowHeight(0), 44)
        self.assertEqual(self.dialog.choice_table.horizontalScrollBar().maximum(), 0)
        self.dialog.choice_table.selectRow(0)
        self.dialog.choice_preview.cellClicked.emit(1, 0)
        self.assertEqual(self.dialog.choice_boxes[0].currentData(),
                         self.dialog.current_issues({"MULTIPLE_BILLS"})[0].candidates[1].candidate_id)
        self.assertTrue(self.dialog.choice_apply.isEnabled())

    def test_conflict_table_is_the_main_editing_area(self):
        conflicts = [
            FieldConflict("cargo-demo", 2, "Loại hàng", "Vôi rời", "Vôi", "Tháng 6"),
            FieldConflict("carrier-demo", 2, "VT biển", "Vietsun", "VIETSUN LINES", "Tháng 6"),
            FieldConflict("seal-demo", 2, "Số chì/Seal", "123", "456", "Tháng 6"),
            FieldConflict("vessel-demo", 2, "Tên tàu", "Tàu A", "Tàu B", "Tháng 6"),
            FieldConflict("port-demo", 2, "Nơi đóng", "Cảng A", "Cảng B", "Tháng 6"),
            FieldConflict("weight-demo", 2, "Số tấn", 28, 30, "Tháng 6"),
        ]
        self.dialog.analysis.issues.append(ImportIssue(
            issue_id="value-demo", kind=ISSUE_FIELD_CONFLICT, title="", reason="",
            doc_type="BILL", container="MMNN1234567", match_date=None,
            reference="Bill BL-A", json_indexes=[1], conflicts=conflicts,
        ))
        self.dialog.refresh_ui(capture=False)
        self.dialog.tabs.setCurrentIndex(2)
        self.dialog.show()
        self.app.processEvents()
        self.app.processEvents()
        self.assertTrue(self.dialog.diff_table.isVisible())
        self.assertEqual(self.dialog.diff_table.rowCount(), 6)
        self.assertLessEqual(self.dialog.diff_table.rowHeight(0), 44)
        self.assertGreaterEqual(self.dialog.diff_table.viewport().height(), 6 * self.dialog.diff_table.rowHeight(0))
        self.assertEqual(self.dialog.diff_table.horizontalHeaderItem(1).text(), "SQT")
        self.assertFalse(self.dialog.diff_table.item(0, 0).font().bold())
        self.assertFalse(self.dialog.diff_table.item(0, 1).font().bold())
        self.assertTrue(self.dialog.diff_preview.isVisible())
        self.assertEqual(self.dialog.diff_preview.item(0, 0).text(), "250")
        self.assertIn("Loại hàng", self.dialog.diff_detail.text())
        self.dialog.diff_table.selectRow(1)
        self.assertIn("VT biển", self.dialog.diff_detail.text())
        self.dialog.resize(1100, 710)
        self.app.processEvents()
        self.assertGreater(self.dialog.diff_table.height(), self.dialog.diff_detail_panel.height())
        self.assertEqual(self.dialog.diff_table.horizontalScrollBar().maximum(), 0)
        self.dialog.diff_boxes[0].setCurrentIndex(self.dialog.diff_boxes[0].findData("NEW"))
        self.assertEqual(self.dialog.diff_boxes[0].currentData(), "NEW")
        self.dialog.diff_boxes[1].setCurrentIndex(self.dialog.diff_boxes[1].findData("CUSTOM"))
        self.assertEqual(self.dialog.diff_boxes[1].currentData(), "CUSTOM")
        self.dialog.diff_inputs[1].setText("Viet Sun")
        self.assertEqual(self.dialog.diff_inputs[1].text(), "Viet Sun")

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

    def test_deferred_document_stays_visible_and_can_be_undone(self):
        before = self.daily.read_bytes()
        self.dialog.choice_table.selectRow(0)
        self.dialog.defer_selected("choice")

        self.assertEqual(self.dialog.choice_table.rowCount(), 1)
        self.assertEqual(self.dialog.choice_table.item(0, 5).text(), "Xử lý sau")
        self.assertEqual(
            self.dialog.choice_table.item(0, 0).background().color().name(), "#fff1b8"
        )
        self.assertEqual(set(self.dialog.analysis.skipped_json_indexes), {0, 1, 2})
        self.assertEqual(self.daily.read_bytes(), before)

        self.dialog.choice_table.selectRow(0)
        self.dialog.undo_defer("choice")
        self.assertEqual(self.dialog.analysis.skipped_count, 0)
        self.assertEqual(self.dialog.choice_table.item(0, 5).text(), "Chưa chọn")

    def test_saved_defer_reopens_with_yellow_row(self):
        self.dialog.choice_table.selectRow(0)
        self.dialog.defer_selected("choice")
        reopened_analysis = self.service.analyze(str(self.output), str(self.daily))
        reopened = GuidedImportResolutionDialog(self.service, reopened_analysis)
        try:
            self.assertEqual(reopened.choice_table.item(0, 5).text(), "Xử lý sau")
            self.assertEqual(
                reopened.choice_table.item(0, 0).background().color().name(), "#fff1b8"
            )
        finally:
            reopened.accept()

    def test_close_after_tab_save_keeps_draft_without_importing(self):
        before = self.daily.read_bytes()
        self.dialog.choice_table.selectRow(0)
        self.dialog.defer_selected("choice")
        self.dialog.reject()
        self.assertEqual(self.dialog.result(), QDialog.Accepted)
        self.assertTrue(self.dialog.saved_only)
        self.assertEqual(self.daily.read_bytes(), before)

    def test_confirm_import_keeps_only_selected_documents_in_waiting_batch(self):
        before = self.daily.read_bytes()
        self.dialog.choice_table.selectRow(0)
        self.dialog.defer_selected("choice")
        self.dialog.edit_table.selectRow(0)
        self.dialog.defer_selected("edit")
        self.assertEqual(self.dialog.analysis.skipped_count, 4)

        self.dialog.finish()
        self.assertEqual(self.dialog.result(), QDialog.Accepted)
        summary = self.service.commit(self.dialog.final_analysis)
        self.assertGreaterEqual(summary.pending_count, 4)
        self.assertEqual(self.daily.read_bytes(), before)
        batches = self.db.list_deferred_batches()
        self.assertEqual(len(batches), 1)
        waiting = self.db.get_deferred_batch(batches[0]["id"])
        self.assertEqual(len(waiting["payload"]["du_lieu_boc_tach"]), 4)
        self.assertEqual(waiting["resolutions"], {})

    def test_multiple_selected_documents_are_deferred_together(self):
        payload = json.loads(self.output.read_text(encoding="utf-8"))
        extra = dict(payload["du_lieu_boc_tach"][3])
        extra.update({"ma_md5_file": "cost-extra", "so_container": "JJKK1234567"})
        payload["du_lieu_boc_tach"].append(extra)
        self.output.write_text(json.dumps(payload, ensure_ascii=False), encoding="utf-8")
        analysis = self.service.analyze(str(self.output), str(self.daily))
        dialog = GuidedImportResolutionDialog(self.service, analysis)
        try:
            self.assertEqual(dialog.edit_table.rowCount(), 2)
            dialog.edit_table.selectRow(0)
            dialog.edit_table.selectionModel().select(
                dialog.edit_table.model().index(1, 0),
                QItemSelectionModel.Select | QItemSelectionModel.Rows,
            )
            dialog.defer_selected("edit")
            self.assertEqual(set(dialog.analysis.skipped_json_indexes), {3, 4})
            self.assertEqual(
                [dialog.edit_table.item(row, 5).text() for row in range(2)],
                ["Xử lý sau", "Xử lý sau"],
            )
        finally:
            dialog.accept()

    def test_bill_label_and_review_show_exact_update(self):
        issue = ImportIssue(
            issue_id="bill-label", kind="FIELD_CONFLICT", title="", reason="",
            doc_type="BILL", container="IIJJ1234567", match_date=None,
            reference="Bill BL-NEW",
        )
        self.assertEqual(self.dialog.source_label(issue), "Bill BL-NEW")
        conflict = FieldConflict(
            "vessel-conflict", 2, "Tên tàu", "Tàu cũ", "Tàu mới", "Tháng 6",
        )
        self.dialog.sandbox.workbook_rows[("Tháng 6", 2)]["Tên tàu"] = "Tàu cũ"
        change = InfoChange(
            action="UPDATE", target_row=2, sqt=250,
            values={"SQT PM": 250, "Số Container": "IIJJ1234567",
                    "Số B/L": "BL-NEW", "Tên tàu": "Tàu mới"},
            staged_ids=[], document_md5s=[], conflicts=[conflict], target_sheet="Tháng 6",
        )
        self.dialog.sandbox.analysis.info_changes = [change]
        self.dialog.sandbox.analysis.conflict_decisions[conflict.conflict_id] = False
        self.dialog.refresh_review()
        self.assertEqual(self.dialog.review_table.item(0, 1).text(), "Cập nhật dòng")
        self.assertEqual(self.dialog.review_table.item(0, 2).text(), "250")
        self.assertIn("Số B/L", self.dialog.review_details.toPlainText())
        self.assertNotIn("Tàu cũ → Tàu mới", self.dialog.review_details.toPlainText())
        self.dialog.sandbox.analysis.conflict_decisions[conflict.conflict_id] = True
        self.dialog.refresh_review()
        self.assertIn("Tàu cũ → Tàu mới", self.dialog.review_details.toPlainText())


if __name__ == "__main__":
    unittest.main()
