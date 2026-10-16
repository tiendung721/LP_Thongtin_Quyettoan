from __future__ import annotations

import json
import logging
import os
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

try:
    from PySide6.QtWidgets import QApplication, QDialog
except ModuleNotFoundError:  # pragma: no cover - depends on local GUI deps
    raise unittest.SkipTest("PySide6 is not installed in this test environment")

from app.daily_import import (
    FieldConflict,
    ISSUE_FIELD_CONFLICT,
    ISSUE_MULTIPLE_TARGETS,
    ImportCandidate,
    ISSUE_NO_TARGET,
    RESOLUTION_SKIP,
    DailyImportService,
    ImportAnalysis,
    ImportIssue,
    IssueResolution,
)
from app.daily_import_ui import BulkCandidateDialog, BulkFieldDialog, ImportResolutionDialog
from app.database import Database


def _app() -> QApplication:
    return QApplication.instance() or QApplication([])


class ImportResolutionDialogTests(unittest.TestCase):
    def setUp(self) -> None:
        self.app = _app()
        self.temp = tempfile.TemporaryDirectory()
        self.root = Path(self.temp.name)
        self.output = self.root / "output.json"
        self.output.write_text(
            json.dumps(
                {
                    "metadata": {},
                    "du_lieu_boc_tach": [
                        {
                            "loai_chung_tu": "Khoản chi",
                            "so_container": "TEST1234567",
                            "file_nguon": "cost.jpg",
                        }
                    ],
                },
                ensure_ascii=False,
            ),
            encoding="utf-8",
        )
        database = Database(str(self.root / "state.db"))
        database.init_db()
        self.service = DailyImportService(database, logging.getLogger("dialog-test"))
        self.issue = ImportIssue(
            issue_id="issue-1",
            kind=ISSUE_NO_TARGET,
            title="Không tìm thấy SQT",
            reason="Không tìm thấy dòng quyết toán tương ứng",
            doc_type="KHOAN_CHI",
            container="TEST1234567",
            match_date="2026-09-01",
            reference="HĐ 01",
            json_indexes=[0],
            staged_ids=[1],
        )

    def tearDown(self) -> None:
        self.temp.cleanup()

    def _analysis(self, *, resolved: bool = False) -> ImportAnalysis:
        resolutions = {}
        if resolved:
            resolutions[self.issue.issue_id] = IssueResolution(
                self.issue.issue_id, RESOLUTION_SKIP
            )
        return ImportAnalysis(
            output_path=str(self.output),
            daily_path=str(self.root / "daily.xlsx"),
            processed_file_id=1,
            issues=[self.issue],
            resolutions=resolutions,
        )

    def test_continue_is_disabled_until_every_issue_has_a_decision(self) -> None:
        dialog = ImportResolutionDialog(self.service, self._analysis())
        self.assertFalse(dialog.btn_continue.isEnabled())
        self.assertIn("1  chưa xử lý", dialog.lbl_pending.text())

        dialog.resolutions[self.issue.issue_id] = IssueResolution(
            self.issue.issue_id, RESOLUTION_SKIP
        )
        dialog._refresh_all()

        self.assertTrue(dialog.btn_continue.isEnabled())
        self.assertIn("1  để sau", dialog.lbl_skipped.text())

    def test_first_continue_opens_summary_in_same_dialog(self) -> None:
        dialog = ImportResolutionDialog(self.service, self._analysis(resolved=True))

        dialog._continue()

        self.assertTrue(dialog._reviewing)
        self.assertEqual(dialog.btn_continue.text(), "Xác nhận nhập")


    def test_rebuilding_detail_detaches_old_widgets_immediately(self) -> None:
        dialog = ImportResolutionDialog(self.service, self._analysis())
        old_widgets = [
            dialog.detail_layout.itemAt(index).widget()
            for index in range(dialog.detail_layout.count())
            if dialog.detail_layout.itemAt(index).widget() is not None
        ]

        dialog._show_issue(self.issue)

        self.assertTrue(old_widgets)
        self.assertTrue(all(widget.parent() is None for widget in old_widgets))

    def test_refresh_table_renders_selected_issue_only_once(self) -> None:
        dialog = ImportResolutionDialog(self.service, self._analysis())
        dialog.issue_table.clearSelection()

        with patch.object(dialog, "_show_issue", wraps=dialog._show_issue) as show:
            dialog._refresh_issue_table()

        self.assertEqual(show.call_count, 1)

    def test_bulk_menu_splits_mixed_selection_and_keeps_other_rows(self) -> None:
        conflict = ImportIssue(
            issue_id="issue-2", kind=ISSUE_FIELD_CONFLICT,
            title="Dữ liệu khác nhau", reason="Đơn giá khác",
            doc_type="KHOAN_CHI", container="OTHER1234567",
            match_date="2026-09-01", reference="HĐ 02",
            json_indexes=[0], staged_ids=[2],
            conflicts=[FieldConflict("field-1", 2, "Đơn giá", 100, 200)],
        )
        analysis = self._analysis()
        analysis.issues.append(conflict)
        dialog = ImportResolutionDialog(self.service, analysis)
        dialog._select_visible_issues()
        self.assertEqual(len(dialog._selected_issue_ids), 2)
        self.assertTrue(dialog.btn_bulk.isEnabled())
        actions = [action for action in dialog._build_bulk_menu().actions() if not action.isSeparator()]
        self.assertEqual(len(actions), 3)
        self.assertIn("1 dòng", actions[0].text())
        self.assertIn("1 dòng", actions[1].text())

        dialog.kind_filter.setCurrentIndex(1)
        self.assertIn("1 dòng đang ẩn", dialog.lbl_selected.text())
        def choose(group_dialog):
            group_dialog.modes = {"Đơn giá": "NEW"}
            return QDialog.Accepted
        with patch.object(BulkFieldDialog, "exec", new=choose), patch.object(dialog, "_queue_reanalysis"):
            actions[1].trigger()
        self.assertIn("issue-2", dialog.resolutions)
        self.assertNotIn("issue-1", dialog.resolutions)
        self.assertIn("issue-1", dialog._selected_issue_ids)
        self.assertNotIn("issue-2", dialog._selected_issue_ids)
        with patch.object(dialog, "_queue_reanalysis"):
            dialog._undo_last_bulk()
        self.assertEqual(dialog.resolutions, {})
        self.assertEqual(len(dialog._selected_issue_ids), 2)

    def test_bulk_candidate_requires_a_choice_shared_by_every_row(self) -> None:
        shared = ImportCandidate("target-1", "INFO_ROW", "Target 1")
        only_second = ImportCandidate("target-2", "INFO_ROW", "Target 2")
        issues = [
            ImportIssue(
                issue_id=f"candidate-{index}", kind=ISSUE_MULTIPLE_TARGETS,
                title="Multiple targets", reason="Choose one",
                doc_type="KHOAN_CHI", container=f"BOX-{index}",
                match_date="2026-09-01", reference=f"HD-{index}",
                json_indexes=[0], candidates=[shared, only_second],
            )
            for index in (1, 2)
        ]
        analysis = self._analysis()
        analysis.issues = issues
        dialog = ImportResolutionDialog(self.service, analysis)
        dialog._select_visible_issues()

        def choose(group_dialog):
            group_dialog.choice = shared.candidate_id
            return QDialog.Accepted
        with patch.object(BulkCandidateDialog, "exec", new=choose), patch.object(dialog, "_queue_reanalysis"):
            dialog._bulk_candidates(tuple(issue.issue_id for issue in issues))
        self.assertEqual({r.candidate_id for r in dialog.resolutions.values()}, {shared.candidate_id})
        self.assertEqual(len(dialog.resolutions), 2)
        dialog.resolutions.clear()
        dialog._select_visible_issues()
        issues[1].candidates = [ImportCandidate("unrelated", "INFO_ROW", "Unrelated")]
        with patch("app.daily_import_ui.notify.info") as message:
            dialog._bulk_candidates(tuple(issue.issue_id for issue in issues))
        self.assertTrue(message.called)
        self.assertEqual(dialog.resolutions, {})


if __name__ == "__main__":
    unittest.main()
