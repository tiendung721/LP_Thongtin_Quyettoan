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
    from PySide6.QtWidgets import QApplication
except ModuleNotFoundError:  # pragma: no cover - depends on local GUI deps
    raise unittest.SkipTest("PySide6 is not installed in this test environment")

from app.daily_import import (
    ISSUE_NO_TARGET,
    RESOLUTION_SKIP,
    DailyImportService,
    ImportAnalysis,
    ImportIssue,
    IssueResolution,
)
from app.daily_import_ui import ImportResolutionDialog
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
        self.assertIn("1  bỏ qua", dialog.lbl_skipped.text())

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


if __name__ == "__main__":
    unittest.main()
