from __future__ import annotations

import os
import unittest
from unittest.mock import patch

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

try:
    from PySide6.QtWidgets import QApplication, QDialog  # noqa: E402
except ModuleNotFoundError:  # pragma: no cover - depends on local GUI deps
    raise unittest.SkipTest("PySide6 is not installed in this test environment")

from app.dialogs.select_sqt_dialog import SelectSqtDialog  # noqa: E402
from app.services.sqt_selection_service import (  # noqa: E402
    EXPENSE_SELECTION_OPERATION,
    SqtSelectionItem,
)


def _app() -> QApplication:
    return QApplication.instance() or QApplication([])


class SelectSqtDialogTests(unittest.TestCase):
    def setUp(self) -> None:
        self.app = _app()
        self.items = [
            SqtSelectionItem("596", 2),
            SqtSelectionItem("597", 1),
            SqtSelectionItem("598", 5),
        ]

    def test_select_one_or_many_preserves_original_order(self) -> None:
        dialog = SelectSqtDialog(self.items)
        dialog.checkboxes[2].setChecked(True)
        dialog.checkboxes[0].setChecked(True)

        self.assertEqual(dialog.selected_values(), ["596", "598"])

    def test_select_all_and_clear_all(self) -> None:
        dialog = SelectSqtDialog(self.items)
        dialog.select_all()
        self.assertEqual(dialog.selected_values(), ["596", "597", "598"])

        dialog.clear_all()
        self.assertEqual(dialog.selected_values(), [])

    def test_run_without_selection_warns_and_keeps_dialog_open(self) -> None:
        dialog = SelectSqtDialog(self.items)

        with patch("app.dialogs.select_sqt_dialog.notify.warning") as warning:
            dialog._on_run()

        warning.assert_called_once()
        self.assertEqual(dialog.result(), QDialog.Rejected)

    def test_run_with_selection_accepts(self) -> None:
        dialog = SelectSqtDialog(self.items)
        dialog.checkboxes[1].setChecked(True)

        dialog._on_run()

        self.assertEqual(dialog.result(), QDialog.Accepted)
        self.assertEqual(dialog.selected_values(), ["597"])

    def test_pending_and_imported_are_separate_tabs_with_shared_selection(self) -> None:
        items = [
            SqtSelectionItem("700", 2),
            SqtSelectionItem("701", 1, imported_row_count=1),
            SqtSelectionItem("702", 3, imported_row_count=2),
        ]
        dialog = SelectSqtDialog(items)

        self.assertEqual(dialog.tabs.tabText(0), "CẦN NHẬP (2)")
        self.assertEqual(dialog.tabs.tabText(1), "ĐÃ NHẬP (1)")

        dialog.checkboxes[0].setChecked(True)
        dialog.tabs.setCurrentIndex(1)
        dialog.checkboxes[1].setChecked(True)

        self.assertEqual(dialog.selected_values(), ["700", "701"])
        self.assertIn("1 cần nhập", dialog.selection_summary.text())
        self.assertIn("1 nhập lại", dialog.selection_summary.text())

    def test_switching_month_clears_selection_from_previous_month(self) -> None:
        items = [
            SqtSelectionItem("700", 1, "Tháng 7", (2,)),
            SqtSelectionItem("600", 1, "Tháng 6", (2,)),
        ]
        dialog = SelectSqtDialog(items)
        dialog.checkboxes[0].setChecked(True)

        dialog.month_combo.setCurrentIndex(1)
        dialog.checkboxes[1].setChecked(True)

        self.assertFalse(dialog.checkboxes[0].isChecked())
        self.assertEqual(dialog.selected_values(), ["600"])
        self.assertEqual(dialog.selected_items()[0].sheet_name, "Tháng 6")

    def test_imported_selection_requires_confirmation(self) -> None:
        item = SqtSelectionItem("701", 1, imported_row_count=1)
        dialog = SelectSqtDialog([item])
        dialog.checkboxes[0].setChecked(True)

        with patch(
            "app.dialogs.select_sqt_dialog.notify.confirm", return_value=False
        ) as question:
            dialog._on_run()
        question.assert_called_once()
        self.assertEqual(dialog.result(), QDialog.Rejected)

        with patch("app.dialogs.select_sqt_dialog.notify.confirm", return_value=True):
            dialog._on_run()
        self.assertEqual(dialog.result(), QDialog.Accepted)

    def test_expense_dialog_uses_expense_labels(self) -> None:
        dialog = SelectSqtDialog(
            [SqtSelectionItem("700", 2)],
            operation=EXPENSE_SELECTION_OPERATION,
        )

        self.assertIn("khoản chi", dialog.windowTitle())
        self.assertEqual(dialog.btn_run.text(), "Chạy nhập khoản chi")
        self.assertFalse(dialog.month_combo.isVisible())


if __name__ == "__main__":
    unittest.main()
