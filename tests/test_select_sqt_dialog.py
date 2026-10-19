from __future__ import annotations

import os
import unittest
from unittest.mock import patch

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

try:
    from PySide6.QtCore import Qt  # noqa: E402
    from PySide6.QtWidgets import QApplication, QDialog, QLabel  # noqa: E402
except ModuleNotFoundError:  # pragma: no cover - depends on local GUI deps
    raise unittest.SkipTest("PySide6 is not installed in this test environment")

from app.dialogs.select_sqt_dialog import SelectSqtDialog  # noqa: E402
from app.services.sqt_selection_service import (  # noqa: E402
    EXPENSE_SELECTION_OPERATION,
    SqtSelectionItem,
)


def _app() -> QApplication:
    return QApplication.instance() or QApplication([])


def _check(dialog: SelectSqtDialog, sqt: str) -> None:
    for row in range(dialog.table.rowCount()):
        if dialog.table.item(row, 1).text() == sqt:
            dialog.table.item(row, 0).setCheckState(Qt.Checked)
            return
    raise AssertionError(f"Không tìm thấy SQT {sqt}")


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
        _check(dialog, "598")
        _check(dialog, "596")

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
        _check(dialog, "597")

        dialog._on_run()

        self.assertEqual(dialog.result(), QDialog.Accepted)
        self.assertEqual(dialog.selected_values(), ["597"])

    def test_status_filter_preserves_selection_and_summary(self) -> None:
        items = [
            SqtSelectionItem("700", 2),
            SqtSelectionItem("701", 1, imported_row_count=1),
            SqtSelectionItem("702", 3, imported_row_count=2),
        ]
        dialog = SelectSqtDialog(items)

        _check(dialog, "700")
        dialog.status_combo.setCurrentIndex(3)
        self.assertEqual(sum(not dialog.table.isRowHidden(row) for row in range(3)), 1)
        _check(dialog, "701")

        self.assertEqual(dialog.selected_values(), ["700", "701"])
        self.assertIn("1 cần nhập", dialog.top_summary.text())
        self.assertIn("1 nhập lại", dialog.top_summary.text())

    def test_switching_month_clears_selection_from_previous_month(self) -> None:
        items = [
            SqtSelectionItem("700", 1, "Tháng 7", (2,)),
            SqtSelectionItem("600", 1, "Tháng 6", (2,)),
        ]
        dialog = SelectSqtDialog(items)
        _check(dialog, "700")

        dialog.month_combo.setCurrentIndex(1)
        _check(dialog, "600")

        self.assertEqual(dialog.selected_values(), ["600"])
        self.assertEqual(dialog.selected_items()[0].sheet_name, "Tháng 6")

    def test_imported_selection_requires_confirmation(self) -> None:
        item = SqtSelectionItem("701", 1, imported_row_count=1)
        dialog = SelectSqtDialog([item])
        _check(dialog, "701")

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
        self.assertEqual(dialog.table.columnCount(), 7)

    def test_searches_container_and_shows_selected_detail(self) -> None:
        from app.services.sqt_selection_service import SqtSourceRow

        items = [
            SqtSelectionItem("700", 1, source_rows=(SqtSourceRow(2, "ABCD1234567", "2026-10-12"),)),
            SqtSelectionItem("701", 1, source_rows=(SqtSourceRow(3, "WXYZ1234567", "2026-10-13"),)),
        ]
        dialog = SelectSqtDialog(items)
        dialog.search_input.setText("wxyz")

        self.assertEqual(sum(not dialog.table.isRowHidden(row) for row in range(2)), 1)
        self.assertEqual(dialog.detail_title.text(), "SQT 701")
        self.assertIn("13/10/2026", dialog.detail_meta.text())

    def test_expense_total_and_optional_material_price_appear_in_detail(self) -> None:
        from app.services.sqt_selection_service import SqtSourceRow

        item = SqtSelectionItem(
            "722", 2, source_rows=(
                SqtSourceRow(2, "ABCD1234567", "2026-10-12", total_amount=2_000_000, material_price=1_200_000),
                SqtSourceRow(3, "EFGH1234567", "2026-10-13", total_amount=3_000_000),
            ),
        )
        dialog = SelectSqtDialog([item], operation=EXPENSE_SELECTION_OPERATION)
        texts = [child.text() for child in dialog.detail_host.findChildren(QLabel)]

        self.assertEqual(dialog.table.item(0, 6).text(), "5.000.000 đ")
        self.assertEqual(dialog.table.item(0, 4).text(), "12–13/10/2026")
        self.assertIn("Giá VL: 1.200.000 đ", texts)
        self.assertEqual(sum(text.startswith("Giá VL:") for text in texts), 1)


if __name__ == "__main__":
    unittest.main()
