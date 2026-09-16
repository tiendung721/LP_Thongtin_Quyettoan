from __future__ import annotations

import os
import unittest
from unittest.mock import patch

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

try:
    from PySide6.QtCore import Qt
    from PySide6.QtWidgets import QApplication, QLabel, QPushButton
except ModuleNotFoundError:  # pragma: no cover - depends on local GUI deps
    raise unittest.SkipTest("PySide6 is not installed in this test environment")

from app.dialogs.base_dialog import ResizableDialog
from app.dialogs.notifications import (
    NotificationAction,
    NotificationDialog,
    NotificationMetric,
    notify,
)


def _app() -> QApplication:
    return QApplication.instance() or QApplication([])


class WindowControlsTests(unittest.TestCase):
    def setUp(self) -> None:
        self.app = _app()

    def test_resizable_dialog_has_standard_minimize_and_maximize_buttons(self) -> None:
        dialog = ResizableDialog()
        flags = dialog.windowFlags()

        self.assertTrue(flags & Qt.WindowMinimizeButtonHint)
        self.assertTrue(flags & Qt.WindowMaximizeButtonHint)
        self.assertTrue(flags & Qt.WindowCloseButtonHint)
        self.assertFalse(flags & Qt.WindowContextHelpButtonHint)


class NotificationDialogTests(unittest.TestCase):
    def setUp(self) -> None:
        self.app = _app()

    def test_result_dialog_builds_metrics_and_uses_action_key(self) -> None:
        dialog = NotificationDialog(
            None,
            "Hoàn tất",
            "Đã cập nhật file.",
            tone="success",
            metrics=(NotificationMetric("12", "Dòng mới", "success"),),
            actions=(NotificationAction("action", "Mở file", "primary"),),
        )

        self.assertEqual(dialog.selected_action, "close")
        self.assertIn("Hoàn tất", dialog.windowTitle())
        metric_value = dialog.findChild(QLabel, "notificationMetricValue")
        self.assertIsNotNone(metric_value)
        self.assertEqual(metric_value.text(), "12")
        self.assertTrue(
            any(button.text() == "Mở file" for button in dialog.findChildren(QPushButton))
        )

        dialog._choose("action")
        self.assertEqual(dialog.selected_action, "action")

    def test_confirm_returns_false_when_dialog_is_closed(self) -> None:
        with patch.object(NotificationDialog, "exec", return_value=0):
            self.assertFalse(notify.confirm(None, "Xác nhận", "Tiếp tục?"))


if __name__ == "__main__":
    unittest.main()
