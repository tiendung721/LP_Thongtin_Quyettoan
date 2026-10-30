from __future__ import annotations

import os
import unittest
from types import SimpleNamespace
from unittest.mock import Mock, patch

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

try:
    from app.main_window import MainWindow
except ModuleNotFoundError:
    raise unittest.SkipTest("GUI dependencies are not installed")


class CreateNewSqtTests(unittest.TestCase):
    def setUp(self) -> None:
        self.window = SimpleNamespace(
            config=SimpleNamespace(pad_create_new_bat_path=r"C:\project\create.bat"),
            logger=Mock(),
            show_error=Mock(),
            _mark_rpa_launched=Mock(),
        )

    @patch("app.main_window.launch_bat")
    @patch("app.main_window.QInputDialog.getInt", return_value=(5, True))
    def test_selected_count_is_forwarded_to_bat(self, _get_int, launch_bat) -> None:
        MainWindow.on_create_new_sqt(self.window)

        self.assertEqual(launch_bat.call_args.kwargs["arguments"], ("5",))
        self.window._mark_rpa_launched.assert_called_once_with(
            "Đã khởi chạy RPA tạo 5 Số QT mới"
        )

    @patch("app.main_window.launch_bat")
    @patch("app.main_window.QInputDialog.getInt", return_value=(5, False))
    def test_cancel_does_not_launch_bat(self, _get_int, launch_bat) -> None:
        MainWindow.on_create_new_sqt(self.window)

        launch_bat.assert_not_called()
