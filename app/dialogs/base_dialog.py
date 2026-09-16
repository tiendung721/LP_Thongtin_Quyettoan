"""Các thiết lập cửa sổ dùng chung cho giao diện nghiệp vụ."""

from __future__ import annotations

from PySide6.QtCore import Qt
from PySide6.QtWidgets import QDialog, QWidget


def enable_standard_window_controls(window: QWidget) -> None:
    """Bật đầy đủ nút hệ thống mà không làm mất các cờ cửa sổ hiện có."""
    window.setWindowFlag(Qt.WindowMinimizeButtonHint, True)
    window.setWindowFlag(Qt.WindowMaximizeButtonHint, True)
    window.setWindowFlag(Qt.WindowCloseButtonHint, True)
    window.setWindowFlag(Qt.WindowContextHelpButtonHint, False)


class ResizableDialog(QDialog):
    """QDialog có nút thu nhỏ/phóng to như một cửa sổ Windows thông thường."""

    def __init__(self, parent=None):
        super().__init__(parent)
        enable_standard_window_controls(self)

