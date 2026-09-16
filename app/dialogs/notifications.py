"""Hộp thông báo thống nhất, dễ đọc và dùng hoàn toàn bằng tiếng Việt."""

from __future__ import annotations

from dataclasses import dataclass
from typing import Iterable, Optional, Sequence

from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QApplication,
    QDialog,
    QFrame,
    QHBoxLayout,
    QLabel,
    QPlainTextEdit,
    QPushButton,
    QStyle,
    QToolButton,
    QVBoxLayout,
    QWidget,
)


@dataclass(frozen=True)
class NotificationMetric:
    value: str
    label: str
    tone: str = "info"


@dataclass(frozen=True)
class NotificationAction:
    key: str
    text: str
    variant: str = ""


class NotificationDialog(QDialog):
    """Thông báo modal có phân cấp nội dung, chi tiết và hành động rõ ràng."""

    _ICONS = {
        "success": QStyle.SP_DialogApplyButton,
        "info": QStyle.SP_MessageBoxInformation,
        "warning": QStyle.SP_MessageBoxWarning,
        "error": QStyle.SP_MessageBoxCritical,
        "question": QStyle.SP_MessageBoxQuestion,
    }

    def __init__(
        self,
        parent: Optional[QWidget],
        title: str,
        message: str,
        *,
        tone: str = "info",
        details: str = "",
        context: str = "",
        notice: str = "",
        metrics: Sequence[NotificationMetric] = (),
        actions: Sequence[NotificationAction] = (),
        cancel_key: str = "close",
        default_key: str = "",
    ):
        super().__init__(parent)
        self.selected_action = cancel_key
        self.setObjectName("notificationDialog")
        self.setProperty("tone", tone)
        self.setWindowTitle(title)
        self.setWindowFlag(Qt.WindowContextHelpButtonHint, False)
        self.setModal(True)
        self.setMinimumWidth(520)
        self.setMaximumWidth(760)
        self.resize(590, self.sizeHint().height())

        root = QVBoxLayout(self)
        root.setContentsMargins(0, 0, 0, 0)
        root.setSpacing(0)

        accent = QFrame()
        accent.setObjectName("notificationAccent")
        accent.setProperty("tone", tone)
        accent.setFixedHeight(5)
        root.addWidget(accent)

        content = QVBoxLayout()
        content.setContentsMargins(24, 22, 24, 20)
        content.setSpacing(16)
        root.addLayout(content)

        header = QHBoxLayout()
        header.setSpacing(14)
        icon = QLabel()
        icon.setObjectName("notificationIcon")
        icon.setAlignment(Qt.AlignTop | Qt.AlignHCenter)
        standard_icon = self._ICONS.get(tone, self._ICONS["info"])
        pixmap = self.style().standardIcon(standard_icon).pixmap(38, 38)
        icon.setPixmap(pixmap)
        icon.setFixedSize(42, 42)
        header.addWidget(icon)

        text_column = QVBoxLayout()
        text_column.setSpacing(7)
        heading = QLabel(title)
        heading.setObjectName("notificationTitle")
        heading.setWordWrap(True)
        body = QLabel(message)
        body.setObjectName("notificationMessage")
        body.setWordWrap(True)
        body.setTextInteractionFlags(Qt.TextSelectableByMouse)
        text_column.addWidget(heading)
        text_column.addWidget(body)
        header.addLayout(text_column, stretch=1)
        content.addLayout(header)

        if metrics:
            content.addLayout(self._build_metrics(metrics))

        if notice:
            notice_box = QFrame()
            notice_box.setObjectName("notificationNotice")
            notice_layout = QHBoxLayout(notice_box)
            notice_layout.setContentsMargins(12, 10, 12, 10)
            notice_label = QLabel(notice)
            notice_label.setWordWrap(True)
            notice_label.setObjectName("notificationNoticeText")
            notice_layout.addWidget(notice_label)
            content.addWidget(notice_box)

        if context:
            context_label = QLabel(context)
            context_label.setObjectName("notificationContext")
            context_label.setWordWrap(True)
            content.addWidget(context_label)

        if details:
            content.addWidget(self._build_details(details))

        action_row = QHBoxLayout()
        action_row.addStretch(1)
        for action in actions or (NotificationAction("close", "Đóng"),):
            button = QPushButton(action.text)
            if action.variant:
                button.setProperty("variant", action.variant)
            button.clicked.connect(
                lambda _checked=False, key=action.key: self._choose(key)
            )
            if action.key == default_key:
                button.setDefault(True)
                button.setFocus()
            action_row.addWidget(button)
        content.addLayout(action_row)

    @staticmethod
    def _build_metrics(metrics: Sequence[NotificationMetric]) -> QHBoxLayout:
        row = QHBoxLayout()
        row.setSpacing(8)
        for metric in metrics:
            card = QFrame()
            card.setObjectName("notificationMetric")
            card.setProperty("tone", metric.tone)
            layout = QVBoxLayout(card)
            layout.setContentsMargins(9, 9, 9, 9)
            layout.setSpacing(2)
            value = QLabel(metric.value)
            value.setObjectName("notificationMetricValue")
            value.setProperty("tone", metric.tone)
            value.setAlignment(Qt.AlignCenter)
            label = QLabel(metric.label)
            label.setObjectName("notificationMetricLabel")
            label.setAlignment(Qt.AlignCenter)
            label.setWordWrap(True)
            layout.addWidget(value)
            layout.addWidget(label)
            row.addWidget(card, stretch=1)
        return row

    def _build_details(self, details: str) -> QWidget:
        host = QWidget()
        layout = QVBoxLayout(host)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(7)

        toggle = QToolButton()
        toggle.setObjectName("notificationDetailsToggle")
        toggle.setText("Chi tiết kỹ thuật")
        toggle.setCheckable(True)
        toggle.setArrowType(Qt.RightArrow)
        toggle.setToolButtonStyle(Qt.ToolButtonTextBesideIcon)
        layout.addWidget(toggle)

        panel = QFrame()
        panel.setObjectName("notificationDetails")
        panel_layout = QVBoxLayout(panel)
        panel_layout.setContentsMargins(10, 10, 10, 10)
        detail_text = QPlainTextEdit(details)
        detail_text.setReadOnly(True)
        detail_text.setMinimumHeight(92)
        detail_text.setMaximumHeight(170)
        copy_button = QPushButton("Sao chép chi tiết")
        copy_button.clicked.connect(
            lambda: QApplication.clipboard().setText(details)
        )
        copy_row = QHBoxLayout()
        copy_row.addStretch(1)
        copy_row.addWidget(copy_button)
        panel_layout.addWidget(detail_text)
        panel_layout.addLayout(copy_row)
        panel.setVisible(False)
        layout.addWidget(panel)

        def toggle_details(checked: bool) -> None:
            panel.setVisible(checked)
            toggle.setArrowType(Qt.DownArrow if checked else Qt.RightArrow)
            self.adjustSize()

        toggle.toggled.connect(toggle_details)
        return host

    def _choose(self, key: str) -> None:
        self.selected_action = key
        self.accept()


class NotificationService:
    """API duy nhất để toàn ứng dụng hiển thị thông báo."""

    @staticmethod
    def _show(
        parent: Optional[QWidget],
        title: str,
        message: str,
        *,
        tone: str,
        details: str = "",
        context: str = "",
        notice: str = "",
        metrics: Sequence[NotificationMetric] = (),
        actions: Sequence[NotificationAction] = (),
        cancel_key: str = "close",
        default_key: str = "",
    ) -> str:
        dialog = NotificationDialog(
            parent,
            title,
            message,
            tone=tone,
            details=details,
            context=context,
            notice=notice,
            metrics=metrics,
            actions=actions,
            cancel_key=cancel_key,
            default_key=default_key,
        )
        dialog.exec()
        return dialog.selected_action

    def info(self, parent: Optional[QWidget], title: str, message: str) -> None:
        self._show(parent, title, message, tone="info")

    def success(self, parent: Optional[QWidget], title: str, message: str) -> None:
        self._show(parent, title, message, tone="success")

    def warning(self, parent: Optional[QWidget], title: str, message: str) -> None:
        self._show(parent, title, message, tone="warning")

    def error(
        self,
        parent: Optional[QWidget],
        title: str,
        message: str,
        *,
        details: str = "",
    ) -> None:
        self._show(parent, title, message, tone="error", details=details)

    def confirm(
        self,
        parent: Optional[QWidget],
        title: str,
        message: str,
        *,
        confirm_text: str = "Tiếp tục",
        cancel_text: str = "Quay lại",
        tone: str = "question",
        default_confirm: bool = False,
        confirm_variant: str = "primary",
    ) -> bool:
        actions = [NotificationAction("cancel", cancel_text)]
        actions.append(NotificationAction("confirm", confirm_text, confirm_variant))
        result = self._show(
            parent,
            title,
            message,
            tone=tone,
            actions=actions,
            cancel_key="cancel",
            default_key="confirm" if default_confirm else "cancel",
        )
        return result == "confirm"

    def result(
        self,
        parent: Optional[QWidget],
        title: str,
        message: str,
        *,
        metrics: Iterable[NotificationMetric],
        notice: str = "",
        context: str = "",
        action_text: str = "",
    ) -> bool:
        actions = [NotificationAction("close", "Đóng")]
        if action_text:
            actions.append(NotificationAction("action", action_text, "primary"))
        result = self._show(
            parent,
            title,
            message,
            tone="success",
            notice=notice,
            context=context,
            metrics=tuple(metrics),
            actions=actions,
        )
        return result == "action"


notify = NotificationService()
