"""Status-aware dialogs for selecting one or more SQT values."""

from __future__ import annotations

from collections import OrderedDict
from typing import Iterable, List

from PySide6.QtCore import Qt, Slot
from PySide6.QtWidgets import (
    QCheckBox,
    QComboBox,
    QFrame,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QPushButton,
    QScrollArea,
    QTabWidget,
    QVBoxLayout,
    QWidget,
)

from .base_dialog import ResizableDialog
from .notifications import notify
from ..services.sqt_selection_service import (
    EXPENSE_SELECTION_OPERATION,
    SELECTION_OPERATION,
    SqtSelectionItem,
)


class _SqtChoiceRow(QFrame):
    """One selectable SQT row without a separate status column."""

    def __init__(self, item: SqtSelectionItem, *, expense: bool):
        super().__init__()
        self.item = item
        self.setObjectName("sqtSelectionRow")

        layout = QVBoxLayout(self)
        layout.setContentsMargins(12, 9, 12, 9)
        layout.setSpacing(3)

        main_line = QHBoxLayout()
        main_line.setSpacing(10)
        self.checkbox = QCheckBox(f"SQT {item.value}")
        self.checkbox.setObjectName("sqtChoice")
        self.checkbox.setProperty("sqtValue", item.value)
        self.checkbox.setToolTip(item.value)
        main_line.addWidget(self.checkbox)
        main_line.addStretch(1)

        unit = "khoản chi" if expense else "dòng"
        count = QLabel(f"{item.row_count} {unit}")
        count.setObjectName("sqtRowCount")
        main_line.addWidget(count)
        layout.addLayout(main_line)

        if item.is_partially_imported:
            detail = QLabel(
                f"Còn {item.pending_row_count}/{item.row_count} dòng cần nhập; "
                "khi chạy sẽ xử lý lại toàn bộ SQT"
            )
            detail.setObjectName("sqtPartialHint")
            detail.setWordWrap(True)
            layout.addWidget(detail)


class SelectSqtDialog(ResizableDialog):
    """Select SQT values from separate pending and imported display flows."""

    def __init__(
        self,
        items: Iterable[SqtSelectionItem],
        parent=None,
        *,
        operation: str = SELECTION_OPERATION,
    ):
        super().__init__(parent)
        self.items = list(items)
        self.operation = operation
        self.is_expense = operation == EXPENSE_SELECTION_OPERATION
        self.checkboxes: List[QCheckBox] = []
        self._rows: list[tuple[SqtSelectionItem, _SqtChoiceRow]] = []
        self._selected_values: List[str] = []
        self._selected_items: List[SqtSelectionItem] = []

        action_name = "nhập khoản chi" if self.is_expense else "nhập thông tin"
        self.setWindowTitle(f"Chọn Số quyết toán {action_name}")
        self.setModal(True)
        self.resize(760, 650)
        self.setMinimumSize(600, 460)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(20, 18, 20, 18)
        layout.setSpacing(12)

        title = QLabel(
            "CHỌN SQT NHẬP KHOẢN CHI"
            if self.is_expense
            else "CHỌN SQT NHẬP THÔNG TIN"
        )
        title.setObjectName("cardTitle")
        title.setAlignment(Qt.AlignCenter)
        layout.addWidget(title)

        filters = QHBoxLayout()
        filters.setSpacing(10)
        self.month_label = QLabel("Tháng:")
        self.month_label.setObjectName("formLabel")
        self.month_combo = QComboBox()
        self.month_combo.setMinimumWidth(210)
        self._populate_months()
        show_month = not self.is_expense and any(item.sheet_name for item in self.items)
        self.month_label.setVisible(show_month)
        self.month_combo.setVisible(show_month)
        filters.addWidget(self.month_label)
        filters.addWidget(self.month_combo)
        filters.addStretch(1)
        filters.addWidget(QLabel("Tìm SQT:"))
        self.search_input = QLineEdit()
        self.search_input.setPlaceholderText("Nhập Số quyết toán...")
        self.search_input.setClearButtonEnabled(True)
        self.search_input.setMinimumWidth(220)
        filters.addWidget(self.search_input)
        layout.addLayout(filters)

        self.tabs = QTabWidget()
        self.tabs.setObjectName("sqtStatusTabs")
        self.pending_page, self.pending_layout = self._create_selection_page()
        imported_intro = (
            "Các SQT bên dưới đã được nhập. Bạn vẫn có thể chọn để nhập lại."
        )
        self.imported_page, self.imported_layout = self._create_selection_page(
            imported_intro
        )
        self.tabs.addTab(self.pending_page, "CẦN NHẬP")
        self.tabs.addTab(self.imported_page, "ĐÃ NHẬP")
        layout.addWidget(self.tabs, stretch=1)

        for item in self.items:
            row = _SqtChoiceRow(item, expense=self.is_expense)
            row.checkbox.stateChanged.connect(self._update_summary)
            self.checkboxes.append(row.checkbox)
            self._rows.append((item, row))
            target_layout = (
                self.imported_layout if item.is_fully_imported else self.pending_layout
            )
            target_layout.addWidget(row)
        self.pending_layout.addStretch(1)
        self.imported_layout.addStretch(1)

        bulk = QHBoxLayout()
        self.btn_select_all = QPushButton("Chọn tất cả trong tab")
        self.btn_clear_all = QPushButton("Bỏ chọn trong tab")
        self.btn_select_all.clicked.connect(self.select_all)
        self.btn_clear_all.clicked.connect(self.clear_all)
        bulk.addWidget(self.btn_select_all)
        bulk.addWidget(self.btn_clear_all)
        bulk.addStretch(1)
        layout.addLayout(bulk)

        actions = QHBoxLayout()
        btn_cancel = QPushButton("Hủy")
        self.selection_summary = QLabel()
        self.selection_summary.setObjectName("sqtSelectionSummary")
        self.btn_run = QPushButton(
            "Chạy nhập khoản chi" if self.is_expense else "Chạy nhập thông tin"
        )
        self.btn_run.setProperty("variant", "primary")
        btn_cancel.clicked.connect(self.reject)
        self.btn_run.clicked.connect(self._on_run)
        actions.addWidget(btn_cancel)
        actions.addWidget(self.selection_summary)
        actions.addStretch(1)
        actions.addWidget(self.btn_run)
        layout.addLayout(actions)

        self.month_combo.currentIndexChanged.connect(self._on_month_changed)
        self.search_input.textChanged.connect(self._apply_filters)
        self.tabs.currentChanged.connect(self._apply_filters)
        self._apply_filters()
        self._update_summary()

    def _populate_months(self) -> None:
        grouped: "OrderedDict[str, int]" = OrderedDict()
        for item in self.items:
            key = item.sheet_name or ""
            grouped[key] = grouped.get(key, 0) + 1
        for sheet_name, count in grouped.items():
            label = f"{sheet_name} — {count} SQT" if sheet_name else f"{count} SQT"
            self.month_combo.addItem(label, sheet_name)

    @staticmethod
    def _create_selection_page(intro: str = "") -> tuple[QWidget, QVBoxLayout]:
        page = QWidget()
        page_layout = QVBoxLayout(page)
        page_layout.setContentsMargins(2, 8, 2, 2)
        page_layout.setSpacing(8)
        if intro:
            hint = QLabel(intro)
            hint.setObjectName("cardHint")
            hint.setWordWrap(True)
            page_layout.addWidget(hint)

        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.NoFrame)
        host = QWidget()
        list_layout = QVBoxLayout(host)
        list_layout.setContentsMargins(2, 2, 8, 2)
        list_layout.setSpacing(7)
        scroll.setWidget(host)
        page_layout.addWidget(scroll, stretch=1)
        return page, list_layout

    def _active_sheet(self) -> str:
        data = self.month_combo.currentData()
        return str(data) if data is not None else ""

    def _row_matches_filters(self, item: SqtSelectionItem) -> bool:
        if (item.sheet_name or "") != self._active_sheet():
            return False
        query = self.search_input.text().strip().casefold()
        return not query or query in item.value.casefold()

    @Slot()
    def _apply_filters(self, *_args) -> None:
        active_sheet = self._active_sheet()
        pending_count = 0
        imported_count = 0
        for item, row in self._rows:
            in_sheet = (item.sheet_name or "") == active_sheet
            if in_sheet:
                if item.is_fully_imported:
                    imported_count += 1
                else:
                    pending_count += 1
            row.setVisible(in_sheet and self._row_matches_filters(item))
        self.tabs.setTabText(0, f"CẦN NHẬP ({pending_count})")
        self.tabs.setTabText(1, f"ĐÃ NHẬP ({imported_count})")

    @Slot()
    def _on_month_changed(self, *_args) -> None:
        active_sheet = self._active_sheet()
        for item, row in self._rows:
            if (item.sheet_name or "") != active_sheet:
                row.checkbox.setChecked(False)
        self._apply_filters()
        self._update_summary()

    def _active_tab_rows(self) -> list[tuple[SqtSelectionItem, _SqtChoiceRow]]:
        imported_tab = self.tabs.currentIndex() == 1
        return [
            (item, row)
            for item, row in self._rows
            if item.is_fully_imported == imported_tab
            and self._row_matches_filters(item)
        ]

    @Slot()
    def select_all(self) -> None:
        for _item, row in self._active_tab_rows():
            row.checkbox.setChecked(True)

    @Slot()
    def clear_all(self) -> None:
        for _item, row in self._active_tab_rows():
            row.checkbox.setChecked(False)

    def selected_values(self) -> List[str]:
        if self._selected_values:
            return list(self._selected_values)
        return [
            item.value
            for item, row in self._rows
            if row.checkbox.isChecked()
        ]

    def selected_items(self) -> List[SqtSelectionItem]:
        if self._selected_items:
            return list(self._selected_items)
        return [
            item
            for item, row in self._rows
            if row.checkbox.isChecked()
        ]

    @Slot()
    def _update_summary(self, *_args) -> None:
        selected = self.selected_items()
        row_count = sum(item.row_count for item in selected)
        imported = sum(1 for item in selected if item.is_fully_imported)
        pending = len(selected) - imported
        unit = "khoản chi" if self.is_expense else "dòng"
        detail = ""
        if selected:
            detail = f" · {pending} cần nhập · {imported} nhập lại"
        self.selection_summary.setText(
            f"Đã chọn {len(selected)} SQT · {row_count} {unit}{detail}"
        )

    @Slot()
    def _on_run(self) -> None:
        selected_items = self.selected_items()
        if not selected_items:
            notify.warning(
                self,
                "Chưa chọn SQT",
                "Vui lòng chọn ít nhất một Số quyết toán.",
            )
            return

        fully_imported = sum(1 for item in selected_items if item.is_fully_imported)
        partial = sum(1 for item in selected_items if item.is_partially_imported)
        if fully_imported or partial:
            details = []
            if fully_imported:
                details.append(f"{fully_imported} SQT đã nhập sẽ được nhập lại")
            if partial:
                details.append(
                    f"{partial} SQT mới nhập một phần sẽ được xử lý lại toàn bộ"
                )
            confirmed = notify.confirm(
                self,
                "Xác nhận nhập lại",
                "Bạn đang chọn:\n• "
                + "\n• ".join(details)
                + "\n\nBạn có muốn tiếp tục?",
                confirm_text="Tiếp tục nhập lại",
                cancel_text="Quay lại kiểm tra",
                tone="warning",
            )
            if not confirmed:
                return

        self._selected_items = selected_items
        self._selected_values = [item.value for item in selected_items]
        self.accept()
