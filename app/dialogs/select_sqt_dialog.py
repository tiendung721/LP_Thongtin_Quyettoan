"""Chọn SQT bằng bảng tổng quan và khung chi tiết từng dòng Excel."""

from __future__ import annotations

from collections import OrderedDict
from datetime import date
from typing import Iterable, List

from PySide6.QtCore import Qt, Slot
from PySide6.QtGui import QColor
from PySide6.QtWidgets import (
    QComboBox,
    QFrame,
    QGridLayout,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QLineEdit,
    QPushButton,
    QScrollArea,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)

from .base_dialog import ResizableDialog
from .notifications import notify
from ..daily_import import month_sheet_number
from ..services.sqt_selection_service import (
    EXPENSE_SELECTION_OPERATION,
    SELECTION_OPERATION,
    SqtSelectionItem,
    SqtSourceRow,
)


def _label(text: str, name: str = "") -> QLabel:
    result = QLabel(text)
    if name:
        result.setObjectName(name)
    return result


def _money(value: float | None) -> str:
    return f"{value:,.0f}".replace(",", ".") + " đ" if value is not None else "—"


def _date_text(value: str) -> str:
    if not value:
        return "—"
    try:
        return date.fromisoformat(value).strftime("%d/%m/%Y")
    except ValueError:
        return value


def _date_preview(item: SqtSelectionItem) -> str:
    values = list(dict.fromkeys(row.close_date for row in item.source_rows if row.close_date))
    if not values:
        return "—"
    if len(values) == 1:
        return _date_text(values[0])
    try:
        ordered = sorted(date.fromisoformat(value) for value in values)
        first, last = ordered[0], ordered[-1]
    except ValueError:
        return f"{_date_text(values[0])} – {_date_text(values[-1])}"
    if (first.month, first.year) == (last.month, last.year):
        return f"{first.day}–{last.strftime('%d/%m/%Y')}"
    return f"{first.strftime('%d/%m/%Y')} – {last.strftime('%d/%m/%Y')}"


def _containers(item: SqtSelectionItem) -> list[str]:
    return list(dict.fromkeys(row.container for row in item.source_rows if row.container))


def _container_preview(item: SqtSelectionItem) -> str:
    values = _containers(item)
    if not values:
        return "—"
    preview = ", ".join(values[:2])
    return f"{preview} +{len(values) - 2}" if len(values) > 2 else preview


def _total_text(item: SqtSelectionItem) -> str:
    amounts = [row.total_amount for row in item.source_rows if row.total_amount is not None]
    if not amounts:
        return "—"
    if len(amounts) < item.row_count:
        return f"{_money(sum(amounts))} (thiếu)"
    return _money(sum(amounts))


class SelectSqtDialog(ResizableDialog):
    """Chọn SQT trong một sheet tháng và xem nhanh dữ liệu trước khi chạy PAD."""

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
        self._display_items: list[SqtSelectionItem] = []
        self._checked_keys: set[tuple[str, str]] = set()
        self._selected_values: List[str] = []
        self._selected_items: List[SqtSelectionItem] = []

        action = "nhập khoản chi" if self.is_expense else "nhập thông tin"
        self.setWindowTitle(f"Chọn Số quyết toán {action}")
        self.setModal(True)
        self.resize(1260, 760)
        self.setMinimumSize(960, 580)

        root = QVBoxLayout(self)
        root.setContentsMargins(18, 16, 18, 16)
        root.setSpacing(11)

        heading = QHBoxLayout()
        heading.addWidget(_label("Chọn số quyết toán", "sqtDialogTitle"))
        heading.addWidget(_label(f"{len(self.items)} SQT trong file", "sqtCountBadge"))
        heading.addStretch(1)
        heading.addWidget(_label("NHẬP KHOẢN CHI" if self.is_expense else "NHẬP THÔNG TIN", "sqtDialogKind"))
        root.addLayout(heading)

        filters = QGridLayout()
        filters.setHorizontalSpacing(12)
        filters.setVerticalSpacing(5)
        filters.addWidget(_label("Tìm SQT hoặc số Cont", "sqtFilterLabel"), 0, 0)
        filters.addWidget(_label("Trạng thái nhập", "sqtFilterLabel"), 0, 1)
        self.search_input = QLineEdit()
        self.search_input.setPlaceholderText("Nhập SQT hoặc số Container...")
        self.search_input.setClearButtonEnabled(True)
        filters.addWidget(self.search_input, 1, 0)
        self.status_combo = QComboBox()
        self.status_combo.addItem("Tất cả", "all")
        self.status_combo.addItem("Chưa nhập", "pending")
        if not self.is_expense:
            self.status_combo.addItem("Nhập một phần", "partial")
        self.status_combo.addItem("Đã nhập", "done")
        filters.addWidget(self.status_combo, 1, 1)
        self.month_label = _label("Tháng theo SQT" if self.is_expense else "Sheet tháng", "sqtFilterLabel")
        filters.addWidget(self.month_label, 2, 0)
        filters.addWidget(_label("Sắp xếp", "sqtFilterLabel"), 2, 1)
        self.month_combo = QComboBox()
        self._populate_months()
        show_month = any(
            item.month_sheet_name if self.is_expense else item.sheet_name
            for item in self.items
        )
        self.month_combo.setVisible(show_month)
        filters.addWidget(self.month_combo, 3, 0)
        self.source_label = _label("Khoan_Chi" if self.is_expense else "Thong_Tin_Quyet_Toan", "sqtSourceLabel")
        self.source_label.setVisible(not show_month)
        filters.addWidget(self.source_label, 3, 0)
        self.sort_combo = QComboBox()
        self.sort_combo.addItem("Chưa nhập trước", "status")
        self.sort_combo.addItem("SQT tăng dần", "sqt")
        self.sort_combo.addItem("Ngày đóng mới trước", "date")
        filters.addWidget(self.sort_combo, 3, 1)
        bulk = QHBoxLayout()
        self.btn_select_all = QPushButton("Chọn hiển thị")
        self.btn_clear_all = QPushButton("Bỏ chọn")
        self.btn_select_all.clicked.connect(self.select_all)
        self.btn_clear_all.clicked.connect(self.clear_all)
        bulk.addWidget(self.btn_select_all)
        bulk.addWidget(self.btn_clear_all)
        filters.addLayout(bulk, 3, 2)
        filters.setColumnStretch(0, 6)
        filters.setColumnStretch(1, 3)
        filters.setColumnStretch(2, 3)
        root.addLayout(filters)

        self.top_summary = _label("Chưa chọn số quyết toán nào", "sqtTopSummary")
        root.addWidget(self.top_summary)

        content = QHBoxLayout()
        content.setSpacing(12)
        self.table = QTableWidget()
        self.table.setObjectName("sqtSelectionTable")
        headers = ["Chọn", "SQT", "Số dòng" if not self.is_expense else "Khoản chi", "Số Cont", "Ngày đóng", "Tiến độ nhập"]
        if self.is_expense:
            headers.append("Tổng tiền")
        self.table.setColumnCount(len(headers))
        self.table.setHorizontalHeaderLabels(headers)
        self.table.verticalHeader().hide()
        self.table.setAlternatingRowColors(True)
        self.table.setSelectionBehavior(QTableWidget.SelectRows)
        self.table.setSelectionMode(QTableWidget.SingleSelection)
        self.table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.table.setWordWrap(False)
        header = self.table.horizontalHeader()
        header.setSectionResizeMode(QHeaderView.Fixed)
        header.setSectionResizeMode(3, QHeaderView.Stretch)
        for index, width in enumerate((48, 75, 78, 240, 118, 130, 130)[:len(headers)]):
            if index != 3:
                self.table.setColumnWidth(index, width)
        self.table.cellChanged.connect(self._cell_changed)
        self.table.itemSelectionChanged.connect(self._show_selected_detail)
        content.addWidget(self.table, 8)

        panel = QFrame()
        panel.setObjectName("sqtDetailPanel")
        panel_layout = QVBoxLayout(panel)
        panel_layout.setContentsMargins(11, 11, 11, 11)
        panel_layout.setSpacing(9)
        self.detail_title = _label("Chi tiết SQT", "sqtDetailTitle")
        panel_layout.addWidget(self.detail_title)
        self.detail_meta = _label("", "sqtDetailMeta")
        self.detail_meta.setWordWrap(True)
        panel_layout.addWidget(self.detail_meta)
        self.detail_status = _label("", "sqtDetailStatus")
        panel_layout.addWidget(self.detail_status, alignment=Qt.AlignLeft)
        divider = QFrame()
        divider.setObjectName("sqtDivider")
        panel_layout.addWidget(divider)
        panel_layout.addWidget(_label("CÁC DÒNG TRONG SQT", "sqtFilterLabel"))
        self.detail_scroll = QScrollArea()
        self.detail_scroll.setWidgetResizable(True)
        self.detail_host = QWidget()
        self.detail_host.setObjectName("sqtDetailHost")
        self.detail_list = QVBoxLayout(self.detail_host)
        self.detail_list.setContentsMargins(0, 0, 3, 0)
        self.detail_list.setSpacing(8)
        self.detail_scroll.setWidget(self.detail_host)
        panel_layout.addWidget(self.detail_scroll, 1)
        content.addWidget(panel, 3)
        root.addLayout(content, 1)

        footer = QHBoxLayout()
        self.selection_summary = _label("", "sqtSelectionSummary")
        footer.addWidget(self.selection_summary)
        footer.addStretch(1)
        self.btn_run = QPushButton("Chạy nhập khoản chi" if self.is_expense else "Chạy nhập thông tin")
        self.btn_run.setProperty("variant", "primary")
        self.btn_run.clicked.connect(self._on_run)
        footer.addWidget(self.btn_run)
        cancel = QPushButton("Hủy")
        cancel.clicked.connect(self.reject)
        footer.addWidget(cancel)
        root.addLayout(footer)

        self.search_input.textChanged.connect(self._apply_filters)
        self.status_combo.currentIndexChanged.connect(self._apply_filters)
        self.sort_combo.currentIndexChanged.connect(self._rebuild_table)
        self.month_combo.currentIndexChanged.connect(self._on_month_changed)
        self._rebuild_table()

    def _populate_months(self) -> None:
        grouped: OrderedDict[str, int] = OrderedDict()
        for item in self.items:
            sheet = (item.month_sheet_name if self.is_expense else item.sheet_name) or ""
            grouped[sheet] = grouped.get(sheet, 0) + 1
        if self.is_expense:
            self.month_combo.addItem(f"Tất cả — {len(self.items)} SQT", "__all__")
            for sheet in sorted(grouped, key=lambda name: month_sheet_number(name) or 99):
                label = f"{sheet} — {grouped[sheet]} SQT" if sheet else f"Chưa xác định tháng — {grouped[sheet]} SQT"
                self.month_combo.addItem(label, sheet or "__unknown__")
            return
        for sheet, count in grouped.items():
            self.month_combo.addItem(f"{sheet} — {count} SQT" if sheet else f"{count} SQT", sheet)

    def _active_sheet(self) -> str:
        return str(self.month_combo.currentData() or "")

    @staticmethod
    def _key(item: SqtSelectionItem) -> tuple[str, str]:
        return item.sheet_name, item.value

    @staticmethod
    def _sqt_sort_key(item: SqtSelectionItem) -> tuple[int, int | str]:
        return (0, int(item.value)) if item.value.isdigit() else (1, item.value.casefold())

    def _sorted_items(self) -> list[SqtSelectionItem]:
        result = list(self.items)
        mode = self.sort_combo.currentData()
        if mode == "status":
            rank = {"pending": 0, "partial": 1, "done": 2}
            result.sort(key=lambda item: (rank[item.status_kind], self._sqt_sort_key(item)))
        elif mode == "sqt":
            result.sort(key=self._sqt_sort_key)
        elif mode == "date":
            result.sort(key=lambda item: max((row.close_date for row in item.source_rows), default=""), reverse=True)
        return result

    def _current_item(self) -> SqtSelectionItem | None:
        row = self.table.currentRow()
        return self._display_items[row] if 0 <= row < len(self._display_items) else None

    @Slot()
    def _rebuild_table(self, *_args) -> None:
        previous = self._current_item()
        self._display_items = self._sorted_items()
        self.table.blockSignals(True)
        self.table.setRowCount(len(self._display_items))
        for row, item in enumerate(self._display_items):
            self.table.setRowHeight(row, 54)
            checkbox = QTableWidgetItem()
            checkbox.setFlags(Qt.ItemIsEnabled | Qt.ItemIsSelectable | Qt.ItemIsUserCheckable)
            checkbox.setCheckState(Qt.Checked if self._key(item) in self._checked_keys else Qt.Unchecked)
            self.table.setItem(row, 0, checkbox)
            values = [item.value, str(item.row_count), _container_preview(item), _date_preview(item), item.status_text]
            if self.is_expense:
                values.append(_total_text(item))
            for column, value in enumerate(values, start=1):
                cell = QTableWidgetItem(value)
                cell.setFlags(Qt.ItemIsEnabled | Qt.ItemIsSelectable)
                if column in (1, 2):
                    cell.setTextAlignment(Qt.AlignCenter)
                if column == 3:
                    cell.setToolTip(", ".join(_containers(item)) or "Chưa có số Cont")
                if column == 5:
                    cell.setForeground(QColor("#15803D" if item.is_fully_imported else "#B45309"))
                if column == 6 and "(thiếu)" in value:
                    cell.setToolTip("Một số khoản chi chưa có Tổng tiền; số hiển thị chỉ cộng các khoản đã có giá trị.")
                self.table.setItem(row, column, cell)
        self.table.blockSignals(False)
        self._apply_filters(preferred=previous)

    def _matches_filters(self, item: SqtSelectionItem) -> bool:
        active_sheet = self._active_sheet()
        if self.is_expense:
            if active_sheet != "__all__" and (item.month_sheet_name or "__unknown__") != active_sheet:
                return False
        elif item.sheet_name != active_sheet:
            return False
        status = self.status_combo.currentData()
        if status != "all" and item.status_kind != status:
            return False
        query = self.search_input.text().strip().casefold()
        return not query or query in item.value.casefold() or any(query in row.container.casefold() for row in item.source_rows)

    @Slot()
    def _apply_filters(self, *_args, preferred: SqtSelectionItem | None = None) -> None:
        visible: list[int] = []
        for row, item in enumerate(self._display_items):
            matches = self._matches_filters(item)
            self.table.setRowHidden(row, not matches)
            if matches:
                visible.append(row)
        current = self._current_item()
        target = next((row for row in visible if self._display_items[row] == preferred), None)
        if target is None and current is not None:
            target = next((row for row in visible if self._display_items[row] == current), None)
        if target is None and visible:
            target = visible[0]
        if target is None:
            self.table.clearSelection()
            self._render_detail(None)
        else:
            self.table.selectRow(target)
            self._render_detail(self._display_items[target])
        self._update_summary(visible_count=len(visible))

    @Slot()
    def _on_month_changed(self, *_args) -> None:
        if self.is_expense:
            self._apply_filters()
            return
        self._checked_keys.clear()
        self.table.blockSignals(True)
        for row in range(self.table.rowCount()):
            self.table.item(row, 0).setCheckState(Qt.Unchecked)
        self.table.blockSignals(False)
        self._apply_filters()

    @Slot()
    def _cell_changed(self, row: int, column: int) -> None:
        if column != 0 or row >= len(self._display_items):
            return
        key = self._key(self._display_items[row])
        if self.table.item(row, 0).checkState() == Qt.Checked:
            self._checked_keys.add(key)
        else:
            self._checked_keys.discard(key)
        self._update_summary()

    @Slot()
    def select_all(self) -> None:
        for row, item in enumerate(self._display_items):
            if not self.table.isRowHidden(row):
                self._checked_keys.add(self._key(item))
                self.table.item(row, 0).setCheckState(Qt.Checked)
        self._update_summary()

    @Slot()
    def clear_all(self) -> None:
        for row, item in enumerate(self._display_items):
            if not self.table.isRowHidden(row):
                self._checked_keys.discard(self._key(item))
                self.table.item(row, 0).setCheckState(Qt.Unchecked)
        self._update_summary()

    def selected_items(self) -> List[SqtSelectionItem]:
        if self._selected_items:
            return list(self._selected_items)
        return [item for item in self.items if self._key(item) in self._checked_keys]

    def selected_values(self) -> List[str]:
        if self._selected_values:
            return list(self._selected_values)
        return [item.value for item in self.selected_items()]

    @Slot()
    def _update_summary(self, *_args, visible_count: int | None = None) -> None:
        selected = self.selected_items()
        count = sum(item.row_count for item in selected)
        pending = sum(not item.is_fully_imported for item in selected)
        imported = len(selected) - pending
        unit = "khoản chi" if self.is_expense else "dòng"
        if selected:
            self.top_summary.setText(f"Đã chọn {len(selected)} SQT · {count} {unit} · {pending} cần nhập · {imported} nhập lại")
        else:
            self.top_summary.setText("Chưa chọn số quyết toán nào")
        if visible_count is None:
            visible_count = sum(not self.table.isRowHidden(row) for row in range(self.table.rowCount()))
        self.selection_summary.setText(f"Hiển thị {visible_count} SQT · Đã chọn {len(selected)} SQT / {count} {unit}")

    @Slot()
    def _show_selected_detail(self) -> None:
        self._render_detail(self._current_item())

    def _clear_detail(self) -> None:
        while self.detail_list.count():
            child = self.detail_list.takeAt(0)
            widget = child.widget()
            if widget:
                widget.hide()
                widget.setParent(None)
                widget.deleteLater()

    def _render_detail(self, item: SqtSelectionItem | None) -> None:
        self._clear_detail()
        if item is None:
            self.detail_title.setText("Không có SQT phù hợp")
            self.detail_meta.setText("Thử thay đổi bộ lọc hoặc từ khóa tìm kiếm.")
            self.detail_status.hide()
            return
        self.detail_title.setText(f"SQT {item.value}")
        unit = "khoản chi" if self.is_expense else "dòng"
        meta = f"{item.row_count} {unit} · {len(_containers(item))} Cont · Ngày đóng {_date_preview(item)}"
        if self.is_expense:
            meta += f" · {item.month_sheet_name or 'Chưa xác định tháng'}\nTổng tiền: {_total_text(item)}"
        self.detail_meta.setText(meta)
        self.detail_status.setText(item.status_text)
        self.detail_status.setProperty("tone", item.status_kind)
        self.detail_status.show()
        self.detail_status.style().unpolish(self.detail_status)
        self.detail_status.style().polish(self.detail_status)
        if not item.source_rows:
            self.detail_list.addWidget(_label("File chưa có thông tin chi tiết để hiển thị.", "sqtDetailMeta"))
        for source in item.source_rows:
            self.detail_list.addWidget(self._detail_card(source))
        self.detail_list.addStretch(1)
        self.detail_scroll.verticalScrollBar().setValue(0)

    def _detail_card(self, source: SqtSourceRow) -> QFrame:
        card = QFrame()
        card.setObjectName("sqtDetailCard")
        body = QVBoxLayout(card)
        body.setContentsMargins(11, 9, 11, 9)
        body.setSpacing(5)
        top = QHBoxLayout()
        title = _label(f"Dòng Excel {source.row_number} · {source.container or 'Chưa có Cont'}", "sqtDetailRowTitle")
        title.setWordWrap(True)
        top.addWidget(title, 1)
        top.addWidget(_label("Đã nhập" if source.imported else "Chưa nhập", "sqtDetailMeta"))
        body.addLayout(top)
        body.addWidget(_label(f"Ngày đóng: {_date_text(source.close_date)}", "sqtDetailMeta"))
        if self.is_expense:
            body.addWidget(_label(f"Tổng tiền: {_money(source.total_amount)}", "sqtDetailAmount"))
            if source.material_price is not None:
                body.addWidget(_label(f"Giá VL: {_money(source.material_price)}", "sqtDetailMeta"))
        return card

    @Slot()
    def _on_run(self) -> None:
        selected_items = self.selected_items()
        if not selected_items:
            notify.warning(self, "Chưa chọn SQT", "Vui lòng chọn ít nhất một Số quyết toán.")
            return

        fully_imported = sum(item.is_fully_imported for item in selected_items)
        partial = sum(item.is_partially_imported for item in selected_items)
        if fully_imported or partial:
            details = []
            if fully_imported:
                details.append(f"{fully_imported} SQT đã nhập sẽ được nhập lại")
            if partial:
                details.append(f"{partial} SQT mới nhập một phần sẽ được xử lý lại toàn bộ")
            confirmed = notify.confirm(
                self,
                "Xác nhận nhập lại",
                "Bạn đang chọn:\n• " + "\n• ".join(details) + "\n\nBạn có muốn tiếp tục?",
                confirm_text="Tiếp tục nhập lại",
                cancel_text="Quay lại kiểm tra",
                tone="warning",
            )
            if not confirmed:
                return

        self._selected_items = selected_items
        self._selected_values = [item.value for item in selected_items]
        self.accept()
