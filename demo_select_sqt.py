"""Demo độc lập: chọn SQT bằng bảng và khung chi tiết.

Chạy: python demo_select_sqt.py
Xuất ảnh: python demo_select_sqt.py --capture
Chỉ dùng dữ liệu mẫu; không đọc Excel, ghi JSON hay chạy RPA.
"""

from __future__ import annotations

import sys
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path

from PySide6.QtCore import Qt
from PySide6.QtGui import QColor
from PySide6.QtWidgets import (
    QApplication,
    QComboBox,
    QDialog,
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


STYLE = """
* { font-family: 'Segoe UI', Arial, sans-serif; font-size: 10pt; color: #122B4C; }
QDialog, QWidget#root { background: #F4F7FB; }
QLabel#title { font-size: 17pt; font-weight: 700; color: #0B315C; }
QLabel#subtitle, QLabel#muted { color: #667D99; }
QLabel#filterLabel { color: #315986; font-size: 9pt; }
QLabel#countBadge { color: #1654C4; background: #E5EFFF; padding: 7px 11px; border-radius: 9px; font-weight: 600; }
QLabel#summary { color: #173C68; background: #EDF4FD; border: 1px solid #D3E1F3; border-radius: 7px; padding: 8px 12px; font-weight: 600; }
QLabel#detailTitle { font-size: 12pt; font-weight: 700; color: #0E3159; }
QLabel#detailStatusPending { color: #A65E05; background: #FFF0D6; border-radius: 8px; padding: 4px 8px; font-weight: 600; }
QLabel#detailStatusPartial { color: #A65E05; background: #FFF0D6; border-radius: 8px; padding: 4px 8px; font-weight: 600; }
QLabel#detailStatusDone { color: #15744A; background: #DFF5E9; border-radius: 8px; padding: 4px 8px; font-weight: 600; }
QLabel#detailMeta { color: #4A6381; }
QLabel#itemTitle { color: #17385F; font-weight: 700; }
QLabel#itemValue { color: #0D3569; font-weight: 600; }
QLabel#footer { color: #315986; }
QFrame#detailPanel, QFrame#detailCard { background: #FFFFFF; border: 1px solid #D8E3F0; border-radius: 9px; }
QFrame#detailCard:hover { border-color: #9BC2F7; }
QFrame#divider { background: #DBE5F0; max-height: 1px; min-height: 1px; }
QLineEdit, QComboBox { background: #FFFFFF; border: 1px solid #C8D6E7; border-radius: 7px; padding: 8px 10px; min-height: 22px; }
QLineEdit:focus, QComboBox:focus { border-color: #3A78E8; }
QPushButton { background: #FFFFFF; border: 1px solid #C8D6E7; border-radius: 8px; padding: 8px 15px; }
QPushButton:hover { background: #EBF3FF; }
QPushButton#primary { color: #FFFFFF; background: #2862DF; border-color: #2862DF; font-weight: 700; }
QPushButton#primary:hover { background: #1D51BD; }
QPushButton#switch { color: #245CB1; border-color: #AFCBF2; }
QTableWidget { background: #FFFFFF; alternate-background-color: #F8FAFD; border: 1px solid #D8E3F0; border-radius: 8px; gridline-color: #E0E8F2; selection-background-color: #DEEBFF; selection-color: #0C315C; outline: 0; }
QTableWidget::item { padding: 7px 8px; }
QTableWidget::item:selected { background: #DCEAFF; color: #0C315C; }
QHeaderView::section { background: #EAF0F7; color: #173D68; border: 0; border-right: 1px solid #D6E1EE; border-bottom: 1px solid #D6E1EE; padding: 9px 6px; font-weight: 600; }
QScrollArea { border: 0; background: transparent; }
"""


@dataclass(frozen=True)
class DetailLine:
    container: str
    close_date: str
    imported: bool = False
    total: int | None = None
    material_price: int | None = None


@dataclass(frozen=True)
class SqtGroup:
    sqt: str
    lines: tuple[DetailLine, ...]
    month: str = "Tháng 10"

    @property
    def imported_count(self) -> int:
        return sum(line.imported for line in self.lines)

    @property
    def status_kind(self) -> str:
        if self.imported_count == len(self.lines):
            return "done"
        return "partial" if self.imported_count else "pending"

    @property
    def status_text(self) -> str:
        if self.status_kind == "done":
            return "Đã nhập"
        if self.status_kind == "partial":
            return f"Đã nhập {self.imported_count}/{len(self.lines)}"
        return "Chưa nhập"


INFO_GROUPS = [
    SqtGroup("958", (DetailLine("TEMU4827316", "12/10/2026"),)),
    SqtGroup("959", (DetailLine("MSCU1298475", "12/10/2026"), DetailLine("TGHU8261540", "12/10/2026"))),
    SqtGroup("960", (DetailLine("TCLU5507241", "13/10/2026", True), DetailLine("CMAU4392810", "13/10/2026"), DetailLine("FSCU7038129", "14/10/2026"))),
    SqtGroup("961", (DetailLine("FSCU7038129", "14/10/2026"),)),
    SqtGroup("962", (DetailLine("OOLU6291483", "15/10/2026"), DetailLine("MSKU7603812", "15/10/2026"))),
    SqtGroup("963", (DetailLine("ABCD1234567", "15/10/2026"),)),
    SqtGroup("964", (DetailLine("ONEU9876543", "16/10/2026", True),)),
    SqtGroup("902", (DetailLine("HLBU3498710", "28/09/2026"),), "Tháng 9"),
]

EXPENSE_GROUPS = [
    SqtGroup("722", (DetailLine("TEMU4827316", "12/10/2026", total=4_950_000, material_price=3_000_000),)),
    SqtGroup("723", (DetailLine("MSCU1298475", "12/10/2026", total=3_140_000, material_price=1_800_000), DetailLine("TGHU8261540", "12/10/2026", total=5_100_000, material_price=3_400_000))),
    SqtGroup("716", (DetailLine("TCLU5507241", "13/10/2026", True, 4_080_000, 2_600_000), DetailLine("CMAU4392810", "13/10/2026", total=3_900_000, material_price=2_400_000), DetailLine("FSCU7038129", "14/10/2026", total=4_700_000, material_price=2_500_000))),
    SqtGroup("727", (DetailLine("FSCU7038129", "14/10/2026", total=3_850_000),)),
    SqtGroup("738", (DetailLine("OOLU6291483", "15/10/2026", total=6_120_000, material_price=4_100_000),)),
    SqtGroup("734", (DetailLine("MSKU7603812", "15/10/2026", total=5_670_000, material_price=3_450_000),)),
    SqtGroup("737", (DetailLine("ONEU9876543", "16/10/2026", True, 4_300_000, 2_800_000), DetailLine("ABCD1234567", "16/10/2026", True, 5_000_000, 3_200_000))),
]


def money(value: int | None) -> str:
    return f"{value:,}".replace(",", ".") + " đ" if value is not None else "—"


def label(text: str, name: str = "") -> QLabel:
    result = QLabel(text)
    if name:
        result.setObjectName(name)
    return result


def containers(group: SqtGroup) -> str:
    values = list(dict.fromkeys(line.container for line in group.lines if line.container))
    if not values:
        return "—"
    preview = ", ".join(values[:2])
    return f"{preview} +{len(values) - 2}" if len(values) > 2 else preview


def close_dates(group: SqtGroup) -> str:
    dates = sorted({datetime.strptime(line.close_date, "%d/%m/%Y") for line in group.lines if line.close_date})
    if not dates:
        return "—"
    first, last = dates[0], dates[-1]
    if first == last:
        return first.strftime("%d/%m/%Y")
    if (first.month, first.year) == (last.month, last.year):
        return f"{first.day}–{last.strftime('%d/%m/%Y')}"
    return f"{first.strftime('%d/%m/%Y')} – {last.strftime('%d/%m/%Y')}"


class SqtTableDemo(QDialog):
    def __init__(self, *, expense: bool):
        super().__init__()
        self.expense = expense
        self.groups = EXPENSE_GROUPS if expense else INFO_GROUPS
        self.visible_groups: list[SqtGroup] = []
        self.selected_keys: set[tuple[str, str]] = set()
        self.setWindowTitle("Demo — Chọn SQT nhập khoản chi" if expense else "Demo — Chọn SQT nhập thông tin")
        self.resize(1320, 790)
        self.setMinimumSize(1040, 620)

        root = QVBoxLayout(self)
        root.setContentsMargins(18, 16, 18, 16)
        root.setSpacing(11)

        heading = QHBoxLayout()
        heading.addWidget(label("Chọn số quyết toán", "title"))
        heading.addWidget(label(f"{len(self.groups)} SQT trong file", "countBadge"))
        heading.addStretch(1)
        heading.addWidget(label("NHẬP KHOẢN CHI" if expense else "NHẬP THÔNG TIN", "subtitle"))
        root.addLayout(heading)

        filters = QGridLayout()
        filters.setHorizontalSpacing(12)
        filters.setVerticalSpacing(5)
        filters.addWidget(label("Tìm SQT hoặc số Cont", "filterLabel"), 0, 0)
        filters.addWidget(label("Trạng thái nhập", "filterLabel"), 0, 1)
        self.search = QLineEdit()
        self.search.setPlaceholderText("Nhập SQT hoặc số Container...")
        self.search.setClearButtonEnabled(True)
        filters.addWidget(self.search, 1, 0)
        self.status = QComboBox()
        self.status.addItem("Tất cả", "all")
        self.status.addItem("Chưa nhập", "pending")
        self.status.addItem("Nhập một phần", "partial")
        self.status.addItem("Đã nhập", "done")
        filters.addWidget(self.status, 1, 1)
        filters.addWidget(label("Sheet tháng" if not expense else "Nguồn khoản chi", "filterLabel"), 2, 0)
        filters.addWidget(label("Sắp xếp", "filterLabel"), 2, 1)
        if expense:
            source = QLineEdit("Khoan_Chi")
            source.setReadOnly(True)
            filters.addWidget(source, 3, 0)
        else:
            self.month = QComboBox()
            self.month.addItem("Tháng 10 — 7 SQT", "Tháng 10")
            self.month.addItem("Tháng 9 — 1 SQT", "Tháng 9")
            filters.addWidget(self.month, 3, 0)
            self.month.currentIndexChanged.connect(self._change_month)
        self.sort = QComboBox()
        self.sort.addItem("Chưa nhập trước", "status")
        self.sort.addItem("SQT tăng dần", "sqt")
        self.sort.addItem("Ngày đóng mới trước", "date")
        filters.addWidget(self.sort, 3, 1)
        actions = QHBoxLayout()
        self.select_visible = QPushButton("Chọn hiển thị")
        self.select_visible.clicked.connect(lambda: self._bulk(True))
        clear = QPushButton("Bỏ chọn")
        clear.clicked.connect(lambda: self._bulk(False))
        actions.addWidget(self.select_visible)
        actions.addWidget(clear)
        filters.addLayout(actions, 3, 2)
        filters.setColumnStretch(0, 6)
        filters.setColumnStretch(1, 3)
        filters.setColumnStretch(2, 3)
        root.addLayout(filters)

        self.top_summary = label("Chưa chọn số quyết toán nào", "summary")
        root.addWidget(self.top_summary)

        content = QHBoxLayout()
        content.setSpacing(12)
        self.table = QTableWidget()
        headers = ["Chọn", "SQT", "Số dòng" if not expense else "Khoản chi", "Số Cont", "Ngày đóng", "Tiến độ nhập"]
        if expense:
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
        panel.setObjectName("detailPanel")
        panel_layout = QVBoxLayout(panel)
        panel_layout.setContentsMargins(11, 11, 11, 11)
        panel_layout.setSpacing(9)
        self.detail_title = label("Chi tiết SQT", "detailTitle")
        panel_layout.addWidget(self.detail_title)
        self.detail_meta = label("", "detailMeta")
        self.detail_meta.setWordWrap(True)
        panel_layout.addWidget(self.detail_meta)
        self.detail_status = label("", "detailStatusPending")
        panel_layout.addWidget(self.detail_status, alignment=Qt.AlignLeft)
        divider = QFrame()
        divider.setObjectName("divider")
        panel_layout.addWidget(divider)
        panel_layout.addWidget(label("CÁC DÒNG TRONG SQT", "filterLabel"))
        self.detail_scroll = QScrollArea()
        self.detail_scroll.setWidgetResizable(True)
        self.detail_host = QWidget()
        self.detail_host.setObjectName("root")
        self.detail_list = QVBoxLayout(self.detail_host)
        self.detail_list.setContentsMargins(0, 0, 3, 0)
        self.detail_list.setSpacing(8)
        self.detail_scroll.setWidget(self.detail_host)
        panel_layout.addWidget(self.detail_scroll, 1)
        content.addWidget(panel, 3)
        root.addLayout(content, 1)

        footer = QHBoxLayout()
        self.footer_label = label("", "footer")
        footer.addWidget(self.footer_label)
        footer.addStretch(1)
        switch = QPushButton("Xem cửa sổ thông tin" if expense else "Xem cửa sổ khoản chi")
        switch.setObjectName("switch")
        switch.clicked.connect(self._switch)
        footer.addWidget(switch)
        run = QPushButton("Chạy nhập khoản chi" if expense else "Chạy nhập thông tin")
        run.setObjectName("primary")
        run.setToolTip("Bản demo: không chạy RPA")
        footer.addWidget(run)
        cancel = QPushButton("Đóng demo")
        cancel.clicked.connect(self.close)
        footer.addWidget(cancel)
        root.addLayout(footer)

        self.search.textChanged.connect(self._refresh)
        self.status.currentIndexChanged.connect(self._refresh)
        self.sort.currentIndexChanged.connect(self._refresh)
        self._refresh()
        initial = next((i for i, group in enumerate(self.visible_groups) if group.status_kind == "partial"), 0)
        self.table.selectRow(initial)

    @staticmethod
    def _key(group: SqtGroup) -> tuple[str, str]:
        return group.month, group.sqt

    def _change_month(self) -> None:
        self.selected_keys.clear()
        self._refresh()

    def _filtered_groups(self) -> list[SqtGroup]:
        query = self.search.text().strip().casefold()
        chosen_status = self.status.currentData()
        month = None if self.expense else self.month.currentData()
        result = [
            group for group in self.groups
            if (month is None or group.month == month)
            and (chosen_status == "all" or group.status_kind == chosen_status)
            and (not query or query in group.sqt.casefold() or any(query in line.container.casefold() for line in group.lines))
        ]
        order = self.sort.currentData()
        if order == "status":
            rank = {"pending": 0, "partial": 1, "done": 2}
            result.sort(key=lambda group: (rank[group.status_kind], int(group.sqt)))
        elif order == "sqt":
            result.sort(key=lambda group: int(group.sqt))
        elif order == "date":
            result.sort(key=lambda group: max(datetime.strptime(line.close_date, "%d/%m/%Y") for line in group.lines), reverse=True)
        return result

    def _refresh(self) -> None:
        previous = self._current_group()
        self.visible_groups = self._filtered_groups()
        self.table.blockSignals(True)
        self.table.setRowCount(len(self.visible_groups))
        for row, group in enumerate(self.visible_groups):
            self.table.setRowHeight(row, 58)
            check = QTableWidgetItem()
            check.setFlags(Qt.ItemIsEnabled | Qt.ItemIsSelectable | Qt.ItemIsUserCheckable)
            check.setCheckState(Qt.Checked if self._key(group) in self.selected_keys else Qt.Unchecked)
            self.table.setItem(row, 0, check)
            values = [group.sqt, str(len(group.lines)), containers(group), close_dates(group), group.status_text]
            if self.expense:
                totals = [line.total for line in group.lines if line.total is not None]
                values.append(money(sum(totals)) if totals else "—")
            for column, value in enumerate(values, start=1):
                item = QTableWidgetItem(value)
                item.setFlags(Qt.ItemIsEnabled | Qt.ItemIsSelectable)
                if column in (1, 2):
                    item.setTextAlignment(Qt.AlignCenter)
                if column == 5:
                    item.setForeground(QColor("#147549" if group.status_kind == "done" else "#B56A0C"))
                self.table.setItem(row, column, item)
        self.table.blockSignals(False)
        if self.visible_groups:
            matched_row = next((i for i, group in enumerate(self.visible_groups) if group == previous), 0)
            self.table.selectRow(matched_row)
            self._show_selected_detail()
        else:
            self._render_detail(None)
        self._update_summary()

    def _current_group(self) -> SqtGroup | None:
        index = self.table.currentRow()
        return self.visible_groups[index] if 0 <= index < len(self.visible_groups) else None

    def _cell_changed(self, row: int, column: int) -> None:
        if column != 0 or row >= len(self.visible_groups):
            return
        key = self._key(self.visible_groups[row])
        if self.table.item(row, 0).checkState() == Qt.Checked:
            self.selected_keys.add(key)
        else:
            self.selected_keys.discard(key)
        self._update_summary()

    def _bulk(self, checked: bool) -> None:
        for group in self.visible_groups:
            if checked:
                self.selected_keys.add(self._key(group))
            else:
                self.selected_keys.discard(self._key(group))
        self._refresh()

    def _update_summary(self) -> None:
        selected = [group for group in self.groups if self._key(group) in self.selected_keys]
        count = sum(len(group.lines) for group in selected)
        unit = "khoản chi" if self.expense else "dòng"
        if selected:
            self.top_summary.setText(f"Đã chọn {len(selected)} SQT · {count} {unit} · {sum(group.status_kind == 'done' for group in selected)} SQT sẽ nhập lại")
        else:
            self.top_summary.setText("Chưa chọn số quyết toán nào")
        self.footer_label.setText(f"Hiển thị {len(self.visible_groups)} SQT  ·  Đã chọn {len(selected)} SQT / {count} {unit}")

    def _show_selected_detail(self) -> None:
        self._render_detail(self._current_group())

    def _render_detail(self, group: SqtGroup | None) -> None:
        while self.detail_list.count():
            child = self.detail_list.takeAt(0)
            widget = child.widget()
            if widget:
                widget.hide()
                widget.setParent(None)
                widget.deleteLater()
        if group is None:
            self.detail_title.setText("Không có SQT phù hợp")
            self.detail_meta.setText("Thử thay đổi bộ lọc hoặc từ khóa tìm kiếm.")
            self.detail_status.hide()
            return
        self.detail_title.setText(f"SQT {group.sqt}")
        meta = f"{len(group.lines)} {'khoản chi' if self.expense else 'dòng'} · {len({line.container for line in group.lines})} Cont · Ngày đóng {close_dates(group)}"
        if self.expense:
            totals = [line.total for line in group.lines if line.total is not None]
            if totals:
                meta += f"\nTổng tiền: {money(sum(totals))}"
        self.detail_meta.setText(meta)
        self.detail_status.setObjectName({"pending": "detailStatusPending", "partial": "detailStatusPartial", "done": "detailStatusDone"}[group.status_kind])
        self.detail_status.setText(group.status_text)
        self.detail_status.show()
        self.detail_status.style().unpolish(self.detail_status)
        self.detail_status.style().polish(self.detail_status)
        for index, line in enumerate(group.lines, start=1):
            card = QFrame()
            card.setObjectName("detailCard")
            box = QVBoxLayout(card)
            box.setContentsMargins(11, 9, 11, 9)
            box.setSpacing(5)
            card_top = QHBoxLayout()
            card_top.addWidget(label(f"Dòng {index} · {line.container}", "itemTitle"))
            card_top.addStretch(1)
            card_top.addWidget(label("Đã nhập" if line.imported else "Chưa nhập", "muted"))
            box.addLayout(card_top)
            box.addWidget(label(f"Ngày đóng: {line.close_date}", "detailMeta"))
            if self.expense:
                box.addWidget(label(f"Tổng tiền: {money(line.total)}", "itemValue"))
                if line.material_price is not None:
                    box.addWidget(label(f"Giá VL: {money(line.material_price)}", "detailMeta"))
            self.detail_list.addWidget(card)
        self.detail_list.addStretch(1)
        self.detail_scroll.verticalScrollBar().setValue(0)

    def _switch(self) -> None:
        self.other = SqtTableDemo(expense=not self.expense)
        self.other.show()
        self.close()


def main() -> int:
    app = QApplication(sys.argv)
    app.setStyle("Fusion")
    app.setStyleSheet(STYLE)
    if "--capture" in sys.argv:
        output = Path(__file__).resolve().parent / "demo_preview"
        output.mkdir(exist_ok=True)
        for expense, filename in ((False, "sqt_info_table.png"), (True, "sqt_expense_table.png")):
            dialog = SqtTableDemo(expense=expense)
            dialog.show()
            app.processEvents()
            path = output / filename
            dialog.grab().save(str(path))
            print(path)
            dialog.close()
        return 0
    dialog = SqtTableDemo(expense=False)
    dialog.show()
    return app.exec()


if __name__ == "__main__":
    raise SystemExit(main())
