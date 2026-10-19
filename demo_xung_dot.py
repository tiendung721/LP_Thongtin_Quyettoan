"""Demo độc lập các cửa sổ xử lý xung đột A1–A4, B1–B3.

Chạy: python demo_xung_dot.py
Mọi dữ liệu trong cửa sổ này là giả lập và chỉ tồn tại trong bộ nhớ.
"""

from __future__ import annotations

import copy
import sys
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path

from PySide6.QtCore import Qt
from PySide6.QtGui import QFontDatabase
from PySide6.QtWidgets import (
    QApplication, QComboBox, QDialog, QFormLayout, QFrame, QHBoxLayout,
    QHeaderView, QLabel, QLineEdit, QMessageBox, QPushButton, QRadioButton,
    QScrollArea, QTableWidget, QTableWidgetItem, QVBoxLayout, QWidget,
)


STYLE = """
QWidget { font-family: 'Segoe UI', Arial; font-size: 10pt; color: #1E293B; }
QDialog { background: #F4F7FB; }
QFrame#card { background: white; border: 1px solid #DCE5F0; border-radius: 10px; }
QLabel#title { font-size: 18pt; font-weight: 700; color: #102541; }
QLabel#heading { font-size: 13pt; font-weight: 700; color: #102541; }
QLabel#hint { color: #60718A; }
QLabel#badge { color: #9A5A10; background: #FFF1D8; border-radius: 6px; padding: 5px 9px; }
QPushButton { background: white; border: 1px solid #C7D3E2; border-radius: 7px; padding: 7px 12px; }
QPushButton:hover { background: #EDF3FB; }
QPushButton:disabled { color: #9BA8B8; background: #F1F3F6; }
QPushButton[primary='true'] { color: white; background: #285DB7; border-color: #285DB7; font-weight: 700; }
QLineEdit, QComboBox { background: white; border: 1px solid #C7D3E2; border-radius: 6px; padding: 6px; }
QTableWidget { background: white; gridline-color: #E4EAF2; selection-background-color: #DDEBFF; selection-color: #172C4A; }
QHeaderView::section { background: #EDF2F8; padding: 8px; border: 0; border-right: 1px solid #DCE5F0; }
"""


@dataclass
class FieldDifference:
    name: str
    current: str
    incoming: str


@dataclass
class DemoIssue:
    code: str
    title: str
    doc: str
    reference: str
    container: str
    date: str
    reason: str
    candidates: list[tuple[str, str]] = field(default_factory=list)
    differences: list[FieldDifference] = field(default_factory=list)
    next_code: str = ""
    status: str = "Chưa xử lý"
    decision: str = ""
    extra_fields: dict[str, str] = field(default_factory=dict)


def sample_issues() -> list[DemoIssue]:
    return [
        DemoIssue(
            "A1", "Nhiều Bill", "Phiếu cân", "Seal SC-014", "ABCD1234567", "12/06/2026",
            "Một Container có nhiều Bill trong dữ liệu nhập. Chọn Bill đúng để ghép với Phiếu cân.",
            [("BL-001", "Bill BL-001 · Tàu AN PHÚ · chạy 13/06/2026"),
             ("BL-002", "Bill BL-002 · Tàu MINH LONG · chạy 14/06/2026")],
            next_code="A2",
        ),
        DemoIssue(
            "A2", "Nhiều SQT", "Phiếu cân / Bill", "Seal SC-014", "ABCD1234567", "12/06/2026",
            "Có nhiều dòng quyết toán đích phù hợp. Chọn đúng dòng Excel.",
            [("A-SQT-200", "Tháng 6 · Dòng 18 · SQT 200 · Tàu AN PHÚ"),
             ("A-SQT-201", "Tháng 6 · Dòng 26 · SQT 201 · Tàu MINH LONG")],
            next_code="A4",
        ),
        DemoIssue(
            "A3", "Thiếu Ngày đóng", "Bill", "BL-003", "KLMN7654321", "",
            "File Excel chia sheet theo tháng; chưa xác định được sheet đích để tạo SQT.",
        ),
        DemoIssue(
            "A4", "Dữ liệu khác Excel", "Phiếu cân", "Seal SC-014", "ABCD1234567", "12/06/2026",
            "Giá trị JSON khác dòng quyết toán đã có. Chọn cách xử lý cho từng trường.",
            differences=[FieldDifference("Số tấn", "28", "30"),
                         FieldDifference("Nơi đóng", "Kho A", "Kho B")],
        ),
        DemoIssue(
            "B1", "Không tìm thấy SQT", "Khoản chi", "HĐ 88", "EFGH1234567", "20/06/2026",
            "Không tìm thấy SQT có ngày và Container tương ứng. Sửa thông tin để ghép lại.",
            next_code="B2",
        ),
        DemoIssue(
            "B2", "Nhiều SQT", "Khoản chi", "HĐ 89", "EFGH1234567", "03/06/2026",
            "Khoản chi khớp nhiều số quyết toán khác nhau. Chọn SQT tương ứng.",
            [("B-SQT-200", "Tháng 6 · Dòng 12 · SQT 200 · EFGH1234567"),
             ("B-SQT-300", "Tháng 6 · Dòng 19 · SQT 300 · EFGH1234567")],
            next_code="B3",
        ),
        DemoIssue(
            "B3", "Dữ liệu khác Excel", "Khoản chi", "HĐ 89", "EFGH1234567", "03/06/2026",
            "Khoản chi đã có trong Excel nhưng giá trị mới khác giá trị cũ.",
            differences=[FieldDifference("Đơn giá", "100.000", "120.000"),
                         FieldDifference("Tổng tiền", "2.800.000", "3.360.000")],
        ),
    ]


def label(text: str, role: str = "", wrap: bool = True) -> QLabel:
    widget = QLabel(text)
    widget.setWordWrap(wrap)
    if role:
        widget.setObjectName(role)
    return widget


def button(text: str, action, primary: bool = False) -> QPushButton:
    widget = QPushButton(text)
    if primary:
        widget.setProperty("primary", "true")
    widget.clicked.connect(action)
    return widget


def clear_layout(layout: QVBoxLayout) -> None:
    while layout.count():
        item = layout.takeAt(0)
        if item.widget():
            widget = item.widget()
            widget.hide()
            widget.setParent(None)
            widget.deleteLater()
        elif item.layout():
            clear_layout(item.layout())


class EditRowDialog(QDialog):
    EXTRA_LABELS = (
        "File nguồn", "Mã MD5 file", "Trạng thái", "Biển số xe", "Số tấn",
        "Loại hàng", "Nơi đóng", "Người nhận", "Tên tàu", "Ngày chạy",
        "VT biển", "Giá vật liệu", "Số HĐ", "Số B/L", "Số chì/Seal",
        "Đơn giá", "Thành tiền", "VAT", "Tổng tiền", "Trường khác",
        "Độ tin cậy", "Cảnh báo",
    )

    def __init__(self, issue: DemoIssue, parent=None):
        super().__init__(parent)
        self.issue = issue
        self.setWindowTitle(f"Chỉnh sửa dòng bóc tách · {issue.code} · Demo")
        self.resize(620, 670)
        outer = QVBoxLayout(self)
        outer.addWidget(label("Chỉnh sửa dòng bóc tách", "heading"))
        outer.addWidget(label("Sửa ngày và Container; demo sẽ phân tích lại tình huống này.", "hint"))
        form_host = QWidget()
        form = QFormLayout(form_host)
        self.doc = QLineEdit(issue.doc)
        self.reference = QLineEdit(issue.reference)
        self.date = QLineEdit(issue.date)
        self.container = QLineEdit(issue.container)
        for caption, widget in (("Loại chứng từ", self.doc), ("Số chứng từ", self.reference),
                                ("Ngày đóng / ngày ghép", self.date), ("Container", self.container)):
            form.addRow(caption, widget)
        self.extra_editors = {}
        for caption in self.EXTRA_LABELS:
            editor = QLineEdit(issue.extra_fields.get(caption, ""))
            self.extra_editors[caption] = editor
            form.addRow(caption, editor)
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setWidget(form_host)
        outer.addWidget(scroll, 1)
        outer.addWidget(label("Ví dụ: A3 thêm 12/06/2026; B1 sửa ngày hoặc Container.", "hint"))
        actions = QHBoxLayout()
        actions.addStretch()
        actions.addWidget(button("Quay lại", self.reject))
        actions.addWidget(button("Áp dụng chỉnh sửa", self.apply, True))
        outer.addLayout(actions)

    def apply(self):
        value = self.date.text().strip()
        if value:
            try:
                datetime.strptime(value, "%d/%m/%Y")
            except ValueError:
                QMessageBox.warning(self, "Ngày chưa hợp lệ", "Hãy nhập ngày theo dạng dd/mm/yyyy.")
                return
        self.issue.doc = self.doc.text().strip()
        self.issue.reference = self.reference.text().strip()
        self.issue.date = value
        self.issue.container = self.container.text().strip().upper()
        self.issue.extra_fields = {name: editor.text().strip() for name, editor in self.extra_editors.items() if editor.text().strip()}
        self.accept()


class BulkEditDialog(QDialog):
    def __init__(self, issues: list[DemoIssue], parent=None):
        super().__init__(parent)
        self.issues = issues
        self.setWindowTitle("Sửa nhanh nhiều dòng · Demo")
        self.resize(760, 420)
        outer = QVBoxLayout(self)
        outer.addWidget(label(f"Sửa thông tin để ghép lại {len(issues)} dòng", "heading"))
        self.table = QTableWidget(len(issues), 3)
        self.table.setHorizontalHeaderLabels(["Chứng từ", "Ngày đóng / ngày ghép", "Container"])
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        for row, issue in enumerate(issues):
            item = QTableWidgetItem(f"{issue.code} · {issue.reference}")
            item.setFlags(item.flags() & ~Qt.ItemIsEditable)
            self.table.setItem(row, 0, item)
            self.table.setItem(row, 1, QTableWidgetItem(issue.date))
            self.table.setItem(row, 2, QTableWidgetItem(issue.container))
        outer.addWidget(self.table)
        actions = QHBoxLayout()
        actions.addStretch()
        actions.addWidget(button("Hủy", self.reject))
        actions.addWidget(button("Kiểm tra và áp dụng", self.apply, True))
        outer.addLayout(actions)

    def apply(self):
        changes = []
        for row in range(self.table.rowCount()):
            date = self.table.item(row, 1).text().strip()
            container = self.table.item(row, 2).text().strip().upper()
            if date:
                try:
                    datetime.strptime(date, "%d/%m/%Y")
                except ValueError:
                    QMessageBox.warning(self, "Ngày chưa hợp lệ", f"Ngày ở dòng {row + 1} chưa hợp lệ.")
                    return
            changes.append((date, container))
        for issue, (date, container) in zip(self.issues, changes):
            issue.date, issue.container = date, container
            if date and container:
                issue.status = "Đã xử lý"
                issue.decision = "Đã sửa thông tin ghép"
        self.accept()


class BulkCandidateDialog(QDialog):
    def __init__(self, issues: list[DemoIssue], common: list[tuple[str, str]], parent=None):
        super().__init__(parent)
        self.issues = issues
        self.setWindowTitle("Chọn cùng một kết quả ghép · Demo")
        self.resize(780, 450)
        outer = QVBoxLayout(self)
        outer.addWidget(label(f"Áp dụng cho {len(issues)} dòng đã chọn", "heading"))
        self.choice = QComboBox()
        for code, description in common:
            self.choice.addItem(description, code)
        outer.addWidget(self.choice)
        table = QTableWidget(len(issues), 3)
        table.setHorizontalHeaderLabels(["Chứng từ", "Container", "Vấn đề"])
        table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        for row, issue in enumerate(issues):
            for col, value in enumerate((issue.reference, issue.container, issue.title)):
                table.setItem(row, col, QTableWidgetItem(value))
        outer.addWidget(table)
        actions = QHBoxLayout()
        actions.addStretch()
        actions.addWidget(button("Hủy", self.reject))
        actions.addWidget(button("Áp dụng cho các dòng này", self.apply, True))
        outer.addLayout(actions)

    def apply(self):
        for issue in self.issues:
            issue.status = "Đã xử lý"
            issue.decision = f"Đã chọn {self.choice.currentData()}"
        self.accept()


class BulkFieldDialog(QDialog):
    def __init__(self, issues: list[DemoIssue], parent=None):
        super().__init__(parent)
        self.issues = issues
        self.setWindowTitle("Chọn dữ liệu cho nhiều dòng · Demo")
        self.resize(860, 560)
        outer = QVBoxLayout(self)
        outer.addWidget(label("Chọn cách xử lý các trường khác nhau", "heading"))
        outer.addWidget(label("Hàng loạt chỉ có Giữ Excel hoặc Dùng mới; nhập thủ công ở từng dòng.", "hint"))
        names = sorted({item.name for issue in issues for item in issue.differences})
        choices = QTableWidget(len(names), 3)
        choices.setHorizontalHeaderLabels(["Trường", "Số dòng", "Cách xử lý"])
        choices.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.combos = {}
        for row, name in enumerate(names):
            choices.setItem(row, 0, QTableWidgetItem(name))
            choices.setItem(row, 1, QTableWidgetItem(str(sum(any(d.name == name for d in i.differences) for i in issues))))
            combo = QComboBox()
            combo.addItems(["Chọn cách xử lý", "Giữ dữ liệu Excel", "Dùng dữ liệu mới"])
            choices.setCellWidget(row, 2, combo)
            self.combos[name] = combo
        outer.addWidget(choices)
        preview = QTableWidget(sum(len(i.differences) for i in issues), 4)
        preview.setHorizontalHeaderLabels(["Chứng từ", "Trường", "Hiện tại", "Dữ liệu mới"])
        preview.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        row = 0
        for issue in issues:
            for diff in issue.differences:
                for col, value in enumerate((issue.reference, diff.name, diff.current, diff.incoming)):
                    preview.setItem(row, col, QTableWidgetItem(value))
                row += 1
        outer.addWidget(preview)
        actions = QHBoxLayout()
        actions.addStretch()
        actions.addWidget(button("Hủy", self.reject))
        actions.addWidget(button("Áp dụng cho các dòng này", self.apply, True))
        outer.addLayout(actions)

    def apply(self):
        if any(combo.currentIndex() == 0 for combo in self.combos.values()):
            QMessageBox.warning(self, "Chưa chọn đủ", "Hãy chọn cách xử lý cho từng trường.")
            return
        for issue in self.issues:
            issue.status = "Đã xử lý"
            issue.decision = "Đã chọn giá trị cho các trường"
        self.accept()


class ResolutionDemo(QDialog):
    def __init__(self):
        super().__init__()
        self.issues = sample_issues()
        self.selected: set[str] = set()
        self.visible: list[str] = []
        self.undo_state = None
        self.active_code = "A1"
        self.reviewing = False
        self.setWindowTitle("Xử lý dữ liệu trước khi nhập · Demo A1–B3")
        self.resize(1380, 850)
        self.setMinimumSize(1050, 670)
        outer = QVBoxLayout(self)
        outer.addWidget(label("Xử lý dữ liệu trước khi nhập", "title"))
        outer.addWidget(label("DEMO · 7 tình huống mẫu · Không đọc hoặc ghi JSON, Excel hay database thật", "badge"))
        self.summary = label("", "hint")
        outer.addWidget(self.summary)

        body = QHBoxLayout()
        left = QFrame()
        left.setObjectName("card")
        left_layout = QVBoxLayout(left)
        self.search = QLineEdit()
        self.search.setPlaceholderText("Tìm A1–B3, Container, chứng từ...")
        self.search.textChanged.connect(self.refresh)
        left_layout.addWidget(self.search)
        filters = QHBoxLayout()
        self.status_filter = QComboBox()
        self.status_filter.addItems(["Tất cả trạng thái", "Chưa xử lý", "Đã xử lý", "Để sau"])
        self.status_filter.currentIndexChanged.connect(self.refresh)
        self.type_filter = QComboBox()
        self.type_filter.addItems(["Mọi vấn đề", "Nhiều Bill", "Nhiều SQT", "Thiếu Ngày đóng", "Không tìm thấy SQT", "Dữ liệu khác Excel"])
        self.type_filter.currentIndexChanged.connect(self.refresh)
        filters.addWidget(self.status_filter)
        filters.addWidget(self.type_filter)
        left_layout.addLayout(filters)
        controls = QHBoxLayout()
        controls.addWidget(button("Chọn tất cả đang hiện", self.select_visible))
        controls.addWidget(button("Bỏ chọn", self.clear_selected))
        left_layout.addLayout(controls)
        self.table = QTableWidget(0, 7)
        self.table.setHorizontalHeaderLabels(["Chọn", "Luồng", "Chứng từ", "Số chứng từ", "Container", "Vấn đề", "Trạng thái"])
        self.table.setSelectionBehavior(QTableWidget.SelectRows)
        self.table.setSelectionMode(QTableWidget.SingleSelection)
        self.table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.Interactive)
        self.table.horizontalHeader().setStretchLastSection(True)
        for col, width in enumerate((48, 60, 116, 100, 124, 155)):
            self.table.setColumnWidth(col, width)
        self.table.itemChanged.connect(self.checked)
        self.table.itemSelectionChanged.connect(self.row_changed)
        left_layout.addWidget(self.table)
        self.shown = label("", "hint")
        left_layout.addWidget(self.shown)
        self.detail = QFrame()
        self.detail.setObjectName("card")
        self.detail_layout = QVBoxLayout(self.detail)
        body.addWidget(left, 58)
        body.addWidget(self.detail, 42)
        outer.addLayout(body, 1)

        bulk = QHBoxLayout()
        bulk.addWidget(label("Đánh dấu ô Chọn để thử xử lý hàng loạt.", "hint"))
        bulk.addStretch()
        bulk.addWidget(button("Sửa nhanh", self.bulk_edit))
        bulk.addWidget(button("Chọn ứng viên", self.bulk_candidate))
        bulk.addWidget(button("Chọn dữ liệu", self.bulk_fields))
        bulk.addWidget(button("Để xử lý sau", self.defer_selected))
        bulk.addWidget(button("Hoàn tác", self.undo))
        outer.addLayout(bulk)
        footer = QHBoxLayout()
        self.progress = label("", "hint")
        footer.addWidget(self.progress)
        footer.addStretch()
        footer.addWidget(button("Đặt lại demo", self.reset))
        footer.addWidget(button("Lưu và đóng", self.save_demo))
        self.continue_button = button("Tiếp tục nhập", self.continue_flow, True)
        footer.addWidget(self.continue_button)
        outer.addLayout(footer)
        self.refresh()

    def issue(self, code: str) -> DemoIssue:
        return next(item for item in self.issues if item.code == code)

    def refresh(self, *_args):
        pending = sum(i.status == "Chưa xử lý" for i in self.issues)
        done = sum(i.status == "Đã xử lý" for i in self.issues)
        skipped = sum(i.status == "Để sau" for i in self.issues)
        self.summary.setText(f"7 vấn đề trong lô     {done} đã xử lý     {skipped} để sau     {pending} chưa xử lý")
        self.progress.setText(f"Đã quyết định {done + skipped}/7 · Còn {pending} vấn đề")
        self.continue_button.setEnabled(pending == 0)
        query = self.search.text().strip().casefold()
        self.visible = [i.code for i in self.issues if
                        (self.status_filter.currentIndex() == 0 or i.status == self.status_filter.currentText())
                        and (self.type_filter.currentIndex() == 0 or i.title == self.type_filter.currentText())
                        and (not query or query in " ".join((i.code, i.doc, i.reference, i.container, i.title)).casefold())]
        self.table.blockSignals(True)
        self.table.setRowCount(len(self.visible))
        for row, code in enumerate(self.visible):
            issue = self.issue(code)
            check = QTableWidgetItem()
            check.setFlags(Qt.ItemIsEnabled | Qt.ItemIsUserCheckable)
            check.setCheckState(Qt.Checked if code in self.selected else Qt.Unchecked)
            self.table.setItem(row, 0, check)
            for col, value in enumerate((issue.code, issue.doc, issue.reference, issue.container, issue.title, issue.status), 1):
                self.table.setItem(row, col, QTableWidgetItem(value))
        if self.visible:
            self.table.selectRow(self.visible.index(self.active_code) if self.active_code in self.visible else 0)
            self.active_code = self.visible[self.table.currentRow()]
        self.table.blockSignals(False)
        self.shown.setText(f"Đang hiện {len(self.visible)}/7 vấn đề")
        self.show_detail()

    def row_changed(self):
        row = self.table.currentRow()
        if 0 <= row < len(self.visible):
            self.active_code = self.visible[row]
            self.reviewing = False
            self.continue_button.setText("Tiếp tục nhập")
            self.show_detail()

    def checked(self, item: QTableWidgetItem):
        if item.column() != 0 or item.row() >= len(self.visible):
            return
        code = self.visible[item.row()]
        if item.checkState() == Qt.Checked:
            self.selected.add(code)
        else:
            self.selected.discard(code)

    def select_visible(self):
        self.selected.update(self.visible)
        self.refresh()

    def clear_selected(self):
        self.selected.clear()
        self.refresh()

    def show_detail(self):
        clear_layout(self.detail_layout)
        if not self.visible:
            self.detail_layout.addWidget(label("Không có dòng phù hợp với bộ lọc.", "hint"))
            return
        issue = self.issue(self.active_code)
        self.detail_layout.addWidget(label(f"{issue.code}. {issue.title}", "heading"))
        self.detail_layout.addWidget(label(issue.reason, "hint"))
        card = QFrame()
        card.setObjectName("card")
        card_layout = QVBoxLayout(card)
        card_layout.addWidget(label("DỮ LIỆU BÓC TÁCH", "hint"))
        card_layout.addWidget(label(f"{issue.doc} · {issue.reference}\nNgày ghép: {issue.date or '—'}\nContainer: {issue.container or '—'}"))
        actions = QHBoxLayout()
        actions.addWidget(button("Chỉnh sửa dòng", lambda: self.edit_one(issue)))
        actions.addWidget(button("Để xử lý sau", lambda: self.defer_one(issue)))
        card_layout.addLayout(actions)
        self.detail_layout.addWidget(card)
        if issue.status != "Chưa xử lý":
            self.detail_layout.addWidget(label(f"Trạng thái: {issue.status}. {issue.decision}", "hint"))
            self.detail_layout.addWidget(button("Đặt lại lựa chọn", lambda: self.reset_one(issue)))
        if issue.candidates:
            self.candidate_detail(issue)
        elif issue.differences:
            self.field_detail(issue)
        else:
            self.detail_layout.addWidget(label("Sửa ngày/Container để ghép lại, hoặc để xử lý sau.", "hint"))
        if issue.next_code:
            self.detail_layout.addWidget(label(f"Luồng mẫu tiếp theo sau bước này: {issue.next_code}", "hint"))
        self.detail_layout.addStretch()

    def candidate_detail(self, issue: DemoIssue):
        self.detail_layout.addWidget(label("Chọn Bill phù hợp" if issue.code == "A1" else "Chọn số quyết toán phù hợp", "heading"))
        for code, description in issue.candidates:
            row = QHBoxLayout()
            radio = QRadioButton(description)
            radio.setChecked(issue.decision == f"Đã chọn {code}")
            radio.toggled.connect(lambda checked, choice=code: self.choose_candidate(issue, choice) if checked else None)
            row.addWidget(radio)
            if issue.code == "A1":
                row.addWidget(button("Sửa Bill này", lambda _checked=False, value=code: self.edit_bill(value)))
            self.detail_layout.addLayout(row)

    def field_detail(self, issue: DemoIssue):
        self.detail_layout.addWidget(label("Chọn dữ liệu sử dụng cho từng trường", "heading"))
        table = QTableWidget(len(issue.differences), 5)
        table.setHorizontalHeaderLabels(["Trường", "Excel", "JSON", "Lựa chọn", "Giá trị nhập"])
        table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeToContents)
        table.horizontalHeader().setStretchLastSection(True)
        self.field_widgets = []
        for row, diff in enumerate(issue.differences):
            for col, value in enumerate((diff.name, diff.current, diff.incoming)):
                table.setItem(row, col, QTableWidgetItem(value))
            combo = QComboBox()
            combo.addItems(["Giữ dữ liệu Excel", "Dùng dữ liệu mới", "Nhập thủ công"])
            editor = QLineEdit()
            editor.setEnabled(False)
            combo.currentIndexChanged.connect(lambda index, target=editor: target.setEnabled(index == 2))
            table.setCellWidget(row, 3, combo)
            table.setCellWidget(row, 4, editor)
            self.field_widgets.append((diff, combo, editor))
        self.detail_layout.addWidget(table)
        self.detail_layout.addWidget(button("Áp dụng lựa chọn", lambda: self.apply_fields(issue), True))

    def snapshot(self):
        self.undo_state = (copy.deepcopy(self.issues), set(self.selected))

    def choose_candidate(self, issue: DemoIssue, candidate: str):
        if issue.status == "Đã xử lý" and issue.decision == f"Đã chọn {candidate}":
            return
        issue.status = "Đã xử lý"
        issue.decision = f"Đã chọn {candidate}"
        self.after_resolution(issue)

    def after_resolution(self, issue: DemoIssue):
        if issue.next_code:
            self.active_code = issue.next_code
        self.refresh()

    def apply_fields(self, issue: DemoIssue):
        answers = []
        for diff, combo, editor in self.field_widgets:
            if combo.currentIndex() == 2 and not editor.text().strip():
                QMessageBox.warning(self, "Chưa nhập giá trị", f"Hãy nhập giá trị cho trường {diff.name}.")
                return
            value = editor.text().strip() if combo.currentIndex() == 2 else combo.currentText()
            answers.append(f"{diff.name}: {value}")
        issue.status = "Đã xử lý"
        issue.decision = "; ".join(answers)
        self.refresh()

    def edit_one(self, issue: DemoIssue):
        if EditRowDialog(issue, self).exec() == QDialog.Accepted:
            if issue.code in ("A3", "B1") and issue.date and issue.container:
                issue.status = "Đã xử lý"
                issue.decision = "Đã sửa thông tin để ghép lại (mô phỏng)"
                self.after_resolution(issue)
            else:
                self.refresh()

    def edit_bill(self, bill: str):
        QMessageBox.information(self, "Sửa Bill này · Demo", f"Trong PM thật, nút này mở form sửa dòng JSON của {bill}.\nDemo dùng nút 'Chỉnh sửa dòng' để xem form tương ứng.")

    def defer_one(self, issue: DemoIssue):
        issue.status = "Để sau"
        issue.decision = "Giữ trong danh sách chờ"
        self.refresh()

    def reset_one(self, issue: DemoIssue):
        issue.status = "Chưa xử lý"
        issue.decision = ""
        self.refresh()

    def chosen_pending(self, predicate):
        return [i for i in self.issues if i.code in self.selected and i.status == "Chưa xử lý" and predicate(i)]

    def bulk_edit(self):
        issues = self.chosen_pending(lambda i: i.code in ("A3", "B1"))
        if not issues:
            QMessageBox.information(self, "Sửa nhanh", "Hãy đánh dấu A3 hoặc B1 đang chờ xử lý.")
            return
        self.snapshot()
        if BulkEditDialog(issues, self).exec() == QDialog.Accepted:
            self.refresh()

    def bulk_candidate(self):
        issues = self.chosen_pending(lambda i: bool(i.candidates))
        if not issues:
            QMessageBox.information(self, "Chọn ứng viên", "Hãy đánh dấu A1, A2 hoặc B2 đang chờ xử lý.")
            return
        common_codes = set(code for code, _ in issues[0].candidates)
        for issue in issues[1:]:
            common_codes &= {code for code, _ in issue.candidates}
        common = [(code, text) for code, text in issues[0].candidates if code in common_codes]
        if not common:
            QMessageBox.information(self, "Không có lựa chọn chung", "Các dòng này không có cùng một ứng viên. Hãy chọn từng dòng hoặc một nhóm nhỏ hơn.")
            return
        self.snapshot()
        if BulkCandidateDialog(issues, common, self).exec() == QDialog.Accepted:
            self.refresh()

    def bulk_fields(self):
        issues = self.chosen_pending(lambda i: bool(i.differences))
        if not issues:
            QMessageBox.information(self, "Chọn dữ liệu", "Hãy đánh dấu A4 hoặc B3 đang chờ xử lý.")
            return
        self.snapshot()
        if BulkFieldDialog(issues, self).exec() == QDialog.Accepted:
            self.refresh()

    def defer_selected(self):
        issues = self.chosen_pending(lambda i: True)
        if not issues:
            QMessageBox.information(self, "Để xử lý sau", "Hãy đánh dấu ít nhất một vấn đề chưa xử lý.")
            return
        self.snapshot()
        for issue in issues:
            issue.status = "Để sau"
            issue.decision = "Giữ trong danh sách chờ"
        self.selected.clear()
        self.refresh()

    def undo(self):
        if not self.undo_state:
            QMessageBox.information(self, "Hoàn tác", "Chưa có thao tác hàng loạt để hoàn tác.")
            return
        self.issues, self.selected = self.undo_state
        self.undo_state = None
        self.refresh()

    def reset(self):
        self.issues = sample_issues()
        self.selected.clear()
        self.active_code = "A1"
        self.undo_state = None
        self.reviewing = False
        self.status_filter.setCurrentIndex(0)
        self.type_filter.setCurrentIndex(0)
        self.search.clear()
        self.refresh()

    def save_demo(self):
        QMessageBox.information(self, "Lưu và đóng · Demo", "Trong PM thật, lựa chọn sẽ được lưu để tiếp tục sau.\nDemo này chỉ giữ trạng thái trong bộ nhớ khi cửa sổ còn mở.")

    def continue_flow(self):
        if any(i.status == "Chưa xử lý" for i in self.issues):
            return
        if not self.reviewing:
            self.reviewing = True
            clear_layout(self.detail_layout)
            self.detail_layout.addWidget(label("Xác nhận nhập dữ liệu", "heading"))
            skipped = sum(i.status == "Để sau" for i in self.issues)
            done = sum(i.status == "Đã xử lý" for i in self.issues)
            self.detail_layout.addWidget(label(f"Dòng đã xử lý: {done}\nDòng để xử lý sau: {skipped}\n\nTrong PM thật, Excel chỉ được ghi sau khi bấm 'Xác nhận nhập'."))
            self.detail_layout.addStretch()
            self.continue_button.setText("Xác nhận nhập")
            return
        skipped = sum(i.status == "Để sau" for i in self.issues)
        QMessageBox.information(self, "Nhập dữ liệu hoàn tất · Demo", f"Đã mô phỏng bước xác nhận nhập.\nCòn {skipped} dòng trong danh sách chờ.\nKhông có file thật nào bị thay đổi.")
        self.reviewing = False
        self.continue_button.setText("Tiếp tục nhập")
        self.show_detail()


def main() -> int:
    app = QApplication(sys.argv)
    font_file = Path("C:/Windows/Fonts/segoeui.ttf")
    if font_file.is_file():
        QFontDatabase.addApplicationFont(str(font_file))
    app.setStyleSheet(STYLE)
    dialog = ResolutionDemo()
    dialog.show()
    return app.exec()


if __name__ == "__main__":
    raise SystemExit(main())
