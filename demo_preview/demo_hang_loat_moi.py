"""Demo thiết kế mới cho xử lý nhiều xung đột. Chạy: python demo_hang_loat_moi.py.

Dữ liệu hoàn toàn giả lập và chỉ nằm trong bộ nhớ; không mở file PM thật.
"""

from __future__ import annotations

import copy
import sys
from datetime import datetime
from pathlib import Path

from PySide6.QtCore import Qt
from PySide6.QtGui import QFontDatabase, QKeySequence
from PySide6.QtWidgets import (
    QApplication, QComboBox, QDialog, QFrame, QHBoxLayout, QHeaderView,
    QLabel, QLineEdit, QMessageBox, QPushButton, QTableWidget,
    QTableWidgetItem, QTabWidget, QVBoxLayout, QWidget,
)


STYLE = """
QWidget { font-family: 'Segoe UI', Arial; font-size: 10pt; color: #1D304A; }
QDialog { background: #F4F7FB; }
QFrame#card { background: white; border: 1px solid #DCE5F0; border-radius: 10px; }
QLabel#title { font-size: 18pt; font-weight: 700; color: #102541; }
QLabel#heading { font-size: 12pt; font-weight: 700; color: #102541; }
QLabel#hint { color: #5E718B; }
QLabel#notice { background: #EAF3FF; color: #245A9D; padding: 9px; border-radius: 7px; }
QPushButton { background: white; border: 1px solid #C6D3E2; border-radius: 7px; padding: 7px 12px; }
QPushButton:hover { background: #EDF3FB; }
QPushButton:disabled { color: #98A7B8; background: #F2F4F7; }
QPushButton[primary='true'] { color: white; background: #285DB7; border-color: #285DB7; font-weight: 700; }
QLineEdit, QComboBox { background: white; border: 1px solid #C7D3E2; border-radius: 6px; padding: 5px; }
QTableWidget { background: white; gridline-color: #E4EAF2; selection-background-color: #DDEBFF; selection-color: #172C4A; }
QHeaderView::section { background: #EDF2F8; padding: 8px; border: 0; border-right: 1px solid #DCE5F0; }
QTabWidget::pane { border: 1px solid #DCE5F0; background: white; }
QTabBar::tab { background: #E9EFF7; padding: 10px 16px; margin-right: 3px; }
QTabBar::tab:selected { background: white; color: #285DB7; font-weight: 700; }
"""


def text_label(value: str, kind: str = "") -> QLabel:
    result = QLabel(value)
    result.setWordWrap(True)
    if kind:
        result.setObjectName(kind)
    return result


def button(value: str, slot, primary: bool = False) -> QPushButton:
    result = QPushButton(value)
    if primary:
        result.setProperty("primary", "true")
    result.clicked.connect(slot)
    return result


def item(value: str, editable: bool = False) -> QTableWidgetItem:
    result = QTableWidgetItem(value)
    if not editable:
        result.setFlags(result.flags() & ~Qt.ItemIsEditable)
    return result


class PasteTable(QTableWidget):
    """Cho phép dán nhiều ô Ngày/Container từ Excel vào bảng sửa nhanh."""

    def keyPressEvent(self, event):  # noqa: N802 - Qt API
        if event.matches(QKeySequence.Paste):
            values = QApplication.clipboard().text().replace("\r\n", "\n").strip("\n")
            start_row, start_col = self.currentRow(), self.currentColumn()
            if start_row >= 0 and start_col in (3, 4):
                for row_offset, line in enumerate(values.split("\n")):
                    for col_offset, value in enumerate(line.split("\t")):
                        row, col = start_row + row_offset, start_col + col_offset
                        if row < self.rowCount() and col in (3, 4):
                            self.setItem(row, col, item(value.strip(), editable=True))
                return
        super().keyPressEvent(event)


class BulkWorkspaceDemo(QDialog):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Thiết kế mới · Xử lý xung đột theo nhóm công việc")
        self.resize(1390, 850)
        self.setMinimumSize(1080, 700)
        self.edit_rows = [
            dict(code="A3", doc="Bill BL-003", source="Ngày đóng", date="", container="KLMN7654321",
                 expected_date="12/06/2026", expected_container="KLMN7654321", status="Chưa xử lý", decision=""),
            dict(code="B1", doc="Khoản chi HĐ 88", source="Ngày chạy (đang dùng)", date="20/06/2026", container="EFGH1234567",
                 expected_date="03/06/2026", expected_container="EFGH1234567", status="Chưa xử lý", decision=""),
            dict(code="B1", doc="Khoản chi HĐ 90", source="Ngày đóng (đang dùng)", date="05/06/2026", container="MNOP3334445",
                 expected_date="05/06/2026", expected_container="MNOP3334446", status="Chưa xử lý", decision=""),
        ]
        for data in self.edit_rows:
            data["last_checked"] = (data["date"], data["container"])
        self.choice_rows = [
            dict(code="A1", doc="Phiếu cân SC-014", container="ABCD1234567", date="12/06/2026",
                 options=[("BL-001", "Bill BL-001 · tàu AN PHÚ"), ("BL-002", "Bill BL-002 · tàu MINH LONG")],
                 status="Chưa xử lý", decision=""),
            dict(code="A1", doc="Phiếu cân SC-015", container="QRST1234567", date="13/06/2026",
                 options=[("BL-101", "Bill BL-101 · tàu HẢI NAM"), ("BL-102", "Bill BL-102 · tàu ĐẠI DƯƠNG")],
                 status="Chưa xử lý", decision=""),
            dict(code="A2", doc="Bill BL-010", container="ABCD1234567", date="12/06/2026",
                 options=[("A-200", "Tháng 6 · dòng 18 · SQT 200"), ("A-201", "Tháng 6 · dòng 26 · SQT 201")],
                 status="Chưa xử lý", decision=""),
            dict(code="B2", doc="Khoản chi HĐ 89", container="EFGH1234567", date="03/06/2026",
                 options=[("B-300", "Tháng 6 · dòng 12 · SQT 300"), ("B-301", "Tháng 6 · dòng 19 · SQT 301")],
                 status="Chưa xử lý", decision=""),
            dict(code="B2", doc="Khoản chi HĐ 91", container="MNOP3334446", date="05/06/2026",
                 options=[("C-400", "Tháng 6 · dòng 30 · SQT 400"), ("C-401", "Tháng 6 · dòng 32 · SQT 401")],
                 status="Chưa xử lý", decision=""),
        ]
        self.diff_rows = [
            dict(code="A4", doc="Phiếu cân SC-014", field="Số tấn", old="28", new="30", status="Chưa xử lý", decision=""),
            dict(code="A4", doc="Phiếu cân SC-014", field="Nơi đóng", old="Kho A", new="Kho B", status="Chưa xử lý", decision=""),
            dict(code="A4", doc="Phiếu cân SC-015", field="Nơi đóng", old="Kho A", new="Kho C", status="Chưa xử lý", decision=""),
            dict(code="B3", doc="Khoản chi HĐ 89", field="Đơn giá", old="100.000", new="120.000", status="Chưa xử lý", decision=""),
            dict(code="B3", doc="Khoản chi HĐ 89", field="Tổng tiền", old="2.800.000", new="3.360.000", status="Chưa xử lý", decision=""),
            dict(code="B3", doc="Khoản chi HĐ 91", field="Đơn giá", old="95.000", new="110.000", status="Chưa xử lý", decision=""),
        ]
        self.initial_rows = copy.deepcopy((self.edit_rows, self.choice_rows, self.diff_rows))
        self._build()
        self.refresh_counts()

    def _build(self):
        outer = QVBoxLayout(self)
        outer.addWidget(text_label("Xử lý dữ liệu trước khi nhập", "title"))
        outer.addWidget(text_label("DEMO PHƯƠNG ÁN MỚI · Sửa và chọn trực tiếp trong bảng · Không dùng dữ liệu thật", "notice"))
        self.summary = text_label("", "hint")
        outer.addWidget(self.summary)
        self.tabs = QTabWidget()
        self.tabs.addTab(self._build_edit(), "")
        self.tabs.addTab(self._build_choice(), "")
        self.tabs.addTab(self._build_diff(), "")
        self.tabs.addTab(self._build_review(), "Xem trước nhập")
        self.tabs.currentChanged.connect(self._tab_changed)
        outer.addWidget(self.tabs, 1)
        footer = QHBoxLayout()
        self.feedback = text_label("Mỗi thao tác chỉ thay đổi dữ liệu mẫu trong cửa sổ này.", "hint")
        footer.addWidget(self.feedback, 1)
        footer.addWidget(button("Đặt lại demo", self.reset))
        footer.addWidget(button("Xem trước nhập", lambda: self.tabs.setCurrentIndex(3), True))
        outer.addLayout(footer)

    def _page(self, heading: str, explanation: str):
        page = QWidget()
        layout = QVBoxLayout(page)
        layout.setContentsMargins(18, 18, 18, 16)
        layout.addWidget(text_label(heading, "heading"))
        layout.addWidget(text_label(explanation, "hint"))
        return page, layout

    def _build_edit(self):
        page, layout = self._page(
            "Sửa thông tin để ghép lại",
            "Sửa trực tiếp ngày và Container; có thể dán nhiều ô từ Excel. Cột ‘Ngày đang dùng’ cho biết PM đang ghép bằng trường nào.",
        )
        self.edit_table = PasteTable(len(self.edit_rows), 6)
        self.edit_table.setHorizontalHeaderLabels(["Luồng", "Chứng từ", "Ngày đang dùng", "Ngày", "Container", "Kết quả"])
        self.edit_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.edit_table.verticalHeader().setVisible(False)
        self.edit_table.setSelectionBehavior(QTableWidget.SelectItems)
        for row, data in enumerate(self.edit_rows):
            for col, value in enumerate((data["code"], data["doc"], data["source"], data["date"], data["container"], data["status"])):
                self.edit_table.setItem(row, col, item(value, editable=col in (3, 4)))
        layout.addWidget(self.edit_table, 1)
        actions = QHBoxLayout()
        actions.addWidget(button("Điền ví dụ để thử", self.fill_edit_examples))
        actions.addWidget(button("Để dòng đang chọn xử lý sau", lambda: self.defer_row("edit")))
        actions.addStretch()
        actions.addWidget(button("Kiểm tra lại các dòng đã sửa", self.apply_edits, True))
        layout.addLayout(actions)
        return page

    def _build_choice(self):
        page, layout = self._page(
            "Chọn Bill hoặc SQT cho từng dòng",
            "Mỗi chứng từ có lựa chọn riêng. Chọn ngay trên dòng, sau đó áp dụng tất cả lựa chọn đã điền một lần.",
        )
        self.choice_table = QTableWidget(len(self.choice_rows), 6)
        self.choice_table.setHorizontalHeaderLabels(["Luồng", "Chứng từ", "Ngày", "Container", "Chọn Bill / SQT", "Kết quả"])
        self.choice_table.horizontalHeader().setSectionResizeMode(QHeaderView.Interactive)
        self.choice_table.horizontalHeader().setStretchLastSection(True)
        for col, width in enumerate((66, 210, 105, 150, 415)):
            self.choice_table.setColumnWidth(col, width)
        self.choice_table.verticalHeader().setVisible(False)
        self.choice_table.setSelectionBehavior(QTableWidget.SelectRows)
        self.choice_table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.choice_boxes = []
        for row, data in enumerate(self.choice_rows):
            for col, value in enumerate((data["code"], data["doc"], data["date"], data["container"])):
                self.choice_table.setItem(row, col, item(value))
            box = QComboBox()
            box.addItem("Chọn kết quả cho dòng này…", "")
            for key, label in data["options"]:
                box.addItem(label, key)
            self.choice_table.setCellWidget(row, 4, box)
            self.choice_boxes.append(box)
            self.choice_table.setItem(row, 5, item(data["status"]))
        layout.addWidget(self.choice_table, 1)
        self.choice_help = text_label("Chọn một dòng để xem lưu ý. Với A1, Bill không được chọn sẽ ở lại danh sách chờ.", "notice")
        self.choice_table.itemSelectionChanged.connect(self.show_choice_help)
        layout.addWidget(self.choice_help)
        actions = QHBoxLayout()
        actions.addWidget(button("Điền ví dụ để thử", self.fill_choice_examples))
        actions.addWidget(button("Để dòng đang chọn xử lý sau", lambda: self.defer_row("choice")))
        actions.addStretch()
        self.choice_apply = button("Áp dụng 0 lựa chọn", self.apply_choices, True)
        actions.addWidget(self.choice_apply)
        layout.addLayout(actions)
        for box in self.choice_boxes:
            box.currentIndexChanged.connect(self.update_choice_button)
        self.update_choice_button()
        return page

    def _build_diff(self):
        page, layout = self._page(
            "So sánh dữ liệu Excel và JSON",
            "Đặt một quy tắc cho trường đang lọc, sau đó sửa ngoại lệ ở từng dòng trước khi áp dụng.",
        )
        rule = QHBoxLayout()
        rule.addWidget(text_label("Trường:"))
        self.field_filter = QComboBox()
        self.field_filter.addItem("Tất cả trường")
        self.field_filter.addItems(sorted({row["field"] for row in self.diff_rows}))
        self.field_filter.currentIndexChanged.connect(self.filter_diffs)
        rule.addWidget(self.field_filter)
        rule.addWidget(text_label("Quy tắc cho các dòng đang hiện:"))
        self.rule_box = QComboBox()
        self.rule_box.addItems(["Chọn cách xử lý", "Giữ dữ liệu Excel", "Dùng dữ liệu JSON"])
        rule.addWidget(self.rule_box)
        rule.addWidget(button("Điền quy tắc", self.apply_rule))
        rule.addStretch()
        layout.addLayout(rule)
        self.diff_table = QTableWidget(len(self.diff_rows), 7)
        self.diff_table.setHorizontalHeaderLabels(["Luồng", "Chứng từ", "Trường", "Excel", "JSON", "Lựa chọn / giá trị nhập", "Kết quả"])
        self.diff_table.horizontalHeader().setSectionResizeMode(QHeaderView.Interactive)
        self.diff_table.horizontalHeader().setStretchLastSection(True)
        for col, width in enumerate((66, 185, 110, 125, 125, 300)):
            self.diff_table.setColumnWidth(col, width)
        self.diff_table.verticalHeader().setVisible(False)
        self.diff_table.setSelectionBehavior(QTableWidget.SelectRows)
        self.diff_table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.diff_boxes = []
        self.diff_custom = []
        for row, data in enumerate(self.diff_rows):
            for col, value in enumerate((data["code"], data["doc"], data["field"], data["old"], data["new"])):
                self.diff_table.setItem(row, col, item(value))
            cell = QWidget()
            cell_layout = QHBoxLayout(cell)
            cell_layout.setContentsMargins(2, 1, 2, 1)
            box = QComboBox()
            box.addItems(["Chưa chọn", "Giữ Excel", "Dùng JSON", "Nhập tay"])
            custom = QLineEdit()
            custom.setPlaceholderText("Giá trị")
            custom.setEnabled(False)
            box.currentIndexChanged.connect(lambda index, target=custom: target.setEnabled(index == 3))
            cell_layout.addWidget(box, 2)
            cell_layout.addWidget(custom, 1)
            self.diff_table.setCellWidget(row, 5, cell)
            self.diff_boxes.append(box)
            self.diff_custom.append(custom)
            self.diff_table.setItem(row, 6, item(data["status"]))
        layout.addWidget(self.diff_table, 1)
        actions = QHBoxLayout()
        self.diff_scope = text_label("Đang hiện 6/6 trường", "hint")
        actions.addWidget(self.diff_scope)
        actions.addWidget(button("Để dòng đang chọn xử lý sau", lambda: self.defer_row("diff")))
        actions.addStretch()
        actions.addWidget(button("Áp dụng các lựa chọn đã điền", self.apply_diffs, True))
        layout.addLayout(actions)
        return page

    def _build_review(self):
        page, layout = self._page(
            "Xem trước trước khi ghi Excel",
            "Kiểm tra các quyết định và những dòng được để sau. Demo này không ghi vào Excel.",
        )
        self.review_summary = text_label("", "notice")
        layout.addWidget(self.review_summary)
        self.review_table = QTableWidget(0, 4)
        self.review_table.setHorizontalHeaderLabels(["Nhóm việc", "Chứng từ", "Trường / ngày", "Quyết định"])
        self.review_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.review_table.setEditTriggers(QTableWidget.NoEditTriggers)
        layout.addWidget(self.review_table, 1)
        actions = QHBoxLayout()
        actions.addStretch()
        self.finish_button = button("Mô phỏng xác nhận nhập", self.finish, True)
        actions.addWidget(self.finish_button)
        layout.addLayout(actions)
        return page

    def refresh_counts(self):
        groups = (self.edit_rows, self.choice_rows, self.diff_rows)
        pending = [sum(row["status"] == "Chưa xử lý" for row in group) for group in groups]
        done = sum(row["status"] == "Đã xử lý" for group in groups for row in group)
        skipped = sum(row["status"] == "Để sau" for group in groups for row in group)
        self.summary.setText(f"{done} mục đã xử lý    ·    {sum(pending)} mục còn chờ    ·    {skipped} mục để sau")
        self.tabs.setTabText(0, f"1. Sửa thông tin ({pending[0]})")
        self.tabs.setTabText(1, f"2. Chọn Bill/SQT ({pending[1]})")
        self.tabs.setTabText(2, f"3. So sánh giá trị ({pending[2]})")
        self.finish_button.setEnabled(sum(pending) == 0)
        if self.tabs.currentIndex() == 3:
            self.refresh_review()

    def set_feedback(self, message: str):
        self.feedback.setText(message)
        self.refresh_counts()

    def fill_edit_examples(self):
        for row, data in enumerate(self.edit_rows):
            if data["status"] == "Chưa xử lý":
                self.edit_table.item(row, 3).setText(data["expected_date"])
                self.edit_table.item(row, 4).setText(data["expected_container"])
        self.set_feedback("Đã điền giá trị mẫu. Bấm ‘Kiểm tra lại các dòng đã sửa’ để xem kết quả.")

    def apply_edits(self):
        changed = []
        for row, data in enumerate(self.edit_rows):
            if data["status"] == "Để sau":
                continue
            date = self.edit_table.item(row, 3).text().strip()
            container = self.edit_table.item(row, 4).text().strip().upper()
            if (date, container) == data["last_checked"]:
                continue
            if date:
                try:
                    datetime.strptime(date, "%d/%m/%Y")
                except ValueError:
                    QMessageBox.warning(self, "Ngày chưa hợp lệ", f"Dòng {row + 1}: hãy nhập dạng dd/mm/yyyy.")
                    return
            changed.append((row, data, date, container))
        matched = failed = 0
        for row, data, date, container in changed:
            data["last_checked"] = (date, container)
            if (date, container) == (data["expected_date"], data["expected_container"]):
                data["status"] = "Đã xử lý"
                data["decision"] = f"{data['source']}: {date} · {container}"
                matched += 1
            else:
                data["status"] = "Chưa xử lý"
                data["decision"] = ""
                failed += 1
            self.edit_table.item(row, 5).setText("Đã ghép" if data["status"] == "Đã xử lý" else "Còn chưa khớp")
        self.set_feedback(f"Kiểm tra {len(changed)} dòng đã sửa: {matched} đã ghép, {failed} còn chưa khớp.")

    def fill_choice_examples(self):
        for row, data in enumerate(self.choice_rows):
            if data["status"] == "Chưa xử lý":
                self.choice_boxes[row].setCurrentIndex(1)
        self.set_feedback("Mỗi dòng đã có một lựa chọn mẫu riêng. Có thể đổi từng dòng trước khi áp dụng.")

    def update_choice_button(self, *_args):
        count = sum(data["status"] == "Chưa xử lý" and box.currentIndex() > 0
                    for data, box in zip(self.choice_rows, self.choice_boxes))
        self.choice_apply.setText(f"Áp dụng {count} lựa chọn")
        self.choice_apply.setEnabled(count > 0)

    def show_choice_help(self):
        row = self.choice_table.currentRow()
        if row < 0:
            return
        data = self.choice_rows[row]
        choices = "  |  ".join(value for _, value in data["options"])
        ending = "Bill không được chọn sẽ được để xử lý sau." if data["code"] == "A1" else "Có thể phát sinh bước so sánh dữ liệu sau khi chọn."
        self.choice_help.setText(f"{data['doc']} · {choices}\n{ending}")

    def apply_choices(self):
        count = 0
        for row, data in enumerate(self.choice_rows):
            box = self.choice_boxes[row]
            if data["status"] != "Chưa xử lý" or box.currentIndex() == 0:
                continue
            data["status"] = "Đã xử lý"
            data["decision"] = box.currentText()
            self.choice_table.item(row, 5).setText("Đã chọn")
            count += 1
        self.update_choice_button()
        self.set_feedback(f"Đã chọn kết quả cho {count} dòng. Các dòng chưa chọn vẫn nằm trong danh sách.")

    def filter_diffs(self):
        wanted = self.field_filter.currentText()
        shown = 0
        for row, data in enumerate(self.diff_rows):
            visible = wanted == "Tất cả trường" or wanted == data["field"]
            self.diff_table.setRowHidden(row, not visible)
            shown += visible
        self.diff_scope.setText(f"Đang hiện {shown}/{len(self.diff_rows)} trường")

    def apply_rule(self):
        index = self.rule_box.currentIndex()
        if index == 0:
            QMessageBox.information(self, "Chưa chọn quy tắc", "Hãy chọn Giữ Excel hoặc Dùng JSON.")
            return
        count = 0
        for row, data in enumerate(self.diff_rows):
            if not self.diff_table.isRowHidden(row) and data["status"] == "Chưa xử lý":
                self.diff_boxes[row].setCurrentIndex(index)
                count += 1
        self.set_feedback(f"Đã điền quy tắc vào {count} trường đang hiện. Kiểm tra ngoại lệ rồi bấm Áp dụng.")

    def apply_diffs(self):
        decisions = []
        for row, data in enumerate(self.diff_rows):
            if data["status"] != "Chưa xử lý" or self.diff_boxes[row].currentIndex() == 0:
                continue
            box = self.diff_boxes[row]
            custom = self.diff_custom[row].text().strip()
            if box.currentIndex() == 3 and not custom:
                QMessageBox.warning(self, "Thiếu giá trị", f"Hãy nhập giá trị cho {data['doc']} · {data['field']}.")
                return
            decisions.append((row, custom if box.currentIndex() == 3 else box.currentText()))
        for row, decision in decisions:
            self.diff_rows[row]["status"] = "Đã xử lý"
            self.diff_rows[row]["decision"] = decision
            self.diff_table.item(row, 6).setText("Đã chọn")
        self.set_feedback(f"Đã quyết định {len(decisions)} trường. Các trường chưa chọn vẫn nằm trong danh sách.")

    def defer_row(self, group: str):
        table, rows = {"edit": (self.edit_table, self.edit_rows),
                       "choice": (self.choice_table, self.choice_rows),
                       "diff": (self.diff_table, self.diff_rows)}[group]
        row = table.currentRow()
        if row < 0:
            QMessageBox.information(self, "Chưa chọn dòng", "Hãy chọn một dòng trong bảng.")
            return
        rows[row]["status"] = "Để sau"
        rows[row]["decision"] = "Để xử lý sau"
        status_col = {"edit": 5, "choice": 5, "diff": 6}[group]
        table.item(row, status_col).setText("Để sau")
        self.update_choice_button()
        self.set_feedback("Dòng đã chọn được để lại cho lần xử lý sau.")

    def _tab_changed(self, index: int):
        if index == 3:
            self.refresh_review()

    def refresh_review(self):
        rows = []
        rows.extend(("Sửa thông tin", data["doc"], data["source"], data["decision"] or data["status"])
                    for data in self.edit_rows)
        rows.extend(("Chọn Bill/SQT", data["doc"], data["code"], data["decision"] or data["status"])
                    for data in self.choice_rows)
        rows.extend(("So sánh giá trị", data["doc"], data["field"], data["decision"] or data["status"])
                    for data in self.diff_rows)
        self.review_table.setRowCount(len(rows))
        for row, values in enumerate(rows):
            for col, value in enumerate(values):
                self.review_table.setItem(row, col, item(value))
        pending = sum(data["status"] == "Chưa xử lý" for group in (self.edit_rows, self.choice_rows, self.diff_rows) for data in group)
        self.review_summary.setText(
            f"Còn {pending} mục chưa xử lý. Bạn có thể quay lại các nhóm để hoàn tất hoặc để sau."
            if pending else
            "Mọi mục đã có quyết định. Trong PM thật, đây là bước kiểm tra cuối trước khi ghi Excel."
        )

    def finish(self):
        QMessageBox.information(self, "Xác nhận nhập · Demo", "Đã mô phỏng bước xác nhận. Không file thật nào được thay đổi.")

    def reset(self):
        self.edit_rows, self.choice_rows, self.diff_rows = copy.deepcopy(self.initial_rows)
        for row, data in enumerate(self.edit_rows):
            self.edit_table.item(row, 3).setText(data["date"])
            self.edit_table.item(row, 4).setText(data["container"])
            self.edit_table.item(row, 5).setText(data["status"])
        for row, data in enumerate(self.choice_rows):
            self.choice_boxes[row].setCurrentIndex(0)
            self.choice_table.item(row, 5).setText(data["status"])
        for row, data in enumerate(self.diff_rows):
            self.diff_boxes[row].setCurrentIndex(0)
            self.diff_custom[row].clear()
            self.diff_table.item(row, 6).setText(data["status"])
        self.field_filter.setCurrentIndex(0)
        self.rule_box.setCurrentIndex(0)
        self.tabs.setCurrentIndex(0)
        self.refresh_counts()
        self.feedback.setText("Demo đã trở về dữ liệu mẫu ban đầu.")


def main() -> int:
    app = QApplication(sys.argv)
    font_file = Path("C:/Windows/Fonts/segoeui.ttf")
    if font_file.is_file():
        QFontDatabase.addApplicationFont(str(font_file))
    app.setStyleSheet(STYLE)
    window = BulkWorkspaceDemo()
    window.show()
    return app.exec()


if __name__ == "__main__":
    raise SystemExit(main())
