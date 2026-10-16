"""Demo độc lập cho luồng quyết toán. Chỉ thay đổi dữ liệu mẫu trong bộ nhớ.

Chạy bằng: python demo_giao_dien.py
"""

from __future__ import annotations

import copy
import sys
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path

from PySide6.QtCore import Qt
from PySide6.QtGui import QIcon, QKeySequence
from PySide6.QtWidgets import (
    QApplication,
    QButtonGroup,
    QCheckBox,
    QComboBox,
    QDialog,
    QFileDialog,
    QFormLayout,
    QFrame,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QLineEdit,
    QListWidget,
    QMainWindow,
    QMenu,
    QMessageBox,
    QPlainTextEdit,
    QPushButton,
    QRadioButton,
    QScrollArea,
    QStackedWidget,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)


STYLE = """
* { font-family: 'Segoe UI', Arial, sans-serif; font-size: 10pt; color: #203047; }
QWidget#root, QWidget#page { background: #F5F7FB; }
QWidget#sidebar { background: #14263F; }
QLabel#brand { color: #FFFFFF; font-size: 16pt; font-weight: 700; }
QLabel#sideNote { color: #AFC2DD; font-size: 9pt; }
QListWidget#nav { background: transparent; border: 0; outline: 0; }
QListWidget#nav::item { color: #E0E9F5; padding: 12px 14px; margin: 3px 8px; border-radius: 8px; }
QListWidget#nav::item:hover { background: #243D5D; }
QListWidget#nav::item:selected { background: #2F62B8; color: white; font-weight: 700; }
QLabel#title { color: #102541; font-size: 18pt; font-weight: 700; }
QLabel#section { color: #102541; font-size: 12pt; font-weight: 700; }
QLabel#muted { color: #61718A; }
QLabel#status { color: #225EA8; font-weight: 600; }
QLabel#warning { color: #A35C13; font-weight: 600; }
QLabel#success { color: #16844B; font-weight: 600; }
QFrame#task { background: white; border: 1px solid #DCE4EF; border-radius: 10px; }
QFrame#detail { background: white; border: 1px solid #DCE4EF; border-radius: 10px; }
QFrame#divider { background: #E2E8F0; max-height: 1px; min-height: 1px; }
QPushButton { background: white; border: 1px solid #C7D3E2; border-radius: 8px; padding: 8px 13px; font-weight: 600; }
QPushButton:hover { background: #EDF3FB; }
QPushButton:disabled { background: #F1F3F6; color: #9AA7B8; border-color: #DCE2EA; }
QPushButton[primary='true'] { background: #285DB7; border-color: #285DB7; color: white; }
QPushButton[primary='true']:hover { background: #1F4F9D; }
QPushButton[primary='true']:disabled { background: #B8CBEA; border-color: #B8CBEA; color: white; }
QLineEdit, QComboBox { background: white; border: 1px solid #C8D3E1; border-radius: 7px; padding: 7px 9px; min-height: 19px; }
QLineEdit:focus, QComboBox:focus { border: 1px solid #285DB7; }
QTableWidget { background: white; border: 1px solid #DCE4EF; gridline-color: #E6ECF3; selection-background-color: #DDEBFF; selection-color: #172C4A; }
QTableWidget::item { padding: 4px; }
QHeaderView::section { background: #F2F5F9; border: 0; border-right: 1px solid #E2E8F0; border-bottom: 1px solid #DCE4EF; padding: 9px; color: #344A65; font-weight: 700; }
QTabWidget::pane { border: 1px solid #DCE4EF; background: white; }
QTabBar::tab { background: #EDF1F6; padding: 10px 18px; }
QTabBar::tab:selected { background: white; color: #285DB7; font-weight: 700; }
QScrollArea { border: 0; }
QPlainTextEdit { background: #F8FAFC; border: 1px solid #DCE4EF; border-radius: 7px; }
"""


def lbl(text: str, role: str = "muted", wrap: bool = True) -> QLabel:
    widget = QLabel(text)
    widget.setObjectName(role)
    widget.setWordWrap(wrap)
    return widget


def btn(text: str, action, primary: bool = False) -> QPushButton:
    widget = QPushButton(text)
    widget.setCursor(Qt.PointingHandCursor)
    if primary:
        widget.setProperty("primary", "true")
    widget.clicked.connect(action)
    return widget


def clear_layout(layout: QVBoxLayout) -> None:
    while layout.count():
        item = layout.takeAt(0)
        if item.widget() is not None:
            widget = item.widget()
            widget.hide()
            widget.setParent(None)
            widget.deleteLater()
        elif item.layout() is not None:
            clear_layout(item.layout())


def parse_date(value: str) -> datetime | None:
    try:
        return datetime.strptime(value.strip(), "%d/%m/%Y")
    except ValueError:
        return None


def parse_amount(value: str) -> float | None:
    try:
        return float(value.replace(".", "").replace(",", ".").replace("đ", "").strip())
    except ValueError:
        return None


def amount_text(value: float, kind: str) -> str:
    if kind == "Phiếu cân":
        return f"{value:,.2f}".replace(",", "_").replace(".", ",").replace("_", ".") + " tấn"
    return f"{value:,.0f}".replace(",", ".") + " đ"


@dataclass
class Document:
    id: int
    kind: str
    source: str
    reference: str
    container: str
    date: str
    amount: float
    needs_review: bool
    warnings: int


@dataclass
class Issue:
    id: int
    doc_id: int
    kind: str
    title: str
    reason: str
    status: str = "Chưa xử lý"
    answer: str = ""


class DemoData:
    """Dữ liệu đủ lớn để thử lọc, sắp xếp và các cách xử lý hàng loạt."""

    def __init__(self) -> None:
        self.docs: dict[int, Document] = {}
        self.issues: dict[int, Issue] = {}
        self.expected_targets: dict[int, tuple[str, str]] = {}
        self.imported = False
        self.last_automation = "Chưa chạy tác vụ nào trong phiên demo."
        self._populate()

    def _populate(self) -> None:
        for number in range(1, 59):
            kind = "Khoản chi" if number <= 40 else "Phiếu cân" if number <= 50 else "Bill"
            prefix = "HD" if kind == "Khoản chi" else "PC" if kind == "Phiếu cân" else "BL"
            date = f"{12 + (number - 1) % 14:02d}/05/2026"
            container = f"GAOU{2100000 + number:07d}"
            self.expected_targets[number] = (date, container)
            if number <= 15:
                container = ""
            elif number == 16:
                container = "DRYU3093608"
            source = "219.xlsx" if kind == "Khoản chi" else "phieu_can.pdf" if kind == "Phiếu cân" else "bill.pdf"
            amount = float(1_500_000 + number * 137_000) if kind != "Phiếu cân" else 24.5 + number / 10
            self.docs[number] = Document(
                number, kind, source, f"{prefix}-{number:04d}", container, date,
                amount, number <= 23, 1 if number <= 48 else 0,
            )
        for number in range(1, 17):
            self.issues[number] = Issue(
                number, number, "Không tìm thấy SQT", "Không tìm thấy số quyết toán",
                "Chưa ghép được khoản chi vì ngày đóng hoặc container chưa khớp bảng theo dõi.",
            )
        for number in range(17, 21):
            self.issues[number] = Issue(
                number, number, "Dữ liệu khác nhau", "Dữ liệu mới khác bảng theo dõi",
                "Đơn giá trên chứng từ mới khác giá trị đang có.",
            )
        for number in (21, 22):
            self.docs[number].container = "GAOU9999999"
            self.issues[number] = Issue(
                number, number, "Nhiều Bill phù hợp", "Có nhiều vận đơn phù hợp",
                "Chọn đúng vận đơn để ghép với chứng từ này.",
            )

    def pending_count(self) -> int:
        return sum(item.status == "Chưa xử lý" for item in self.issues.values())

    def reanalyze_missing(self, changed_doc_ids: set[int] | None = None) -> None:
        for issue in self.issues.values():
            if issue.kind != "Không tìm thấy SQT":
                continue
            if issue.status == "Để sau" and (
                changed_doc_ids is None or issue.doc_id not in changed_doc_ids
            ):
                continue
            doc = self.docs.get(issue.doc_id)
            if doc is None:
                continue
            expected = self.expected_targets[doc.id]
            matched = (doc.date.strip(), doc.container.strip().upper()) == expected
            issue.status = "Đã xử lý" if matched else "Chưa xử lý"
            issue.answer = f"Đã ghép theo {doc.date} · {doc.container}" if matched else ""
            if matched:
                doc.needs_review = False


class EditDocumentDialog(QDialog):
    def __init__(self, doc: Document, parent=None):
        super().__init__(parent)
        self.doc = doc
        self.setWindowTitle("Sửa thông tin chứng từ")
        self.resize(540, 390)
        outer = QVBoxLayout(self)
        outer.setContentsMargins(22, 20, 22, 20)
        outer.setSpacing(12)
        outer.addWidget(lbl("Sửa thông tin chứng từ", "title"))
        outer.addWidget(lbl(f"{doc.kind} · Nguồn: {doc.source}"))
        form = QFormLayout()
        form.setSpacing(10)
        self.reference = QLineEdit(doc.reference)
        self.container = QLineEdit(doc.container)
        self.date = QLineEdit(doc.date)
        self.amount = QLineEdit(amount_text(doc.amount, doc.kind).replace(" tấn", "").replace(" đ", ""))
        form.addRow("Số chứng từ", self.reference)
        form.addRow("Container", self.container)
        form.addRow("Ngày đóng (ngày/tháng/năm)", self.date)
        form.addRow("Số tấn" if doc.kind == "Phiếu cân" else "Tổng tiền", self.amount)
        outer.addLayout(form)
        outer.addStretch()
        actions = QHBoxLayout()
        actions.addStretch()
        actions.addWidget(btn("Hủy", self.reject))
        actions.addWidget(btn("Lưu thay đổi", self._save, primary=True))
        outer.addLayout(actions)

    def _save(self) -> None:
        if not self.reference.text().strip():
            QMessageBox.information(self, "Thiếu số chứng từ", "Hãy nhập số chứng từ.")
            return
        if self.date.text().strip() and parse_date(self.date.text()) is None:
            QMessageBox.information(self, "Ngày chưa hợp lệ", "Hãy nhập ngày theo dạng 15/05/2026.")
            return
        value = parse_amount(self.amount.text())
        if value is None:
            QMessageBox.information(self, "Số chưa hợp lệ", "Hãy kiểm tra lại số tấn hoặc tổng tiền.")
            return
        self.doc.reference = self.reference.text().strip()
        self.doc.container = self.container.text().strip().upper()
        self.doc.date = self.date.text().strip()
        self.doc.amount = value
        if self.doc.container and self.doc.date:
            self.doc.needs_review = False
        self.accept()


class ReviewDialog(QDialog):
    HEADERS = ["STT", "Loại chứng từ", "File nguồn", "Container", "Ngày đóng", "Số chứng từ", "Số tấn / Tổng tiền", "Trạng thái"]

    def __init__(self, data: DemoData, parent=None):
        super().__init__(parent)
        self.data = data
        self.backup = copy.deepcopy((data.docs, data.issues))
        self.dirty = False
        self.sort_column = 0
        self.sort_ascending = True
        self.visible_ids: list[int] = []
        self.setWindowTitle("Kiểm tra dữ liệu bóc tách · Demo")
        self.resize(1210, 710)
        self.setMinimumSize(920, 540)
        outer = QVBoxLayout(self)
        outer.setContentsMargins(16, 16, 16, 14)
        outer.setSpacing(11)
        outer.addWidget(lbl("Kiểm tra dữ liệu bóc tách", "title"))
        outer.addWidget(lbl("Dữ liệu mẫu · Chọn dòng cần xem, lọc theo trạng thái hoặc bấm tiêu đề cột để sắp xếp."))
        self.summary = lbl("", "status")
        outer.addWidget(self.summary)
        filters = QHBoxLayout()
        self.status = QComboBox()
        self.status.addItems(["Tất cả trạng thái", "Cần kiểm tra", "Sẵn sàng"])
        self.status.setCurrentIndex(1)
        self.kind = QComboBox()
        self.kind.addItems(["Mọi chứng từ", "Khoản chi", "Phiếu cân", "Bill"])
        self.warning = QComboBox()
        self.warning.addItems(["Mọi cảnh báo", "Có cảnh báo", "Không cảnh báo"])
        self.search = QLineEdit()
        self.search.setPlaceholderText("Tìm container, số chứng từ, file nguồn…")
        for widget in (self.status, self.kind, self.warning):
            widget.currentIndexChanged.connect(self.refresh)
        self.search.textChanged.connect(self.refresh)
        filters.addWidget(self.status)
        filters.addWidget(self.kind)
        filters.addWidget(self.warning)
        filters.addWidget(self.search, stretch=1)
        outer.addLayout(filters)
        self.table = QTableWidget(0, len(self.HEADERS))
        self.table.setHorizontalHeaderLabels(self.HEADERS)
        self.table.setSelectionBehavior(QTableWidget.SelectRows)
        self.table.setSelectionMode(QTableWidget.SingleSelection)
        self.table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.table.setAlternatingRowColors(True)
        self.table.verticalHeader().setVisible(False)
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.Interactive)
        self.table.horizontalHeader().setStretchLastSection(True)
        for column, width in enumerate((55, 125, 95, 145, 115, 125, 175)):
            self.table.setColumnWidth(column, width)
        self.table.horizontalHeader().setSortIndicatorShown(True)
        self.table.horizontalHeader().sectionClicked.connect(self._sort_by)
        self.table.cellDoubleClicked.connect(lambda _r, _c: self._edit())
        outer.addWidget(self.table, stretch=1)
        self.shown = lbl("")
        actions = QHBoxLayout()
        actions.addWidget(self.shown)
        actions.addStretch()
        actions.addWidget(btn("Sửa dòng", self._edit))
        actions.addWidget(btn("Xóa dòng", self._delete))
        actions.addWidget(btn("Lưu", self.accept, primary=True))
        actions.addWidget(btn("Đóng", self.reject))
        outer.addLayout(actions)
        self.refresh()

    def _selected_id(self) -> int | None:
        row = self.table.currentRow()
        if 0 <= row < len(self.visible_ids):
            return self.visible_ids[row]
        return None

    def _sort_by(self, column: int) -> None:
        if self.sort_column == column:
            self.sort_ascending = not self.sort_ascending
        else:
            self.sort_column = column
            self.sort_ascending = True
        self.refresh()

    def _sort_value(self, doc: Document):
        return (
            doc.id,
            doc.kind.casefold(),
            doc.source.casefold(),
            doc.container.casefold() if doc.container else None,
            parse_date(doc.date),
            doc.reference.casefold(),
            doc.amount,
            0 if doc.needs_review else 1,
        )[self.sort_column]

    def refresh(self, *_args) -> None:
        selected_id = self._selected_id()
        all_docs = list(self.data.docs.values())
        pending = sum(doc.needs_review for doc in all_docs)
        warnings = sum(doc.warnings for doc in all_docs)
        self.summary.setText(
            f"{len(all_docs)} chứng từ  ·  {pending} cần kiểm tra  ·  "
            f"{len(all_docs) - pending} sẵn sàng  ·  {warnings} cảnh báo"
        )
        query = self.search.text().strip().casefold()
        chosen: list[Document] = []
        for doc in all_docs:
            if self.status.currentIndex() == 1 and not doc.needs_review:
                continue
            if self.status.currentIndex() == 2 and doc.needs_review:
                continue
            if self.kind.currentIndex() and doc.kind != self.kind.currentText():
                continue
            if self.warning.currentIndex() == 1 and doc.warnings == 0:
                continue
            if self.warning.currentIndex() == 2 and doc.warnings:
                continue
            if query and query not in " ".join((doc.reference, doc.container, doc.source, doc.kind)).casefold():
                continue
            chosen.append(doc)
        filled = [doc for doc in chosen if self._sort_value(doc) is not None]
        blank = [doc for doc in chosen if self._sort_value(doc) is None]
        filled.sort(key=self._sort_value, reverse=not self.sort_ascending)
        chosen = filled + blank
        self.visible_ids = [doc.id for doc in chosen]
        self.table.blockSignals(True)
        self.table.setRowCount(len(chosen))
        for row, doc in enumerate(chosen):
            values = [
                str(doc.id), doc.kind, doc.source, doc.container or "—", doc.date or "—",
                doc.reference, amount_text(doc.amount, doc.kind),
                "Cần kiểm tra" if doc.needs_review else "Sẵn sàng",
            ]
            for col, value in enumerate(values):
                item = QTableWidgetItem(value)
                item.setData(Qt.UserRole, doc.id)
                self.table.setItem(row, col, item)
        self.table.blockSignals(False)
        if chosen:
            self.table.selectRow(self.visible_ids.index(selected_id) if selected_id in self.visible_ids else 0)
        self.table.horizontalHeader().setSortIndicator(
            self.sort_column, Qt.AscendingOrder if self.sort_ascending else Qt.DescendingOrder
        )
        self.shown.setText(f"Đang hiển thị {len(chosen)}/{len(all_docs)} chứng từ")

    def _edit(self) -> None:
        doc_id = self._selected_id()
        if doc_id is None:
            QMessageBox.information(self, "Chưa chọn dòng", "Hãy chọn một chứng từ trong bảng.")
            return
        if EditDocumentDialog(self.data.docs[doc_id], self).exec() == QDialog.Accepted:
            self.dirty = True
            self.data.reanalyze_missing({doc_id})
            self.refresh()

    def _delete(self) -> None:
        doc_id = self._selected_id()
        if doc_id is None:
            QMessageBox.information(self, "Chưa chọn dòng", "Hãy chọn một chứng từ trong bảng.")
            return
        if QMessageBox.question(self, "Xóa chứng từ", "Xóa chứng từ này khỏi lô dữ liệu mẫu?") != QMessageBox.Yes:
            return
        self.data.docs.pop(doc_id)
        for issue_id in [key for key, issue in self.data.issues.items() if issue.doc_id == doc_id]:
            self.data.issues.pop(issue_id)
        self.dirty = True
        self.refresh()

    def reject(self) -> None:
        if self.dirty and QMessageBox.question(
            self, "Chưa lưu thay đổi", "Đóng và bỏ các thay đổi trong lần mở cửa sổ này?"
        ) != QMessageBox.Yes:
            return
        self.data.docs, self.data.issues = self.backup
        super().reject()


class PasteTable(QTableWidget):
    """Dán nhiều ô từ Excel vào cột Ngày đóng và Container."""

    def keyPressEvent(self, event) -> None:  # noqa: N802 - Qt API
        if event.matches(QKeySequence.Paste):
            content = QApplication.clipboard().text().replace("\r\n", "\n").strip("\n")
            start_row = max(0, self.currentRow())
            start_col = max(1, self.currentColumn())
            for row_offset, line in enumerate(content.split("\n")):
                row = start_row + row_offset
                if row >= self.rowCount():
                    break
                for col_offset, value in enumerate(line.split("\t")):
                    col = start_col + col_offset
                    if col > 2:
                        break
                    self.setItem(row, col, QTableWidgetItem(value.strip()))
            return
        super().keyPressEvent(event)


class QuickEditDialog(QDialog):
    def __init__(self, data: DemoData, issues: list[Issue], parent=None):
        super().__init__(parent)
        self.data = data
        self.issues = issues
        self.matched = 0
        self.setWindowTitle("Sửa nhanh nhiều dòng · Demo")
        self.resize(770, 610)
        layout = QVBoxLayout(self)
        layout.setContentsMargins(18, 18, 18, 16)
        layout.addWidget(lbl(f"Sửa nhanh {len(issues)} dòng", "title"))
        layout.addWidget(lbl("Chỉ hiện hai thông tin dùng để tìm số quyết toán. Có thể dán nhiều ô từ Excel vào bảng."))
        self.table = PasteTable(len(issues), 3)
        self.table.setHorizontalHeaderLabels(["Chứng từ", "Ngày đóng", "Container"])
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.table.verticalHeader().setVisible(False)
        for row, issue in enumerate(issues):
            doc = data.docs[issue.doc_id]
            title = QTableWidgetItem(doc.reference)
            title.setFlags(title.flags() & ~Qt.ItemIsEditable)
            self.table.setItem(row, 0, title)
            self.table.setItem(row, 1, QTableWidgetItem(doc.date))
            self.table.setItem(row, 2, QTableWidgetItem(doc.container))
        layout.addWidget(self.table, stretch=1)
        layout.addWidget(lbl("Trong demo, nút điền mẫu giúp thử thao tác. Ứng dụng thật sẽ dùng giá trị bạn nhập hoặc dán.", "muted"))
        actions = QHBoxLayout()
        actions.addWidget(btn("Điền dữ liệu mẫu để thử", self._fill_sample))
        actions.addStretch()
        actions.addWidget(btn("Hủy", self.reject))
        actions.addWidget(btn("Kiểm tra và áp dụng", self._apply, primary=True))
        layout.addLayout(actions)

    def _fill_sample(self) -> None:
        for row, issue in enumerate(self.issues):
            date, container = self.data.expected_targets[issue.doc_id]
            self.table.item(row, 1).setText(date)
            self.table.item(row, 2).setText(container)

    def _apply(self) -> None:
        values: list[tuple[str, str]] = []
        for row in range(self.table.rowCount()):
            date = self.table.item(row, 1).text().strip() if self.table.item(row, 1) else ""
            container = self.table.item(row, 2).text().strip().upper() if self.table.item(row, 2) else ""
            if date and parse_date(date) is None:
                QMessageBox.information(self, "Ngày chưa hợp lệ", f"Dòng {row + 1}: hãy nhập ngày theo dạng 15/05/2026.")
                return
            values.append((date, container))
        for issue, (date, container) in zip(self.issues, values):
            doc = self.data.docs[issue.doc_id]
            doc.date = date
            doc.container = container
        self.data.reanalyze_missing({issue.doc_id for issue in self.issues})
        self.matched = sum(issue.status == "Đã xử lý" for issue in self.issues)
        self.accept()


class BulkChoiceDialog(QDialog):
    """Chọn một phương án chung và xem trước chính xác những dòng sẽ đổi."""

    def __init__(self, data: DemoData, issues: list[Issue], parent=None):
        super().__init__(parent)
        self.issues = issues
        self.choice = ""
        self.setWindowTitle("Xem trước xử lý hàng loạt · Demo")
        self.resize(760, 520)
        layout = QVBoxLayout(self)
        layout.setContentsMargins(18, 18, 18, 16)
        layout.addWidget(lbl(f"Áp dụng cho {len(issues)} dòng đã chọn", "title"))
        kind = issues[0].kind
        self.options = QComboBox()
        if kind == "Dữ liệu khác nhau":
            self.options.addItems(["Giữ đơn giá đang có", "Dùng đơn giá từ chứng từ mới"])
            description = "Xem từng chứng từ trước khi dùng một phương án cho tất cả."
        else:
            self.options.addItems(["Chọn vận đơn BL-481", "Chọn vận đơn BL-482"])
            description = "Các dòng này có cùng hai vận đơn có thể chọn."
        layout.addWidget(lbl(description))
        layout.addWidget(self.options)
        table = QTableWidget(len(issues), 4)
        table.setHorizontalHeaderLabels(["Số chứng từ", "Container", "Hiện tại", "Dữ liệu mới / lựa chọn"])
        table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        table.verticalHeader().setVisible(False)
        table.setEditTriggers(QTableWidget.NoEditTriggers)
        for row, issue in enumerate(issues):
            doc = data.docs[issue.doc_id]
            values = (
                doc.reference,
                doc.container or "—",
                "1.700.000 đ" if kind == "Dữ liệu khác nhau" else "Chưa chọn Bill",
                "1.780.000 đ" if kind == "Dữ liệu khác nhau" else "BL-481 hoặc BL-482",
            )
            for column, value in enumerate(values):
                table.setItem(row, column, QTableWidgetItem(value))
        layout.addWidget(table, stretch=1)
        layout.addWidget(lbl("Chỉ các dòng trong bảng này được thay đổi. Bạn có thể hoàn tác trước khi nhập."))
        actions = QHBoxLayout()
        actions.addStretch()
        actions.addWidget(btn("Quay lại", self.reject))
        actions.addWidget(btn("Áp dụng cho các dòng này", self._accept, primary=True))
        layout.addLayout(actions)

    def _accept(self) -> None:
        self.choice = self.options.currentText()
        self.accept()


class IssueDialog(QDialog):
    HEADERS = ["Chọn", "STT", "Chứng từ", "Số chứng từ", "Container", "Vấn đề", "Trạng thái"]

    def __init__(self, data: DemoData, parent=None):
        super().__init__(parent)
        self.data = data
        self.selected_ids: set[int] = set()
        self.visible_ids: list[int] = []
        self.sort_column = 1
        self.sort_ascending = True
        self.undo_state = None
        self.setWindowTitle("Xử lý dữ liệu trước khi nhập · Demo")
        self.resize(1350, 800)
        self.setMinimumSize(1050, 650)
        outer = QVBoxLayout(self)
        outer.setContentsMargins(16, 15, 16, 14)
        outer.setSpacing(10)
        outer.addWidget(lbl("Xử lý dữ liệu trước khi nhập", "title"))
        outer.addWidget(lbl("Chọn các dòng cần xử lý. Có thể chọn nhiều loại vấn đề cùng lúc; bảng theo dõi chưa thay đổi."))
        self.summary = lbl("", "status")
        outer.addWidget(self.summary)
        controls = QHBoxLayout()
        self.status_filter = QComboBox()
        self.status_filter.addItems(["Chưa xử lý", "Tất cả trạng thái", "Đã xử lý", "Để sau"])
        self.problem_filter = QComboBox()
        self.problem_filter.addItems(["Mọi vấn đề", "Không tìm thấy SQT", "Dữ liệu khác nhau", "Nhiều Bill phù hợp"])
        self.search = QLineEdit()
        self.search.setPlaceholderText("Tìm số chứng từ, container, vấn đề…")
        self.status_filter.currentIndexChanged.connect(self.refresh)
        self.problem_filter.currentIndexChanged.connect(self.refresh)
        self.search.textChanged.connect(self.refresh)
        controls.addWidget(self.status_filter)
        controls.addWidget(self.problem_filter)
        controls.addWidget(self.search, stretch=1)
        outer.addLayout(controls)
        body = QHBoxLayout()
        left = QFrame()
        left.setObjectName("detail")
        left_layout = QVBoxLayout(left)
        left_layout.setContentsMargins(10, 10, 10, 10)
        select_row = QHBoxLayout()
        select_row.addWidget(btn("Chọn tất cả đang hiển thị", self._select_visible))
        select_row.addWidget(btn("Bỏ chọn", self._clear_selection))
        select_row.addStretch()
        left_layout.addLayout(select_row)
        self.table = QTableWidget(0, len(self.HEADERS))
        self.table.setHorizontalHeaderLabels(self.HEADERS)
        self.table.setSelectionBehavior(QTableWidget.SelectRows)
        self.table.setSelectionMode(QTableWidget.SingleSelection)
        self.table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.table.setAlternatingRowColors(True)
        self.table.verticalHeader().setVisible(False)
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.Interactive)
        self.table.horizontalHeader().setStretchLastSection(True)
        for column, width in enumerate((50, 50, 105, 105, 130, 195)):
            self.table.setColumnWidth(column, width)
        self.table.horizontalHeader().setSortIndicatorShown(True)
        self.table.horizontalHeader().sectionClicked.connect(self._sort_by)
        self.table.itemChanged.connect(self._check_changed)
        self.table.itemSelectionChanged.connect(self._show_selected)
        left_layout.addWidget(self.table, stretch=1)
        self.visible_count = lbl("")
        left_layout.addWidget(self.visible_count)
        body.addWidget(left, 62)
        detail = QFrame()
        detail.setObjectName("detail")
        self.detail_layout = QVBoxLayout(detail)
        self.detail_layout.setContentsMargins(17, 16, 17, 16)
        body.addWidget(detail, 38)
        outer.addLayout(body, stretch=1)
        self.bulk_bar = QHBoxLayout()
        self.selected_count = lbl("Chưa chọn dòng nào")
        self.bulk_bar.addWidget(self.selected_count)
        self.bulk_bar.addStretch()
        self.bulk_button = btn("Xử lý hàng loạt ▾", self._show_bulk_menu)
        self.bulk_button.setToolTip("Xem các cách xử lý phù hợp với những dòng đã chọn, theo từng loại vấn đề")
        self.defer_button = btn("Để xử lý sau", self._defer_selected)
        self.undo_button = btn("Hoàn tác", self._undo)
        for widget in (self.bulk_button, self.defer_button, self.undo_button):
            self.bulk_bar.addWidget(widget)
        outer.addLayout(self.bulk_bar)
        divider = QFrame()
        divider.setObjectName("divider")
        outer.addWidget(divider)
        footer = QHBoxLayout()
        self.progress = lbl("")
        footer.addWidget(self.progress)
        footer.addStretch()
        footer.addWidget(btn("Lưu và đóng", self._save_close))
        self.commit_button = btn("Xem và xác nhận nhập", self._commit, primary=True)
        footer.addWidget(self.commit_button)
        outer.addLayout(footer)
        self.refresh()

    def _current_issue_id(self) -> int | None:
        row = self.table.currentRow()
        if 0 <= row < len(self.visible_ids):
            return self.visible_ids[row]
        return None

    def _sort_by(self, column: int) -> None:
        if column == 0:
            return
        if column == self.sort_column:
            self.sort_ascending = not self.sort_ascending
        else:
            self.sort_column = column
            self.sort_ascending = True
        self.refresh()

    def _sort_value(self, issue: Issue):
        doc = self.data.docs[issue.doc_id]
        return {
            1: issue.id,
            2: doc.kind.casefold(),
            3: doc.reference.casefold(),
            4: doc.container.casefold() if doc.container else None,
            5: issue.title.casefold(),
            6: issue.status.casefold(),
        }.get(self.sort_column, issue.id)

    def _matches(self, issue: Issue) -> bool:
        doc = self.data.docs[issue.doc_id]
        status = self.status_filter.currentText()
        if status != "Tất cả trạng thái" and status != issue.status:
            return False
        problem = self.problem_filter.currentText()
        if problem != "Mọi vấn đề" and problem != issue.kind:
            return False
        query = self.search.text().strip().casefold()
        return not query or query in " ".join(
            (doc.reference, doc.container, doc.source, issue.title, issue.reason)
        ).casefold()

    def refresh(self, *_args) -> None:
        current_id = self._current_issue_id()
        issues = [issue for issue in self.data.issues.values() if issue.doc_id in self.data.docs and self._matches(issue)]
        filled = [issue for issue in issues if self._sort_value(issue) is not None]
        blank = [issue for issue in issues if self._sort_value(issue) is None]
        filled.sort(key=self._sort_value, reverse=not self.sort_ascending)
        issues = filled + blank
        self.visible_ids = [issue.id for issue in issues]
        self.table.blockSignals(True)
        self.table.setRowCount(len(issues))
        for row, issue in enumerate(issues):
            doc = self.data.docs[issue.doc_id]
            check = QTableWidgetItem()
            check.setFlags(Qt.ItemIsEnabled | Qt.ItemIsUserCheckable)
            check.setCheckState(Qt.Checked if issue.id in self.selected_ids else Qt.Unchecked)
            self.table.setItem(row, 0, check)
            for col, value in enumerate((
                str(issue.id), doc.kind, doc.reference, doc.container or "—", issue.title, issue.status
            ), start=1):
                cell = QTableWidgetItem(value)
                cell.setFlags(cell.flags() & ~Qt.ItemIsEditable)
                self.table.setItem(row, col, cell)
        self.table.blockSignals(False)
        if issues:
            self.table.selectRow(self.visible_ids.index(current_id) if current_id in self.visible_ids else 0)
        else:
            self._empty_detail()
        self.table.horizontalHeader().setSortIndicator(
            self.sort_column, Qt.AscendingOrder if self.sort_ascending else Qt.DescendingOrder
        )
        pending = self.data.pending_count()
        done = sum(issue.status == "Đã xử lý" for issue in self.data.issues.values())
        deferred = sum(issue.status == "Để sau" for issue in self.data.issues.values())
        self.summary.setText(
            f"{pending} chưa xử lý  ·  {done} đã xử lý  ·  {deferred} để sau  ·  "
            f"{len(self.data.issues)} vấn đề trong lô"
        )
        self.visible_count.setText(f"Đang hiển thị {len(issues)}/{len(self.data.issues)} dòng")
        self.progress.setText(f"Đã có quyết định {done + deferred}/{len(self.data.issues)} dòng")
        self.commit_button.setEnabled(pending == 0 and not self.data.imported)
        self._refresh_bulk_bar()

    def _check_changed(self, item: QTableWidgetItem) -> None:
        if item.column() != 0 or item.row() >= len(self.visible_ids):
            return
        issue_id = self.visible_ids[item.row()]
        if item.checkState() == Qt.Checked:
            self.selected_ids.add(issue_id)
        else:
            self.selected_ids.discard(issue_id)
        self._refresh_bulk_bar()

    def _select_visible(self) -> None:
        self.selected_ids.update(self.visible_ids)
        self.refresh()

    def _clear_selection(self) -> None:
        self.selected_ids.clear()
        self.refresh()

    def _selected_issues(self) -> list[Issue]:
        return [self.data.issues[key] for key in sorted(self.selected_ids) if key in self.data.issues]

    def _refresh_bulk_bar(self) -> None:
        selected = self._selected_issues()
        hidden = len(self.selected_ids - set(self.visible_ids))
        self.selected_count.setText(
            f"Đã chọn {len(selected)} dòng" + (f" · {hidden} dòng đang ẩn bởi bộ lọc" if hidden else "")
            if selected else "Chưa chọn dòng nào"
        )
        pending = [issue for issue in selected if issue.status == "Chưa xử lý"]
        self.bulk_button.setEnabled(bool(pending))
        self.defer_button.setEnabled(bool(pending))
        self.undo_button.setEnabled(self.undo_state is not None)

    def _build_bulk_menu(self) -> QMenu:
        menu = QMenu(self)
        pending = [issue for issue in self._selected_issues() if issue.status == "Chưa xử lý"]
        groups = (
            ("Không tìm thấy SQT", "Sửa nhanh thông tin để ghép lại", self._quick_edit),
            ("Dữ liệu khác nhau", "Chọn phương án cho dữ liệu khác nhau", self._common_action),
            ("Nhiều Bill phù hợp", "Chọn vận đơn phù hợp", self._common_action),
        )
        for kind, label, handler in groups:
            ids = tuple(issue.id for issue in pending if issue.kind == kind)
            if ids:
                action = menu.addAction(f"{label} · {len(ids)} dòng")
                action.triggered.connect(lambda _checked=False, ids=ids, handler=handler: handler(ids))
        if pending:
            menu.addSeparator()
            action = menu.addAction(f"Để cả {len(pending)} dòng chưa xử lý đã chọn lại sau")
            action.triggered.connect(self._defer_selected)
        return menu

    def _show_bulk_menu(self) -> None:
        menu = self._build_bulk_menu()
        menu.exec(self.bulk_button.mapToGlobal(self.bulk_button.rect().bottomLeft()))

    def _snapshot(self) -> None:
        self.undo_state = copy.deepcopy((self.data.docs, self.data.issues))

    def _undo(self) -> None:
        if self.undo_state is None:
            return
        self.data.docs, self.data.issues = self.undo_state
        self.undo_state = None
        self.refresh()

    def _quick_edit(self, issue_ids: tuple[int, ...]) -> None:
        issues = [self.data.issues[key] for key in issue_ids if key in self.selected_ids]
        if not issues or any(issue.kind != "Không tìm thấy SQT" or issue.status != "Chưa xử lý" for issue in issues):
            return
        old_undo = self.undo_state
        self._snapshot()
        dialog = QuickEditDialog(self.data, issues, self)
        if dialog.exec() != QDialog.Accepted:
            self.undo_state = old_undo
            return
        self.selected_ids.difference_update(issue.id for issue in issues if issue.status == "Đã xử lý")
        self.refresh()
        QMessageBox.information(
            self, "Đã kiểm tra lại",
            f"{dialog.matched}/{len(issues)} dòng đã tìm được số quyết toán trong dữ liệu mẫu. "
            "Các dòng còn lại vẫn ở trạng thái chưa xử lý."
        )

    def _common_action(self, issue_ids: tuple[int, ...]) -> None:
        issues = [self.data.issues[key] for key in issue_ids if key in self.selected_ids]
        if not issues:
            return
        kinds = {issue.kind for issue in issues}
        if kinds not in ({"Dữ liệu khác nhau"}, {"Nhiều Bill phù hợp"}) or any(
            issue.status != "Chưa xử lý" for issue in issues
        ):
            QMessageBox.information(self, "Chọn cùng loại vấn đề", "Hãy chọn các dòng có cùng loại vấn đề.")
            return
        dialog = BulkChoiceDialog(self.data, issues, self)
        if dialog.exec() != QDialog.Accepted:
            return
        self._snapshot()
        for issue in issues:
            issue.status = "Đã xử lý"
            issue.answer = dialog.choice
            self.data.docs[issue.doc_id].needs_review = False
        self.selected_ids.difference_update(issue.id for issue in issues)
        self.refresh()

    def _defer_selected(self) -> None:
        issues = [issue for issue in self._selected_issues() if issue.status == "Chưa xử lý"]
        if not issues:
            return
        if QMessageBox.question(
            self, "Để xử lý sau",
            f"Để {len(issues)} dòng đã chọn lại cho lần sau?\n\n"
            "Các dòng này không được nhập vào bảng theo dõi trong lần này và không bị xóa."
        ) != QMessageBox.Yes:
            return
        self._snapshot()
        for issue in issues:
            issue.status = "Để sau"
            issue.answer = ""
        self.selected_ids.difference_update(issue.id for issue in issues)
        self.refresh()

    def _show_selected(self) -> None:
        issue_id = self._current_issue_id()
        if issue_id is None:
            self._empty_detail()
            return
        self._show_issue(self.data.issues[issue_id])

    def _empty_detail(self) -> None:
        clear_layout(self.detail_layout)
        self.detail_layout.addWidget(lbl("Không có dòng phù hợp với bộ lọc."))
        self.detail_layout.addStretch()

    def _show_issue(self, issue: Issue) -> None:
        clear_layout(self.detail_layout)
        doc = self.data.docs[issue.doc_id]
        self.detail_layout.addWidget(lbl(issue.title, "section"))
        self.detail_layout.addWidget(lbl(issue.reason, "warning"))
        self.detail_layout.addSpacing(8)
        self.detail_layout.addWidget(lbl(f"Chứng từ: {doc.reference} · {doc.kind}"))
        self.detail_layout.addWidget(lbl(f"Ngày đóng: {doc.date or 'Chưa có'}"))
        self.detail_layout.addWidget(lbl(f"Container: {doc.container or 'Chưa có'}"))
        self.detail_layout.addWidget(lbl(f"Nguồn: {doc.source}"))
        self.detail_layout.addSpacing(10)
        self.detail_layout.addWidget(lbl("Cách xử lý", "section"))
        if issue.kind == "Không tìm thấy SQT":
            self.detail_layout.addWidget(lbl("Kiểm tra ngày đóng và container. Khi đủ thông tin, phần mềm sẽ thử ghép lại."))
            self.detail_layout.addWidget(btn("Sửa thông tin chứng từ", lambda: self._edit_single(issue)))
        else:
            choices = (
                ["Giữ đơn giá đang có", "Dùng đơn giá từ chứng từ mới"]
                if issue.kind == "Dữ liệu khác nhau"
                else ["Chọn vận đơn BL-481", "Chọn vận đơn BL-482"]
            )
            self.detail_choices = QButtonGroup(self)
            for text in choices:
                radio = QRadioButton(text)
                radio.setChecked(issue.answer == text)
                self.detail_choices.addButton(radio)
                self.detail_layout.addWidget(radio)
            self.detail_layout.addWidget(btn("Áp dụng cho dòng này", lambda: self._apply_single(issue), primary=True))
        if issue.status != "Chưa xử lý":
            self.detail_layout.addWidget(lbl(f"Đã chọn: {issue.answer or 'Để xử lý sau'}", "success"))
        self.detail_layout.addStretch()
        self.detail_layout.addWidget(btn("Để dòng này xử lý sau", lambda: self._defer_one(issue)))

    def _edit_single(self, issue: Issue) -> None:
        old_undo = self.undo_state
        self._snapshot()
        if EditDocumentDialog(self.data.docs[issue.doc_id], self).exec() != QDialog.Accepted:
            self.undo_state = old_undo
            return
        self.data.reanalyze_missing({issue.doc_id})
        self.refresh()
        if issue.status == "Chưa xử lý":
            QMessageBox.information(self, "Chưa tìm thấy số quyết toán", "Thông tin đã lưu nhưng chưa khớp dữ liệu mẫu. Hãy kiểm tra lại ngày đóng và container.")

    def _apply_single(self, issue: Issue) -> None:
        chosen = self.detail_choices.checkedButton()
        if chosen is None:
            QMessageBox.information(self, "Chưa chọn phương án", "Hãy chọn một phương án trước.")
            return
        self._snapshot()
        issue.status = "Đã xử lý"
        issue.answer = chosen.text()
        self.data.docs[issue.doc_id].needs_review = False
        self.refresh()

    def _defer_one(self, issue: Issue) -> None:
        self._snapshot()
        issue.status = "Để sau"
        issue.answer = ""
        self.refresh()

    def _save_close(self) -> None:
        QMessageBox.information(self, "Đã giữ lựa chọn", "Các lựa chọn được giữ trong phiên demo này. Chưa có file thật nào được thay đổi.")
        self.accept()

    def _commit(self) -> None:
        if self.data.pending_count():
            return
        done = sum(issue.status == "Đã xử lý" for issue in self.data.issues.values())
        deferred = sum(issue.status == "Để sau" for issue in self.data.issues.values())
        if QMessageBox.question(
            self, "Xác nhận nhập dữ liệu",
            f"Đã xử lý: {done} dòng\nĐể xử lý sau: {deferred} dòng\n\n"
            "Xác nhận cập nhật bảng theo dõi mẫu? Không có file thật nào được ghi."
        ) != QMessageBox.Yes:
            return
        self.data.imported = True
        self.accept()


class SelectSqtDialog(QDialog):
    def __init__(self, action: str, parent=None):
        super().__init__(parent)
        self.action = action
        self.selected: list[str] = []
        self.setWindowTitle(f"Chọn số quyết toán · {action} · Demo")
        self.resize(670, 570)
        outer = QVBoxLayout(self)
        outer.setContentsMargins(20, 18, 20, 18)
        outer.setSpacing(11)
        outer.addWidget(lbl(f"Chọn số quyết toán để {action.lower()}", "title"))
        outer.addWidget(lbl("Có thể chọn nhiều số. Những số đã nhập sẽ được hỏi lại trước khi chạy."))
        self.search = QLineEdit()
        self.search.setPlaceholderText("Tìm số quyết toán…")
        outer.addWidget(self.search)
        self.tabs = QStackedWidget()
        self.checkboxes: list[tuple[QCheckBox, bool]] = []
        tabs = QHBoxLayout()
        self.pending_button = btn("Cần nhập (4)", lambda: self._set_tab(0), primary=True)
        self.done_button = btn("Đã nhập (2)", lambda: self._set_tab(1))
        tabs.addWidget(self.pending_button)
        tabs.addWidget(self.done_button)
        tabs.addStretch()
        outer.addLayout(tabs)
        for numbers, imported in (
            (["QT-2026-101", "QT-2026-102", "QT-2026-103", "QT-2026-104"], False),
            (["QT-2026-088", "QT-2026-090"], True),
        ):
            page = QWidget()
            layout = QVBoxLayout(page)
            layout.setContentsMargins(12, 12, 12, 12)
            for number in numbers:
                row = QHBoxLayout()
                choice = QCheckBox(number)
                choice.setProperty("number", number)
                choice.stateChanged.connect(self._update_summary)
                self.checkboxes.append((choice, imported))
                row.addWidget(choice)
                row.addStretch()
                row.addWidget(lbl("Đã nhập" if imported else "Cần nhập"))
                layout.addLayout(row)
            layout.addStretch()
            self.tabs.addWidget(page)
        outer.addWidget(self.tabs, stretch=1)
        self.search.textChanged.connect(self._filter)
        select_actions = QHBoxLayout()
        select_actions.addWidget(btn("Chọn tất cả đang hiển thị", self._select_visible))
        select_actions.addWidget(btn("Bỏ chọn", self._clear_visible))
        select_actions.addStretch()
        outer.addLayout(select_actions)
        footer = QHBoxLayout()
        self.summary = lbl("Chưa chọn số quyết toán nào")
        footer.addWidget(self.summary)
        footer.addStretch()
        footer.addWidget(btn("Hủy", self.reject))
        footer.addWidget(btn("Bắt đầu nhập", self._accept, primary=True))
        outer.addLayout(footer)

    def _set_tab(self, index: int) -> None:
        self.tabs.setCurrentIndex(index)
        self.pending_button.setProperty("primary", "true" if index == 0 else "false")
        self.done_button.setProperty("primary", "true" if index == 1 else "false")
        for widget in (self.pending_button, self.done_button):
            widget.style().unpolish(widget)
            widget.style().polish(widget)

    def _filter(self) -> None:
        query = self.search.text().strip().casefold()
        for choice, _imported in self.checkboxes:
            choice.setVisible(query in choice.property("number").casefold())

    def _visible_choices(self):
        imported_tab = self.tabs.currentIndex() == 1
        query = self.search.text().strip().casefold()
        return [
            choice for choice, imported in self.checkboxes
            if imported == imported_tab and query in choice.property("number").casefold()
        ]

    def _select_visible(self) -> None:
        for choice in self._visible_choices():
            choice.setChecked(True)

    def _clear_visible(self) -> None:
        for choice in self._visible_choices():
            choice.setChecked(False)

    def _update_summary(self) -> None:
        count = sum(choice.isChecked() for choice, _ in self.checkboxes)
        self.summary.setText(f"Đã chọn {count} số quyết toán" if count else "Chưa chọn số quyết toán nào")

    def _accept(self) -> None:
        selected = [(choice, imported) for choice, imported in self.checkboxes if choice.isChecked()]
        if not selected:
            QMessageBox.information(self, "Chưa chọn dữ liệu", "Hãy chọn ít nhất một số quyết toán.")
            return
        repeat = sum(imported for _choice, imported in selected)
        if repeat and QMessageBox.question(
            self, "Xác nhận nhập lại",
            f"Có {repeat} số quyết toán đã nhập. Nếu tiếp tục, các số này sẽ được xử lý lại toàn bộ."
        ) != QMessageBox.Yes:
            return
        self.selected = [choice.property("number") for choice, _ in selected]
        self.accept()


class DemoWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.data = DemoData()
        self.setWindowTitle("Trợ lý quyết toán · Demo giao diện")
        self.resize(1130, 780)
        self.setMinimumSize(900, 630)
        self._build_shell()
        self._refresh()

    def _build_shell(self) -> None:
        root = QWidget()
        root.setObjectName("root")
        self.setCentralWidget(root)
        shell = QHBoxLayout(root)
        shell.setContentsMargins(0, 0, 0, 0)
        shell.setSpacing(0)
        sidebar = QWidget()
        sidebar.setObjectName("sidebar")
        sidebar.setFixedWidth(190)
        side_layout = QVBoxLayout(sidebar)
        side_layout.setContentsMargins(15, 22, 15, 18)
        side_layout.addWidget(lbl("KIKAI", "brand"))
        side_layout.addWidget(lbl("Trợ lý quyết toán", "sideNote"))
        side_layout.addSpacing(20)
        self.nav = QListWidget()
        self.nav.setObjectName("nav")
        self.nav.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self.nav.addItems(["Chức năng", "Cài đặt"])
        side_layout.addWidget(self.nav)
        side_layout.addStretch()
        side_layout.addWidget(lbl("DEMO · Dữ liệu mẫu", "sideNote"))
        shell.addWidget(sidebar)
        self.pages = QStackedWidget()
        shell.addWidget(self.pages, stretch=1)
        self.pages.addWidget(self._build_tasks())
        self.pages.addWidget(self._build_settings())
        self.nav.currentRowChanged.connect(self.pages.setCurrentIndex)
        self.nav.setCurrentRow(0)

    @staticmethod
    def _new_page() -> tuple[QScrollArea, QVBoxLayout]:
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        page = QWidget()
        page.setObjectName("page")
        layout = QVBoxLayout(page)
        layout.setContentsMargins(28, 23, 28, 26)
        layout.setSpacing(14)
        scroll.setWidget(page)
        return scroll, layout

    def _task(self, number: str, title: str, description: str, button_text: str, action):
        frame = QFrame()
        frame.setObjectName("task")
        layout = QVBoxLayout(frame)
        layout.setContentsMargins(18, 15, 18, 15)
        layout.setSpacing(8)
        row = QHBoxLayout()
        row.addWidget(lbl(number, "status"))
        row.addWidget(lbl(title, "section", wrap=False))
        row.addStretch()
        state = lbl("")
        row.addWidget(state)
        layout.addLayout(row)
        layout.addWidget(lbl(description))
        layout.addWidget(btn(button_text, action, primary=True), alignment=Qt.AlignRight)
        return frame, state

    def _build_tasks(self) -> QScrollArea:
        scroll, content = self._new_page()
        content.addWidget(lbl("Chức năng", "title"))
        content.addWidget(lbl("Thực hiện theo tác vụ bạn cần. Đây là bản demo; mọi thay đổi chỉ nằm trong dữ liệu mẫu."))
        first, self.state_1 = self._task(
            "01", "Mở trợ lý quyết toán",
            "Mở trợ lý để chuẩn bị dữ liệu chứng từ.", "Mở trợ lý", self._open_assistant,
        )
        content.addWidget(first)
        second, self.state_2 = self._task(
            "02", "Xem dữ liệu bóc tách",
            "Lọc, sắp xếp và sửa chứng từ trước khi nhập.", "Kiểm tra chứng từ", self._open_review,
        )
        content.addWidget(second)
        third, self.state_3 = self._task(
            "03", "Nhập lên file hàng ngày",
            "Giải quyết dữ liệu chưa khớp rồi xác nhận cập nhật.", "Xử lý và nhập dữ liệu", self._open_issues,
        )
        content.addWidget(third)
        fourth = QFrame()
        fourth.setObjectName("task")
        fourth_layout = QVBoxLayout(fourth)
        fourth_layout.setContentsMargins(18, 15, 18, 15)
        fourth_layout.setSpacing(8)
        row = QHBoxLayout()
        row.addWidget(lbl("04", "status"))
        row.addWidget(lbl("Nhập lên phần mềm quyết toán", "section", wrap=False))
        row.addStretch()
        fourth_layout.addLayout(row)
        fourth_layout.addWidget(lbl("Chọn đúng tác vụ cần chạy trên phần mềm quyết toán."))
        actions = QHBoxLayout()
        actions.addWidget(btn("Tạo số quyết toán mới", self._create_sqt))
        actions.addWidget(btn("Nhập thông tin", lambda: self._select_sqt("nhập thông tin")))
        actions.addWidget(btn("Nhập khoản chi", lambda: self._select_sqt("nhập khoản chi")))
        actions.addStretch()
        fourth_layout.addLayout(actions)
        self.state_4 = lbl("")
        fourth_layout.addWidget(self.state_4)
        content.addWidget(fourth)
        content.addStretch()
        return scroll

    def _build_settings(self) -> QScrollArea:
        scroll, content = self._new_page()
        content.addWidget(lbl("Cài đặt", "title"))
        content.addWidget(lbl("Các đường dẫn ở đây chỉ được lưu trong phiên demo; ứng dụng thật không bị thay đổi."))
        frame = QFrame()
        frame.setObjectName("task")
        layout = QVBoxLayout(frame)
        layout.setContentsMargins(18, 17, 18, 17)
        form = QFormLayout()
        form.setSpacing(10)
        self.setting_edits: dict[str, QLineEdit] = {}
        for caption, key, default in (
            ("Trợ lý quyết toán", "assistant", "Launcher/Mo_Tro_Ly_Quyet_Toan.bat"),
            ("Tạo số QT mới", "create", "Launcher/run_create_new_quyet_toan.bat"),
            ("Nhập thông tin", "information", "Launcher/run_input_information.bat"),
            ("Nhập khoản chi", "expense", "Launcher/run_input_expense.bat"),
            ("Thư mục dữ liệu", "output", "output"),
            ("File theo dõi hàng ngày", "daily", "output/quyet_toan_hang_ngay.xlsx"),
        ):
            line = QLineEdit(default)
            self.setting_edits[key] = line
            row = QHBoxLayout()
            row.addWidget(line, stretch=1)
            row.addWidget(btn("Chọn…", lambda _checked=False, k=key: self._browse_setting(k)))
            form.addRow(caption, row)
        layout.addLayout(form)
        actions = QHBoxLayout()
        actions.addWidget(btn("Kiểm tra", self._check_settings))
        actions.addWidget(btn("Lưu trong demo", self._save_settings, primary=True))
        actions.addStretch()
        layout.addLayout(actions)
        content.addWidget(frame)
        content.addWidget(lbl("Nhật ký thao tác demo", "section"))
        self.log = QPlainTextEdit()
        self.log.setReadOnly(True)
        self.log.setMaximumHeight(145)
        content.addWidget(self.log)
        content.addWidget(btn("Đặt lại dữ liệu mẫu", self._reset_demo), alignment=Qt.AlignLeft)
        content.addStretch()
        return scroll

    def _refresh(self) -> None:
        self.state_1.setText("Dữ liệu mẫu đã sẵn sàng")
        self.state_2.setText(f"{len(self.data.docs)} chứng từ · {sum(d.needs_review for d in self.data.docs.values())} cần kiểm tra")
        if self.data.imported:
            self.state_3.setText("Đã cập nhật bảng theo dõi mẫu")
        else:
            self.state_3.setText(f"{self.data.pending_count()} vấn đề chưa xử lý")
        self.state_4.setText(self.data.last_automation)

    def _log(self, text: str) -> None:
        self.log.appendPlainText(text)

    def _open_assistant(self) -> None:
        self._log("Đã mở trợ lý (mô phỏng).")
        QMessageBox.information(
            self, "Trợ lý quyết toán",
            "Bản demo đã chuẩn bị sẵn một lô chứng từ mẫu. Bấm “Kiểm tra chứng từ” ở bước 2 để xem."
        )

    def _open_review(self) -> None:
        ReviewDialog(self.data, self).exec()
        self._refresh()
        self._log("Đã mở bảng kiểm tra chứng từ mẫu.")

    def _open_issues(self) -> None:
        IssueDialog(self.data, self).exec()
        self._refresh()
        self._log("Đã mở cửa sổ xử lý dữ liệu mẫu.")
        if self.data.imported:
            QMessageBox.information(
                self, "Hoàn tất bản demo",
                "Đã cập nhật bảng theo dõi mẫu trong bộ nhớ. Không có file Excel thật nào thay đổi."
            )

    def _create_sqt(self) -> None:
        self.data.last_automation = "Đã mô phỏng tạo số QT mới."
        self._refresh()
        self._log(self.data.last_automation)
        QMessageBox.information(self, "Tạo số QT mới", "Đây là thao tác mô phỏng. Không chạy luồng tự động hóa thật.")

    def _select_sqt(self, action: str) -> None:
        dialog = SelectSqtDialog(action, self)
        if dialog.exec() != QDialog.Accepted:
            return
        self.data.last_automation = f"Đã mô phỏng {action} cho {len(dialog.selected)} số quyết toán."
        self._refresh()
        self._log(self.data.last_automation)
        QMessageBox.information(
            self, "Đã chọn số quyết toán",
            f"{len(dialog.selected)} số quyết toán đã được chọn trong demo. Không chạy luồng tự động hóa thật."
        )

    def _browse_setting(self, key: str) -> None:
        if key == "output":
            path = QFileDialog.getExistingDirectory(self, "Chọn thư mục dữ liệu")
        else:
            path, _selected_filter = QFileDialog.getOpenFileName(self, "Chọn file")
        if path:
            self.setting_edits[key].setText(path)

    def _check_settings(self) -> None:
        lines = []
        for key, edit in self.setting_edits.items():
            value = edit.text().strip()
            exists = Path(value).is_dir() if key == "output" else Path(value).is_file()
            lines.append(f"{key}: {'Tìm thấy' if exists else 'Chưa tìm thấy'}")
        QMessageBox.information(self, "Kiểm tra đường dẫn demo", "\n".join(lines))

    def _save_settings(self) -> None:
        self._log("Đã lưu các đường dẫn trong phiên demo.")
        QMessageBox.information(self, "Đã lưu trong demo", "Các đường dẫn chỉ được giữ đến khi đóng bản demo. Cấu hình ứng dụng thật không thay đổi.")

    def _reset_demo(self) -> None:
        if QMessageBox.question(
            self, "Đặt lại dữ liệu mẫu", "Đưa bản demo về trạng thái ban đầu để thử lại?"
        ) != QMessageBox.Yes:
            return
        self.data = DemoData()
        self._refresh()
        self._log("Đã đặt lại dữ liệu mẫu.")
        self.nav.setCurrentRow(0)


def main() -> int:
    app = QApplication(sys.argv)
    app.setApplicationName("Trợ lý quyết toán · Demo")
    app.setStyleSheet(STYLE)
    icon = QIcon(str(Path(__file__).resolve().parent / "Assets" / "KIKAI_PM.ico"))
    if not icon.isNull():
        app.setWindowIcon(icon)
    window = DemoWindow()
    window.show()
    return app.exec()


if __name__ == "__main__":
    sys.exit(main())
