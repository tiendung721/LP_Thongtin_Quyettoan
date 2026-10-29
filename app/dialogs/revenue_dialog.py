"""Review, edit, and track revenue without opening the daily workbook."""

from __future__ import annotations

from collections import defaultdict
from typing import Any

from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QCheckBox, QComboBox, QDialog, QFormLayout, QGridLayout, QHBoxLayout,
    QLabel, QLineEdit, QPushButton, QTableWidget, QTableWidgetItem, QVBoxLayout,
    QWidget,
)

from ..revenue import RevenueStore, review_errors
from .notifications import notify


def _money(value: int | None) -> str:
    return f"{value:,.0f}".replace(",", ".") + " đ" if value is not None else "—"


def _import_text(rows: list[dict[str, Any]]) -> str:
    done = sum(row["import_status"] == "DA_NHAP" for row in rows)
    if done == len(rows):
        return "Đã nhập"
    return f"Nhập một phần ({done}/{len(rows)})" if done else "Chưa nhập"


class RevenueDialog(QDialog):
    """One screen for revenue verification and manually maintained import state."""

    def __init__(self, store: RevenueStore, parent=None):
        super().__init__(parent)
        self.store = store
        self.rows: list[dict[str, Any]] = []
        self.groups: dict[tuple[str, str], list[dict[str, Any]]] = {}
        self.visible_groups: list[tuple[str, str]] = []
        self.visible_lines: list[dict[str, Any]] = []
        self.selected_ids: set[int] = set()
        self.current_group: tuple[str, str] | None = None
        self.current_id: int | None = None
        self.dirty = False
        self._loading = False
        self._last_filters: tuple[str, Any, Any, Any] = ("", "ALL", "ALL", "ALL")

        self.setWindowTitle("Chọn số quyết toán nhập doanh thu")
        self.resize(1390, 800)
        self.setMinimumSize(1010, 600)
        root = QVBoxLayout(self)
        heading = QHBoxLayout()
        self.title = QLabel("Chọn số quyết toán · Nhập doanh thu")
        self.title.setObjectName("sqtDialogTitle")
        self.count = QLabel()
        heading.addWidget(self.title)
        heading.addWidget(self.count)
        heading.addStretch(1)
        root.addLayout(heading)

        filters = QGridLayout()
        filters.addWidget(QLabel("Tìm SQT, Container hoặc đối tác"), 0, 0)
        filters.addWidget(QLabel("Đối tác"), 0, 1)
        filters.addWidget(QLabel("Trạng thái kiểm tra"), 0, 2)
        filters.addWidget(QLabel("Tiến độ nhập"), 0, 3)
        self.search = QLineEdit()
        self.search.setPlaceholderText("Nhập SQT, Container hoặc đối tác...")
        self.partner_filter = QComboBox()
        self.check_filter = QComboBox()
        self.check_filter.addItem("Tất cả", "ALL")
        self.check_filter.addItem("Cần kiểm tra", "CAN_KIEM_TRA")
        self.check_filter.addItem("OK", "OK")
        self.import_filter = QComboBox()
        self.import_filter.addItem("Tất cả", "ALL")
        self.import_filter.addItem("Chưa nhập", "CHUA_NHAP")
        self.import_filter.addItem("Đã nhập", "DA_NHAP")
        filters.addWidget(self.search, 1, 0)
        filters.addWidget(self.partner_filter, 1, 1)
        filters.addWidget(self.check_filter, 1, 2)
        filters.addWidget(self.import_filter, 1, 3)
        self.select_visible = QPushButton("Chọn dòng OK hiển thị")
        self.clear_selection = QPushButton("Bỏ chọn")
        filters.addWidget(self.select_visible, 2, 2)
        filters.addWidget(self.clear_selection, 2, 3)
        root.addLayout(filters)

        content = QHBoxLayout()
        self.group_table = QTableWidget(0, 8)
        self.group_table.setHorizontalHeaderLabels([
            "Chọn", "SQT", "Số dòng", "Số Cont", "Đối tác", "Tiền gồm VAT",
            "Tiến độ nhập", "Trạng thái kiểm tra",
        ])
        self.group_table.setSelectionBehavior(QTableWidget.SelectRows)
        self.group_table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.group_table.setAlternatingRowColors(True)
        self.group_table.horizontalHeader().setStretchLastSection(True)
        self.group_table.setColumnWidth(0, 58)
        self.group_table.setColumnWidth(1, 85)
        self.group_table.setColumnWidth(2, 75)
        self.group_table.setColumnWidth(3, 180)
        self.group_table.setColumnWidth(4, 140)
        self.group_table.setColumnWidth(5, 140)
        self.group_table.setColumnWidth(6, 145)
        content.addWidget(self.group_table, 3)

        side = QWidget()
        side_layout = QVBoxLayout(side)
        self.group_detail = QLabel("Chọn một SQT để xem các container.")
        self.group_detail.setWordWrap(True)
        side_layout.addWidget(self.group_detail)
        self.line_table = QTableWidget(0, 5)
        self.line_table.setHorizontalHeaderLabels([
            "Chọn", "Container", "Tiền VAT", "Tiến độ", "Kiểm tra",
        ])
        self.line_table.setSelectionBehavior(QTableWidget.SelectRows)
        self.line_table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.line_table.horizontalHeader().setStretchLastSection(True)
        self.line_table.setMinimumHeight(175)
        side_layout.addWidget(self.line_table, 2)

        form = QFormLayout()
        self.partner_edit = QLineEdit()
        self.sqt_edit = QLineEdit()
        self.container_edit = QLineEdit()
        self.amount_edit = QLineEdit()
        self.import_combo = QComboBox()
        self.import_combo.addItem("Chưa nhập", "CHUA_NHAP")
        self.import_combo.addItem("Đã nhập", "DA_NHAP")
        self.warning_resolved = QCheckBox("Đã đối chiếu và xử lý cảnh báo nguồn")
        form.addRow("Đối tác", self.partner_edit)
        form.addRow("SQT", self.sqt_edit)
        form.addRow("Số Container", self.container_edit)
        form.addRow("Tiền gồm VAT (VND)", self.amount_edit)
        form.addRow("Tiến độ nhập", self.import_combo)
        form.addRow("", self.warning_resolved)
        side_layout.addLayout(form)
        self.source_note = QLabel()
        self.source_note.setWordWrap(True)
        side_layout.addWidget(self.source_note)
        self.warning_note = QLabel()
        self.warning_note.setWordWrap(True)
        side_layout.addWidget(self.warning_note)
        self.save_button = QPushButton("Lưu")
        self.save_button.setProperty("variant", "primary")
        self.save_button.clicked.connect(self.save_current)
        side_layout.addWidget(self.save_button)
        content.addWidget(side, 2)
        root.addLayout(content, 1)

        footer = QHBoxLayout()
        self.summary = QLabel()
        footer.addWidget(self.summary, 1)
        self.run_button = QPushButton("Chạy nhập doanh thu")
        self.run_button.setProperty("variant", "primary")
        self.run_button.setToolTip("PAD doanh thu sẽ được kết nối ở giai đoạn sau.")
        self.run_button.clicked.connect(self._run_later)
        close = QPushButton("Đóng")
        close.clicked.connect(self.reject)
        footer.addWidget(self.run_button)
        footer.addWidget(close)
        root.addLayout(footer)

        self.search.textChanged.connect(self._filter_changed)
        self.partner_filter.currentIndexChanged.connect(self._filter_changed)
        self.check_filter.currentIndexChanged.connect(self._filter_changed)
        self.import_filter.currentIndexChanged.connect(self._filter_changed)
        self.select_visible.clicked.connect(self._select_visible)
        self.clear_selection.clicked.connect(self._clear_selected)
        self.group_table.itemSelectionChanged.connect(self._on_group_selected)
        self.group_table.itemChanged.connect(self._on_group_checked)
        self.line_table.itemSelectionChanged.connect(self._on_line_selected)
        self.line_table.itemChanged.connect(self._on_line_checked)
        for edit in (self.partner_edit, self.sqt_edit, self.container_edit, self.amount_edit):
            edit.textChanged.connect(self._mark_dirty)
        self.import_combo.currentIndexChanged.connect(self._mark_dirty)
        self.warning_resolved.stateChanged.connect(self._mark_dirty)
        self._reload()
        self._last_filters = self._filter_values()

    def _filter_values(self) -> tuple[str, Any, Any, Any]:
        return (self.search.text(), self.partner_filter.currentData(),
                self.check_filter.currentData(), self.import_filter.currentData())

    def _filter_changed(self, *_args) -> None:
        if self.dirty and not self._discard_dirty():
            values = self._last_filters
            for widget in (self.search, self.partner_filter, self.check_filter, self.import_filter):
                widget.blockSignals(True)
            self.search.setText(values[0])
            for widget, value in zip(
                (self.partner_filter, self.check_filter, self.import_filter), values[1:]
            ):
                index = widget.findData(value)
                widget.setCurrentIndex(index if index >= 0 else 0)
            for widget in (self.search, self.partner_filter, self.check_filter, self.import_filter):
                widget.blockSignals(False)
            return
        if self.dirty:
            current = next((row for row in self.rows if row["id"] == self.current_id), None)
            self._load_line(current)
        self._last_filters = self._filter_values()
        self._render_groups()

    def _select_visible(self) -> None:
        for key in self.visible_groups:
            self.selected_ids.update(
                row["id"] for row in self.groups[key] if self._eligible(row)
            )
        self._render_groups()

    def _clear_selected(self) -> None:
        self.selected_ids.clear()
        self._render_groups()

    @staticmethod
    def _item(text: str) -> QTableWidgetItem:
        item = QTableWidgetItem(text)
        item.setFlags(item.flags() & ~Qt.ItemIsEditable)
        return item

    def _reload(self, preferred_id: int | None = None) -> None:
        self.rows = self.store.list_rows()
        groups: dict[tuple[str, str], list[dict[str, Any]]] = defaultdict(list)
        for row in self.rows:
            groups[(row["partner"], row["sqt"])].append(row)
        self.groups = dict(groups)
        old_partner = self.partner_filter.currentData()
        self.partner_filter.blockSignals(True)
        self.partner_filter.clear()
        self.partner_filter.addItem("Tất cả đối tác", "ALL")
        for partner in sorted({row["partner"] for row in self.rows}, key=str.casefold):
            self.partner_filter.addItem(partner or "(Thiếu đối tác)", partner)
        index = self.partner_filter.findData(old_partner)
        self.partner_filter.setCurrentIndex(index if index >= 0 else 0)
        self.partner_filter.blockSignals(False)
        self.count.setText(f"{len(self.rows)} dòng doanh thu")
        self._render_groups(preferred_id=preferred_id)

    def _render_groups(self, *_args, preferred_id: int | None = None) -> None:
        query = self.search.text().strip().casefold()
        partner_filter = self.partner_filter.currentData()
        check_filter = self.check_filter.currentData()
        import_filter = self.import_filter.currentData()
        keys = []
        for key, rows in self.groups.items():
            if partner_filter not in (None, "ALL", key[0]):
                continue
            if check_filter != "ALL" and not any(row["review_status"] == check_filter for row in rows):
                continue
            if import_filter != "ALL" and not any(row["import_status"] == import_filter for row in rows):
                continue
            if query and not any(query in " ".join((row["partner"], row["sqt"], row["container"])).casefold() for row in rows):
                continue
            keys.append(key)
        keys.sort(key=lambda key: (key[0].casefold(), key[1]))
        self.visible_groups = keys
        self.group_table.blockSignals(True)
        self.group_table.setRowCount(len(keys))
        preferred_group = next((key for key, rows in self.groups.items() if any(row["id"] == preferred_id for row in rows)), None)
        for index, key in enumerate(keys):
            rows = self.groups[key]
            eligible = [row for row in rows if self._eligible(row)]
            checked = sum(row["id"] in self.selected_ids for row in eligible)
            checkbox = self._item("")
            if eligible:
                checkbox.setFlags(checkbox.flags() | Qt.ItemIsUserCheckable)
            checkbox.setCheckState(Qt.Checked if eligible and checked == len(eligible) else Qt.PartiallyChecked if checked else Qt.Unchecked)
            checkbox.setToolTip(f"{len(eligible)} dòng OK và chưa nhập có thể chọn")
            self.group_table.setItem(index, 0, checkbox)
            values = [key[1] or "—", str(len(rows)), ", ".join(row["container"] for row in rows[:2]),
                      key[0] or "—", _money(sum(row["amount_vnd"] or 0 for row in rows)),
                      _import_text(rows), "OK" if all(row["review_status"] == "OK" for row in rows) else "Cần kiểm tra"]
            for col, value in enumerate(values, 1):
                self.group_table.setItem(index, col, self._item(value))
        self.group_table.blockSignals(False)
        wanted = preferred_group or self.current_group
        selected = keys.index(wanted) if wanted in keys else (0 if keys else -1)
        if selected >= 0:
            self.group_table.selectRow(selected)
        else:
            self.current_group = None
            self.visible_lines = []
            self.line_table.setRowCount(0)
            self._load_line(None)
        self._update_summary()

    @staticmethod
    def _eligible(row: dict[str, Any]) -> bool:
        return row["review_status"] == "OK" and row["import_status"] == "CHUA_NHAP"

    def _on_group_checked(self, item: QTableWidgetItem) -> None:
        if item.column() != 0 or item.row() >= len(self.visible_groups):
            return
        rows = self.groups[self.visible_groups[item.row()]]
        ids = {row["id"] for row in rows if self._eligible(row)}
        if item.checkState() == Qt.Checked:
            self.selected_ids.update(ids)
        else:
            self.selected_ids.difference_update(ids)
        self._render_lines()
        self._update_summary()

    def _on_group_selected(self) -> None:
        index = self.group_table.currentRow()
        if not 0 <= index < len(self.visible_groups):
            return
        key = self.visible_groups[index]
        if key != self.current_group and not self._discard_dirty():
            self.group_table.blockSignals(True)
            if self.current_group in self.visible_groups:
                self.group_table.selectRow(self.visible_groups.index(self.current_group))
            self.group_table.blockSignals(False)
            return
        self.current_group = key
        self.group_detail.setText(f"SQT {key[1] or '—'} · {key[0] or 'Chưa rõ đối tác'}")
        self._render_lines()

    def _render_lines(self, preferred_id: int | None = None) -> None:
        rows = self.groups.get(self.current_group, []) if self.current_group else []
        self.visible_lines = rows
        self.line_table.blockSignals(True)
        self.line_table.setRowCount(len(rows))
        for index, row in enumerate(rows):
            item = self._item("")
            if self._eligible(row):
                item.setFlags(item.flags() | Qt.ItemIsUserCheckable)
            item.setCheckState(Qt.Checked if row["id"] in self.selected_ids else Qt.Unchecked)
            item.setToolTip("Chỉ chọn được dòng OK và Chưa nhập" if not self._eligible(row) else "")
            self.line_table.setItem(index, 0, item)
            values = (row["container"] or "—", _money(row["amount_vnd"]),
                      "Đã nhập" if row["import_status"] == "DA_NHAP" else "Chưa nhập",
                      "OK" if row["review_status"] == "OK" else "Cần kiểm tra")
            for col, value in enumerate(values, 1):
                self.line_table.setItem(index, col, self._item(value))
        self.line_table.blockSignals(False)
        selected = next((index for index, row in enumerate(rows) if row["id"] == preferred_id), None)
        if selected is None:
            selected = next((index for index, row in enumerate(rows) if row["id"] == self.current_id), 0)
        if rows:
            self.line_table.selectRow(selected)
        else:
            self._load_line(None)

    def _on_line_checked(self, item: QTableWidgetItem) -> None:
        if item.column() != 0 or item.row() >= len(self.visible_lines):
            return
        row = self.visible_lines[item.row()]
        if item.checkState() == Qt.Checked and self._eligible(row):
            self.selected_ids.add(row["id"])
        else:
            self.selected_ids.discard(row["id"])
        self._render_groups()

    def _on_line_selected(self) -> None:
        index = self.line_table.currentRow()
        if not 0 <= index < len(self.visible_lines):
            return
        row = self.visible_lines[index]
        if row["id"] != self.current_id and not self._discard_dirty():
            self.line_table.blockSignals(True)
            prior = next((i for i, item in enumerate(self.visible_lines) if item["id"] == self.current_id), None)
            if prior is not None:
                self.line_table.selectRow(prior)
            self.line_table.blockSignals(False)
            return
        self._load_line(row)

    def _load_line(self, row: dict[str, Any] | None) -> None:
        self._loading = True
        self.current_id = row["id"] if row else None
        for widget, value in ((self.partner_edit, row["partner"] if row else ""),
                              (self.sqt_edit, row["sqt"] if row else ""),
                              (self.container_edit, row["container"] if row else ""),
                              (self.amount_edit, str(row["amount_vnd"]) if row and row["amount_vnd"] is not None else "")):
            widget.setText(value)
            widget.setEnabled(row is not None)
        self.import_combo.setCurrentIndex(self.import_combo.findData(row["import_status"]) if row else 0)
        self.import_combo.setEnabled(row is not None)
        self.warning_resolved.setChecked(bool(row["warning_resolved"]) if row else False)
        self.warning_resolved.setEnabled(row is not None)
        self.save_button.setEnabled(row is not None)
        if row:
            where = ", ".join(part for part in (
                row["source_sheet"] and f"sheet {row['source_sheet']}",
                row["source_page"] and f"trang {row['source_page']}",
                row["source_line"] and f"dòng {row['source_line']}",
            ) if part)
            self.source_note.setText(f"Nguồn: {row['source_name']}" + (f" · {where}" if where else ""))
            import json
            warnings = json.loads(row["warning_json"])
            validation = review_errors(
                row["partner"], row["sqt"], row["container"], row["amount_vnd"]
            )
            notices = validation + [str(item) for item in warnings]
            self.warning_note.setText(
                "Cần kiểm tra: " + "; ".join(notices)
                if notices else "Không có cảnh báo dữ liệu."
            )
        else:
            self.source_note.clear()
            self.warning_note.clear()
        self.dirty = False
        self._loading = False

    def _mark_dirty(self, *_args) -> None:
        if not self._loading and self.current_id is not None:
            self.dirty = True

    def _discard_dirty(self) -> bool:
        if not self.dirty:
            return True
        return notify.confirm(
            self, "Chưa lưu chỉnh sửa", "Bạn đã sửa dòng doanh thu nhưng chưa bấm Lưu. Bỏ thay đổi?",
            confirm_text="Bỏ thay đổi", cancel_text="Tiếp tục sửa",
        )

    def save_current(self) -> None:
        if self.current_id is None:
            return
        try:
            saved = self.store.save_row(
                self.current_id, partner=self.partner_edit.text(), sqt=self.sqt_edit.text(),
                container=self.container_edit.text(), amount=self.amount_edit.text(),
                warning_resolved=self.warning_resolved.isChecked(),
                import_status=self.import_combo.currentData(),
            )
        except (ValueError, OSError) as exc:
            notify.error(self, "Không lưu được doanh thu", str(exc))
            return
        self.dirty = False
        if saved["review_status"] != "OK" or saved["import_status"] == "DA_NHAP":
            self.selected_ids.discard(self.current_id)
        self._reload(preferred_id=self.current_id)

    def _update_summary(self) -> None:
        valid = {row["id"] for row in self.rows if self._eligible(row)}
        self.selected_ids.intersection_update(valid)
        need = sum(row["review_status"] != "OK" for row in self.rows)
        self.summary.setText(f"{len(self.rows)} dòng · {need} cần kiểm tra · Đã chọn {len(self.selected_ids)} dòng OK")
        # The button remains visible; the actual PAD flow is a separate phase.
        self.run_button.setEnabled(bool(self.selected_ids))

    def _run_later(self) -> None:
        if self.dirty:
            notify.warning(self, "Chưa lưu", "Hãy bấm Lưu trước khi nhập doanh thu.")
            return
        notify.info(self, "PAD doanh thu chưa cấu hình", "Đã chọn dữ liệu hợp lệ. Flow PAD nhập doanh thu sẽ được nối ở giai đoạn sau; chưa có dòng nào được gửi hoặc đổi trạng thái nhập.")

    def reject(self) -> None:  # noqa: N802
        if self._discard_dirty():
            super().reject()
