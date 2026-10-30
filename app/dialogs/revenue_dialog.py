"""Review revenue in a full-width list; persist all edits with one Save."""

from __future__ import annotations

import json
import os
from typing import Any

from PySide6.QtCore import Qt
from PySide6.QtGui import QColor
from PySide6.QtWidgets import (
    QComboBox, QDialog, QFormLayout, QGridLayout, QHBoxLayout, QLabel,
    QLineEdit, QPushButton, QTableWidget, QTableWidgetItem, QVBoxLayout,
)

from ..revenue import (
    RevenueStore, normalize_container, normalize_sqt, parse_amount_vnd,
    review_errors,
)
from ..revenue_pad import (
    RevenuePadError, build_selection, read_selection, validate_selection, write_selection,
)
from ..services.rpa_launcher import RpaLaunchError, launch_bat
from .base_dialog import ResizableDialog
from .notifications import notify


def _money(value: int | None) -> str:
    return f"{value:,}".replace(",", ".") + " đ" if value is not None else "—"


def _import_text(value: str) -> str:
    return "Đã nhập" if value == "DA_NHAP" else "Chưa nhập"


class RevenueEditDialog(ResizableDialog):
    """Stage one container's changes without writing to SQLite."""

    def __init__(self, row: dict[str, Any], parent=None):
        super().__init__(parent)
        self.row = row
        self.setWindowTitle(f"Sửa doanh thu · SQT {row['sqt'] or '—'}")
        self.resize(650, 500)
        self.setMinimumSize(490, 390)
        layout = QVBoxLayout(self)

        source_parts = [row["source_name"]]
        for key, label in (("source_sheet", "sheet"), ("source_page", "trang"),
                           ("source_line", "dòng")):
            if row.get(key):
                source_parts.append(f"{label} {row[key]}")
        source = QLabel("Nguồn: " + " · ".join(source_parts))
        source.setWordWrap(True)
        layout.addWidget(source)

        form = QFormLayout()
        self.partner_edit = QLineEdit(row["partner"])
        self.sqt_edit = QLineEdit(row["sqt"])
        self.container_edit = QLineEdit(row["container"])
        self.amount_edit = QLineEdit(
            str(row["amount_vnd"]) if row["amount_vnd"] is not None else ""
        )
        self.import_combo = QComboBox()
        self.import_combo.addItem("Chưa nhập", "CHUA_NHAP")
        self.import_combo.addItem("Đã nhập", "DA_NHAP")
        self.import_combo.setCurrentIndex(self.import_combo.findData(row["import_status"]))
        form.addRow("Đối tác (không bắt buộc)", self.partner_edit)
        form.addRow("SQT", self.sqt_edit)
        form.addRow("Số Container", self.container_edit)
        form.addRow("Tiền gồm VAT (VND)", self.amount_edit)
        form.addRow("Tiến độ nhập", self.import_combo)
        layout.addLayout(form)

        errors = review_errors(
            row["partner"], row["sqt"], row["container"], row["amount_vnd"]
        )
        status = QLabel(
            "Trạng thái kiểm tra: " + ("; ".join(errors) if errors else "OK")
        )
        status.setWordWrap(True)
        layout.addWidget(status)
        raw_warnings = json.loads(row["warning_json"])
        if raw_warnings:
            warnings = QLabel("Cảnh báo từ nguồn: " + "; ".join(map(str, raw_warnings)))
            warnings.setWordWrap(True)
            layout.addWidget(warnings)
        layout.addStretch(1)

        buttons = QHBoxLayout()
        buttons.addStretch(1)
        apply_button = QPushButton("Áp dụng vào bảng")
        apply_button.setProperty("variant", "primary")
        apply_button.clicked.connect(self._apply)
        cancel = QPushButton("Hủy")
        cancel.clicked.connect(self.reject)
        buttons.addWidget(apply_button)
        buttons.addWidget(cancel)
        layout.addLayout(buttons)

    def values(self) -> dict[str, Any]:
        amount_text = self.amount_edit.text().strip()
        amount = parse_amount_vnd(amount_text)
        if amount_text and amount is None:
            raise ValueError("Tiền gồm VAT phải là số nguyên VND hợp lệ.")
        return {
            "partner": self.partner_edit.text().strip(),
            "sqt": normalize_sqt(self.sqt_edit.text()),
            "container": normalize_container(self.container_edit.text()),
            "amount_vnd": amount,
            "import_status": self.import_combo.currentData(),
        }

    def _apply(self) -> None:
        try:
            self.values()
        except ValueError as exc:
            notify.warning(self, "Dữ liệu chưa hợp lệ", str(exc))
            return
        self.accept()


class RevenueDialog(ResizableDialog):
    """Full-width revenue table with batch status changes and one atomic Save."""

    COLUMNS = (
        "Chọn", "SQT", "Số Container", "Đối tác", "Tiền gồm VAT",
        "Tiến độ nhập", "Trạng thái kiểm tra", "Nguồn",
    )

    def __init__(
        self, store: RevenueStore, parent=None, *,
        selection_path: str = "", pad_bat_path: str = "", logger=None,
    ):
        super().__init__(parent)
        self.store = store
        self.selection_path = selection_path
        self.pad_bat_path = pad_bat_path
        self.logger = logger
        self.base_rows: dict[int, dict[str, Any]] = {}
        self.rows: dict[int, dict[str, Any]] = {}
        self.visible_ids: list[int] = []
        self.selected_ids: set[int] = set()
        self.drafts: dict[int, dict[str, Any]] = {}

        self.setWindowTitle("Chọn số quyết toán nhập doanh thu")
        self.resize(1390, 800)
        self.setMinimumSize(950, 550)
        root = QVBoxLayout(self)
        heading = QHBoxLayout()
        title = QLabel("Chọn số quyết toán · Nhập doanh thu")
        title.setObjectName("sqtDialogTitle")
        self.count = QLabel()
        heading.addWidget(title)
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
        self.select_visible = QPushButton("Chọn hiển thị")
        self.clear_selection = QPushButton("Bỏ chọn")
        filters.addWidget(self.select_visible, 2, 2)
        filters.addWidget(self.clear_selection, 2, 3)
        root.addLayout(filters)

        self.table = QTableWidget(0, len(self.COLUMNS))
        self.table.setHorizontalHeaderLabels(self.COLUMNS)
        self.table.setAlternatingRowColors(True)
        self.table.setSelectionBehavior(QTableWidget.SelectRows)
        self.table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.table.verticalHeader().setVisible(False)
        self.table.horizontalHeader().setStretchLastSection(True)
        for column, width in enumerate((60, 95, 180, 170, 155, 145, 170)):
            self.table.setColumnWidth(column, width)
        root.addWidget(self.table, 1)
        hint = QLabel("Bấm một dòng để sửa; các thay đổi chỉ được ghi khi bấm Lưu tất cả.")
        root.addWidget(hint)

        bulk = QHBoxLayout()
        bulk.addWidget(QLabel("Đổi tiến độ nhập cho các dòng đã chọn:"))
        self.bulk_status = QComboBox()
        self.bulk_status.addItem("Chưa nhập", "CHUA_NHAP")
        self.bulk_status.addItem("Đã nhập", "DA_NHAP")
        self.bulk_apply = QPushButton("Áp dụng hàng loạt")
        bulk.addWidget(self.bulk_status)
        bulk.addWidget(self.bulk_apply)
        bulk.addStretch(1)
        root.addLayout(bulk)

        footer = QHBoxLayout()
        self.summary = QLabel()
        footer.addWidget(self.summary, 1)
        self.save_button = QPushButton("Lưu tất cả")
        self.save_button.setProperty("variant", "primary")
        self.run_button = QPushButton("Chạy nhập doanh thu")
        self.run_button.setToolTip("Chọn một dòng để nhập tổng doanh thu của toàn bộ SQT.")
        close = QPushButton("Đóng")
        footer.addWidget(self.save_button)
        footer.addWidget(self.run_button)
        footer.addWidget(close)
        root.addLayout(footer)

        self.search.textChanged.connect(self._render)
        self.partner_filter.currentIndexChanged.connect(self._render)
        self.check_filter.currentIndexChanged.connect(self._render)
        self.import_filter.currentIndexChanged.connect(self._render)
        self.select_visible.clicked.connect(self._select_visible)
        self.clear_selection.clicked.connect(self._clear_selected)
        self.bulk_apply.clicked.connect(self._apply_bulk_status)
        self.table.itemChanged.connect(self._on_checked)
        self.table.cellClicked.connect(self._on_cell_clicked)
        self.save_button.clicked.connect(self.save_all)
        self.run_button.clicked.connect(self._run_revenue)
        close.clicked.connect(self.reject)
        self._reload()

    @staticmethod
    def _item(text: str) -> QTableWidgetItem:
        item = QTableWidgetItem(text)
        item.setFlags(item.flags() & ~Qt.ItemIsEditable)
        return item

    def _reload(self) -> None:
        loaded = self.store.list_rows()
        self.base_rows = {row["id"]: row for row in loaded}
        self.rows = {row["id"]: dict(row) for row in loaded}
        self.drafts.clear()
        selected_partner = self.partner_filter.currentData()
        self.partner_filter.blockSignals(True)
        self.partner_filter.clear()
        self.partner_filter.addItem("Tất cả đối tác", "ALL")
        for partner in sorted({row["partner"] for row in loaded}, key=str.casefold):
            self.partner_filter.addItem(partner or "(Chưa rõ đối tác)", partner)
        index = self.partner_filter.findData(selected_partner)
        self.partner_filter.setCurrentIndex(index if index >= 0 else 0)
        self.partner_filter.blockSignals(False)
        self.count.setText(f"{len(loaded)} dòng doanh thu")
        self.selected_ids.intersection_update(self.rows)
        self._render()

    def _render(self, *_args) -> None:
        query = self.search.text().strip().casefold()
        partner = self.partner_filter.currentData()
        check = self.check_filter.currentData()
        progress = self.import_filter.currentData()
        self.visible_ids = [row_id for row_id, row in self.rows.items()
                            if partner in (None, "ALL", row["partner"])
                            and check in (None, "ALL", row["review_status"])
                            and progress in (None, "ALL", row["import_status"])
                            and (not query or query in " ".join((
                                row["sqt"], row["container"], row["partner"]
                            )).casefold())]
        self.table.blockSignals(True)
        self.table.setRowCount(len(self.visible_ids))
        for index, row_id in enumerate(self.visible_ids):
            row = self.rows[row_id]
            check_item = self._item("")
            check_item.setFlags(check_item.flags() | Qt.ItemIsUserCheckable)
            check_item.setCheckState(Qt.Checked if row_id in self.selected_ids else Qt.Unchecked)
            self.table.setItem(index, 0, check_item)
            values = (
                row["sqt"] or "—", row["container"] or "—", row["partner"] or "—",
                _money(row["amount_vnd"]), _import_text(row["import_status"]),
                "OK" if row["review_status"] == "OK" else "Cần kiểm tra",
                row["source_name"],
            )
            for column, value in enumerate(values, 1):
                item = self._item(value)
                if column == 5:
                    item.setForeground(QColor("#15803D" if row["import_status"] == "DA_NHAP" else "#DC2626"))
                elif column == 6:
                    item.setForeground(QColor("#15803D" if row["review_status"] == "OK" else "#B45309"))
                self.table.setItem(index, column, item)
        self.table.blockSignals(False)
        self._update_summary()

    def _on_checked(self, item: QTableWidgetItem) -> None:
        if item.column() != 0 or item.row() >= len(self.visible_ids):
            return
        row_id = self.visible_ids[item.row()]
        if item.checkState() == Qt.Checked:
            self.selected_ids.add(row_id)
        else:
            self.selected_ids.discard(row_id)
        self._update_summary()

    def _on_cell_clicked(self, index: int, column: int) -> None:
        if column == 0 or index >= len(self.visible_ids):
            return
        row_id = self.visible_ids[index]
        dialog = RevenueEditDialog(self.rows[row_id], self)
        if dialog.exec() != QDialog.Accepted:
            return
        self._stage(row_id, dialog.values())
        self._render()

    def _stage(self, row_id: int, values: dict[str, Any]) -> None:
        row = self.rows[row_id]
        row.update(values)
        row["review_status"] = (
            "CAN_KIEM_TRA" if review_errors(
                row["partner"], row["sqt"], row["container"], row["amount_vnd"]
            ) else "OK"
        )
        keys = ("partner", "sqt", "container", "amount_vnd", "import_status")
        if any(row[key] != self.base_rows[row_id][key] for key in keys):
            self.drafts[row_id] = {
                "id": row_id, "partner": row["partner"], "sqt": row["sqt"],
                "container": row["container"], "amount": row["amount_vnd"],
                "import_status": row["import_status"],
            }
        else:
            self.drafts.pop(row_id, None)

    def _select_visible(self) -> None:
        self.selected_ids.update(self.visible_ids)
        self._render()

    def _clear_selected(self) -> None:
        self.selected_ids.clear()
        self._render()

    def _apply_bulk_status(self) -> None:
        if not self.selected_ids:
            notify.warning(self, "Chưa chọn dòng", "Hãy chọn các dòng cần đổi tiến độ nhập.")
            return
        new_status = self.bulk_status.currentData()
        for row_id in self.selected_ids:
            row = self.rows[row_id]
            self._stage(row_id, {
                "partner": row["partner"], "sqt": row["sqt"],
                "container": row["container"], "amount_vnd": row["amount_vnd"],
                "import_status": new_status,
            })
        self._render()

    def _update_summary(self) -> None:
        needs_review = sum(row["review_status"] != "OK" for row in self.rows.values())
        ready = sum(self.rows[row_id]["review_status"] == "OK"
                    and self.rows[row_id]["import_status"] == "CHUA_NHAP"
                    for row_id in self.selected_ids)
        self.summary.setText(
            f"{len(self.rows)} dòng · {needs_review} cần kiểm tra · "
            f"Đã chọn {len(self.selected_ids)} dòng · {ready} dòng đủ điều kiện nhập"
        )
        self.save_button.setEnabled(bool(self.drafts))
        self.run_button.setEnabled(ready > 0 or bool(
            self.selection_path and os.path.exists(self.selection_path)
        ))

    def save_all(self) -> None:
        if not self.drafts:
            return
        try:
            self.store.save_rows(list(self.drafts.values()))
        except (ValueError, OSError) as exc:
            notify.error(self, "Không lưu được doanh thu", str(exc))
            return
        self._reload()

    def _run_revenue(self) -> None:
        if self.drafts:
            notify.warning(self, "Chưa lưu", "Hãy bấm Lưu tất cả trước khi nhập doanh thu.")
            return
        if not self.selection_path or not self.pad_bat_path or self.logger is None:
            notify.error(self, "Chưa cấu hình PAD doanh thu", "Thiếu đường dẫn chạy PAD doanh thu.")
            return
        created = False
        try:
            if os.path.exists(self.selection_path):
                payload = read_selection(self.selection_path)
                validate_selection(payload, self.store.db_path)
                sqt_list = ", ".join(item["sqt"] for item in payload["items"])
                if not notify.confirm(
                    self, "Tiếp tục lượt PAD doanh thu",
                    f"Lượt trước chưa kết thúc (SQT: {sqt_list}). Chạy tiếp lượt này?",
                    confirm_text="Chạy tiếp",
                ):
                    return
            else:
                payload = build_selection(self.store, self.selected_ids)
                write_selection(self.selection_path, payload)
                created = True
        except (RevenuePadError, OSError, KeyError) as exc:
            notify.error(self, "Chưa gửi được doanh thu", str(exc))
            return
        try:
            launch_bat(
                self.pad_bat_path,
                logger=self.logger,
                description="nhập doanh thu",
                missing_message="Không tìm thấy file BAT nhập doanh thu.",
            )
        except RpaLaunchError as exc:
            if created:
                os.remove(self.selection_path)
            notify.error(self, "Không mở được PAD doanh thu", str(exc))
            return
        notify.info(
            self, "Đã gọi PAD doanh thu",
            f"Đã gửi {len(payload['items'])} SQT. "
            "Trạng thái sẽ đổi thành Đã nhập sau khi PAD kiểm tra giá trị trên web. "
            "Đóng rồi mở lại cửa sổ này để xem trạng thái mới.",
        )

    def reject(self) -> None:  # noqa: N802
        if self.drafts and not notify.confirm(
            self, "Chưa lưu thay đổi", "Bạn đã sửa doanh thu nhưng chưa bấm Lưu tất cả. Bỏ thay đổi?",
            confirm_text="Bỏ thay đổi", cancel_text="Tiếp tục sửa",
        ):
            return
        super().reject()
