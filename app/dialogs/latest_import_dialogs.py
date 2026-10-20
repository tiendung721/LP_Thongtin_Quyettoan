"""Read-only views of the latest Excel import and PAD handoff snapshots."""

from __future__ import annotations

from collections import defaultdict
from datetime import datetime
from typing import Any

from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QComboBox, QHBoxLayout, QLabel, QLineEdit, QPushButton,
    QTableWidget, QTableWidgetItem, QVBoxLayout,
)

from .base_dialog import ResizableDialog
from ..daily_import import parse_money
from ..services.sqt_selection_service import EXPENSE_SELECTION_OPERATION


def _text(value: Any) -> str:
    if value in (None, ""):
        return "—"
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value)


def _time(value: str | None) -> str:
    try:
        return datetime.fromisoformat(value or "").strftime("%H:%M:%S %d/%m/%Y")
    except ValueError:
        return value or "—"


def _money(value: Any) -> str:
    try:
        return f"{float(value):,.0f}".replace(",", ".")
    except (TypeError, ValueError):
        return "—"


class _HistoryDialog(ResizableDialog):
    def __init__(self, title: str, description: str, headers: list[str], parent=None):
        super().__init__(parent)
        self.setWindowTitle(title)
        self.resize(1500, 800)
        self.setMinimumSize(950, 560)
        layout = QVBoxLayout(self)
        layout.setContentsMargins(14, 14, 14, 14)
        heading = QLabel(description)
        heading.setWordWrap(True)
        layout.addWidget(heading)
        controls = QHBoxLayout()
        controls.addWidget(QLabel("Lọc kết quả:"))
        self.filter_combo = QComboBox()
        self.filter_combo.addItem("Tất cả")
        controls.addWidget(self.filter_combo)
        controls.addWidget(QLabel("Tìm:"))
        self.search = QLineEdit()
        self.search.setPlaceholderText("Container, SQT hoặc nội dung...")
        self.search.setClearButtonEnabled(True)
        controls.addWidget(self.search, 1)
        layout.addLayout(controls)
        self.table = QTableWidget(0, len(headers))
        self.table.setHorizontalHeaderLabels(headers)
        self.table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.table.setAlternatingRowColors(True)
        self.table.setSelectionBehavior(QTableWidget.SelectRows)
        self.table.setWordWrap(False)
        layout.addWidget(self.table, 1)
        footer = QHBoxLayout()
        footer.addStretch(1)
        close = QPushButton("Đóng")
        close.clicked.connect(self.accept)
        footer.addWidget(close)
        layout.addLayout(footer)
        self._rows: list[list[str]] = []
        self._statuses: list[str] = []
        self.filter_combo.currentTextChanged.connect(self._refresh)
        self.search.textChanged.connect(self._refresh)

    def _set_rows(self, rows: list[list[str]], statuses: list[str] | None = None) -> None:
        self._rows = rows
        self._statuses = statuses or [""] * len(rows)
        for status in dict.fromkeys(self._statuses):
            if status:
                self.filter_combo.addItem(status)
        self._refresh()

    def _refresh(self) -> None:
        status = self.filter_combo.currentText()
        query = self.search.text().strip().casefold()
        visible = [
            row for row, row_status in zip(self._rows, self._statuses)
            if (status == "Tất cả" or row_status == status)
            and (not query or query in " ".join(row).casefold())
        ]
        self.table.setRowCount(len(visible))
        for r, row in enumerate(visible):
            for c, value in enumerate(row):
                item = QTableWidgetItem(value)
                item.setToolTip(value)
                item.setFlags(Qt.ItemIsEnabled | Qt.ItemIsSelectable)
                self.table.setItem(r, c, item)
        self.table.resizeColumnsToContents()
        self.table.horizontalHeader().setStretchLastSection(True)


class LatestDailyImportDialog(_HistoryDialog):
    def __init__(self, run: dict[str, Any], parent=None):
        details = run.get("details")
        summary = run.get("summary") or {}
        description = (
            f"Bước 3 · {_time(run.get('completed_at') or run.get('created_at'))} · "
            f"{('Lỗi: ' + str(run.get('error_message'))) if run.get('status') == 'FAILED' else 'Đã ghi Excel' if summary.get('wrote_workbook') else 'Không thay đổi Excel'}\n"
            f"{_text(run.get('daily_path'))}"
        )
        if details is None:
            description += "\nLượt này được ghi bằng phiên bản cũ nên không có bảng chi tiết."
        super().__init__(
            "Chi tiết lượt cập nhật bảng theo dõi gần nhất", description,
            ["Container", "SQT", "B/L", "Loại dữ liệu", "Nguồn", "Sheet đích",
             "Ô / dòng đích", "Giá trị trước", "Giá trị sau", "Kết quả"],
            parent,
        )
        rows = []
        statuses = []
        for item in details or []:
            result = _text(item.get("result"))
            rows.append([
                _text(item.get("container")), _text(item.get("sqt")),
                _text(item.get("bill")), _text(item.get("field")),
                _text(item.get("source")), _text(item.get("sheet")),
                _text(item.get("cell")), _text(item.get("before")),
                _text(item.get("after")), result,
            ])
            statuses.append(result)
        self._set_rows(rows, statuses)


class LatestRpaLaunchDialog(_HistoryDialog):
    def __init__(self, run: dict[str, Any], parent=None):
        expense = run.get("operation") == EXPENSE_SELECTION_OPERATION
        action = "Nhập khoản chi" if expense else "Nhập thông tin"
        description = (
            f"Bước 4 · {action} · Đã gửi cho PAD lúc {_time(run.get('created_at'))}\n"
            f"Sheet {run.get('sheet_name') or '—'} · "
            f"{len(run.get('selected_sqt') or [])} SQT · "
            "Trạng thái và giá trị dưới đây là ảnh chụp trước khi chạy PAD; "
            "chưa xác nhận PAD nhập thành công."
        )
        headers = (
            ["SQT", "Trạng thái trước khi gửi", "Số khoản chi", "Giá vật liệu",
             "Thành tiền", "VAT", "Tổng tiền"]
            if expense else
            ["SQT", "Trạng thái trước khi gửi", "Sheet", "Dòng Excel",
             "Container", "Số B/L", "Ngày Đóng", "Loại hàng", "Số tấn",
             "Tên tàu", "VT biển"]
        )
        super().__init__("Dữ liệu của lần gửi PAD gần nhất", description, headers, parent)
        snapshot = run.get("snapshot") or []
        if expense:
            groups: dict[str, list[dict[str, Any]]] = defaultdict(list)
            for item in snapshot:
                groups[str(item.get("sqt") or "")].append(item.get("values") or {})
            rows = []
            statuses = []
            for sqt in run.get("selected_sqt") or []:
                values = groups.get(str(sqt), [])
                state_values = {str(v.get("Trạng thái nhập") or "Chưa nhập") for v in values}
                state = next(iter(state_values)) if len(state_values) == 1 else "Nhập một phần"
                rows.append([
                    str(sqt), state, str(len(values)),
                    *(_money(sum(parse_money(v.get(field)) or 0 for v in values))
                      for field in ("Giá vật liệu", "Thành tiền", "VAT", "Tổng tiền")),
                ])
                statuses.append(state)
        else:
            rows = []
            statuses = []
            for item in snapshot:
                values = item.get("values") or {}
                state = _text(values.get("Trạng thái nhập"))
                rows.append([
                    _text(item.get("sqt")), state,
                    _text(item.get("sheet_name")), _text(item.get("row_number")),
                    _text(values.get("Số Container")), _text(values.get("Số B/L")),
                    _text(values.get("Ngày Đóng")), _text(values.get("Loại hàng")),
                    _text(values.get("Số tấn")), _text(values.get("Tên tàu")),
                    _text(values.get("VT biển")),
                ])
                statuses.append(state)
        self._set_rows(rows, statuses)
