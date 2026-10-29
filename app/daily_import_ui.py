"""Các hộp thoại thân thiện cho luồng nhập file theo dõi hàng ngày."""

from __future__ import annotations

import copy
import json
import os
from datetime import date, datetime
from typing import Any, Callable, Dict, List, Optional

from PySide6.QtCore import QObject, QTimer, Qt, Signal, Slot
from PySide6.QtGui import QColor, QKeySequence
from PySide6.QtWidgets import (
    QAbstractItemView,
    QApplication,
    QButtonGroup,
    QComboBox,
    QDialog,
    QDialogButtonBox,
    QFrame,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QLineEdit,
    QMenu,
    QPushButton,
    QRadioButton,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)

from .dialogs.base_dialog import ResizableDialog
from .dialogs.notifications import notify
from .extract_row_editor import ExtractRowEditor, has_low_confidence, needs_review, warning_items, warning_text
from .revenue import is_revenue_row
from .theme import repolish
from .daily_import import (
    DOC_BILL,
    DOC_EXPENSE,
    DOC_SCALE,
    DailyImportError,
    DailyImportService,
    FieldConflict,
    ImportAnalysis,
    ImportCandidate,
    ImportIssue,
    IssueResolution,
    ISSUE_FIELD_CONFLICT,
    ISSUE_MISSING_CLOSE_DATE,
    ISSUE_MULTIPLE_BILLS,
    ISSUE_MULTIPLE_TARGETS,
    ISSUE_NO_TARGET,
    RESOLUTION_FIELDS,
    RESOLUTION_SELECT,
    RESOLUTION_SKIP,
    UnmatchedRow,
    classify_doc_type,
    extract_summary,
    load_extract_payload,
    parse_date,
    parse_money,
    parse_number,
    save_extract_payload,
)

# Dòng lỗi (chưa nhập được) và dòng cần soát lại được tô nền để nhìn ra ngay.
ERROR_BG = QColor("#FEE2E2")
ERROR_FG = QColor("#B91C1C")
WARN_BG = QColor("#FEF3C7")


def _display_date(value: Any) -> str:
    parsed = parse_date(value)
    if not parsed:
        return str(value or "")
    try:
        return datetime.strptime(parsed, "%Y-%m-%d").strftime("%d/%m/%Y")
    except ValueError:
        return str(value or "")


def _display_number(value: Any, decimals: int = 0) -> str:
    """Số theo cách viết Việt Nam: 49.572.000 (tiền) hoặc 27,83 (tấn)."""
    number = parse_number(value) if decimals else parse_money(value)
    if number is None:
        return ""
    text = f"{number:,.{decimals}f}"
    return text.replace(",", "\x00").replace(".", ",").replace("\x00", ".")


def _display_value(value: Any) -> str:
    """Giá trị thô của một trường JSON để đưa vào ô nhập."""
    if value is None:
        return ""
    if isinstance(value, (dict, list)):
        return json.dumps(value, ensure_ascii=False, indent=2)
    return str(value)


def _main_info(doc_type: str, row: Dict[str, Any]) -> str:
    """Thông tin nhận dạng chính của một dòng; thiếu thì để trống."""
    parts: List[str] = []
    if doc_type == DOC_EXPENSE:
        if row.get("so_hd"):
            parts.append(f"HĐ {row['so_hd']}")
        unit_price = _display_number(row.get("don_gia"))
        if unit_price:
            parts.append(f"ĐG {unit_price}")
    elif doc_type == DOC_SCALE:
        if row.get("so_chi_seal"):
            parts.append(f"Seal {row['so_chi_seal']}")
        if row.get("bien_so_xe"):
            parts.append(f"Xe {row['bien_so_xe']}")
    elif doc_type == DOC_BILL:
        if row.get("so_bill"):
            parts.append(f"B/L {row['so_bill']}")
        if row.get("ten_tau"):
            parts.append(str(row["ten_tau"]))
    return "  •  ".join(parts)


def _amount_info(doc_type: str, row: Dict[str, Any]) -> str:
    """Cột “Số tấn / Tổng tiền” tùy theo loại chứng từ."""
    tons = _display_number(row.get("so_tan"), 2)
    if doc_type == DOC_SCALE:
        return f"{tons} tấn" if tons else ""
    if doc_type == DOC_EXPENSE:
        return _display_number(row.get("tong_tien"))
    if doc_type == DOC_BILL:
        return f"{tons} tấn" if tons else _display_number(row.get("tong_tien"))
    return ""


def _warning_count(value: Any) -> int:
    return len(warning_items(value))


def _row_warning_count(row: Dict[str, Any]) -> int:
    return _warning_count(row.get("canh_bao")) + int(has_low_confidence(row))


class FunctionWorker(QObject):
    """Chạy một callable ở QThread và trả kết quả về UI thread."""

    finished = Signal(object)
    failed = Signal(str)

    def __init__(self, function: Callable[[], Any]):
        super().__init__()
        self.function = function

    @Slot()
    def run(self) -> None:
        try:
            self.finished.emit(self.function())
        except Exception as exc:  # noqa: BLE001 - chuyển lỗi sang UI
            self.failed.emit(str(exc))


class _EditExtractRowDialog(ResizableDialog):
    """Cửa sổ sửa chứng từ theo đúng loại giấy tờ."""

    def __init__(self, row: Dict[str, Any], reason: str = "", parent=None):
        super().__init__(parent)
        self.row = row
        self._updated: Dict[str, Any] = dict(row)

        self.setWindowTitle("Sửa dòng bóc tách")
        self.resize(760, 800)
        self.setMinimumSize(600, 500)

        layout = QVBoxLayout(self)

        if reason:
            note = QLabel(f"Dòng này chưa nhập được: {reason}")
            note.setObjectName("noteText")
            note.setWordWrap(True)
            layout.addWidget(note)

        self.editor_form = ExtractRowEditor(row, self)
        layout.addWidget(self.editor_form, stretch=1)

        buttons = QDialogButtonBox(QDialogButtonBox.Save | QDialogButtonBox.Cancel)
        buttons.button(QDialogButtonBox.Save).setText("Xong")
        buttons.button(QDialogButtonBox.Cancel).setText("Hủy")
        buttons.accepted.connect(self._on_accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)

    @Slot()
    def _on_accept(self) -> None:
        try:
            updated = self.editor_form.updated_row()
        except ValueError as exc:
            notify.warning(self, "Dữ liệu chưa hợp lệ", str(exc))
            return
        self._updated = updated
        self.accept()

    def updated_row(self) -> Dict[str, Any]:
        return self._updated


class JsonExtractDialog(ResizableDialog):
    """Bước 2: xem cả lô bóc tách dạng bảng, sửa/xóa từng dòng rồi lưu lại.

    ``error_rows`` là các dòng Bước 3 không ghép được ({vị trí trong
    du_lieu_boc_tach: lý do}); các dòng này được tô đỏ kèm lý do để người dùng
    sửa hoặc xóa ngay tại đây.
    """

    COLUMNS = [
        "STT",
        "Loại chứng từ",
        "File nguồn",
        "Container",
        "Ngày đóng",
        "Loại hàng",
        "Thông tin chung",
        "Số tấn / Tổng tiền",
        "Giá VL",
        "Trạng thái",
    ]

    def __init__(
        self,
        path: str,
        parent=None,
        error_rows: Optional[Dict[int, str]] = None,
    ):
        super().__init__(parent)
        self.path = path
        self.payload: Dict[str, Any] = load_extract_payload(path)
        all_rows = [dict(row) for row in self.payload.get("du_lieu_boc_tach") or []]
        self._revenue_rows = [row for row in all_rows if is_revenue_row(row)]
        regular_indexes = [index for index, row in enumerate(all_rows) if not is_revenue_row(row)]
        self.rows: List[Dict[str, Any]] = [all_rows[index] for index in regular_indexes]
        # Lý do lỗi đi kèm từng dòng, giữ song song với self.rows để khi xóa dòng
        # thì lý do không bị lệch vị trí.
        self.row_errors: List[str] = [
            str((error_rows or {}).get(index, "")) for index in regular_indexes
        ]
        self.saved = False
        self._dirty = False
        self._visible_indices: List[int] = []
        self._sort_column = 0
        self._sort_ascending = True

        self.setWindowTitle("Kiểm tra dữ liệu bóc tách")
        self.resize(1480, 720)
        self.setMinimumSize(900, 480)

        layout = QVBoxLayout(self)

        title = QLabel(f"Đang xem: {os.path.basename(path)}")
        title.setToolTip(path)
        title.setObjectName("metaText")
        title.setWordWrap(True)
        title.setTextInteractionFlags(Qt.TextSelectableByMouse)
        layout.addWidget(title)

        self.lbl_error_hint = QLabel()
        self.lbl_error_hint.setObjectName("noteText")
        self.lbl_error_hint.setWordWrap(True)
        self.lbl_error_hint.setVisible(False)
        layout.addWidget(self.lbl_error_hint)

        summary = QFrame()
        summary.setObjectName("extractSummary")
        summary_layout = QVBoxLayout(summary)
        summary_layout.setContentsMargins(8, 8, 8, 8)
        summary_layout.setSpacing(6)

        metric_row = QHBoxLayout()
        metric_row.setSpacing(6)
        self.summary_values: Dict[str, QLabel] = {}
        self.summary_metrics: Dict[str, QFrame] = {}
        for key, label, tone in (
            ("files", "Tổng số file", "info"),
            ("total", "Dòng bóc tách", "info"),
            ("ok", "Dòng OK", "success"),
            ("need_check", "Cần kiểm tra", "quiet"),
            ("warnings", "Cảnh báo", "quiet"),
        ):
            metric = QFrame()
            metric.setObjectName("extractMetric")
            metric.setProperty("tone", tone)
            metric_layout = QVBoxLayout(metric)
            metric_layout.setContentsMargins(10, 5, 10, 5)
            metric_layout.setSpacing(0)
            value = QLabel("0")
            value.setObjectName("extractMetricValue")
            value.setProperty("tone", tone)
            caption = QLabel(label)
            caption.setObjectName("extractMetricLabel")
            metric_layout.addWidget(value)
            metric_layout.addWidget(caption)
            metric_row.addWidget(metric, stretch=1)
            self.summary_values[key] = value
            self.summary_metrics[key] = metric
        summary_layout.addLayout(metric_row)

        type_row = QHBoxLayout()
        type_row.setSpacing(6)
        self.type_values: Dict[str, QLabel] = {}
        for key, label in (
            (DOC_EXPENSE, "Khoản chi"),
            (DOC_SCALE, "Phiếu cân"),
            (DOC_BILL, "Bill"),
            ("other", "Loại khác"),
        ):
            chip = QFrame()
            chip.setObjectName("extractTypeChip")
            chip_layout = QHBoxLayout(chip)
            chip_layout.setContentsMargins(8, 2, 8, 2)
            chip_layout.setSpacing(6)
            caption = QLabel(label)
            caption.setObjectName("extractTypeLabel")
            value = QLabel("0")
            value.setObjectName("extractTypeValue")
            chip_layout.addWidget(caption)
            chip_layout.addWidget(value)
            type_row.addWidget(chip)
            self.type_values[key] = value
        type_row.addStretch(1)
        summary_layout.addLayout(type_row)
        layout.addWidget(summary)

        filters = QHBoxLayout()
        self.status_filter = QComboBox()
        self.status_filter.addItems(["Mọi trạng thái", "Cần kiểm tra", "Sẵn sàng"])
        self.type_filter = QComboBox()
        self.type_filter.addItem("Mọi chứng từ", "ALL")
        self.type_filter.addItem("Khoản chi", DOC_EXPENSE)
        self.type_filter.addItem("Phiếu cân", DOC_SCALE)
        self.type_filter.addItem("Bill", DOC_BILL)
        self.type_filter.addItem("Loại khác", "OTHER")
        self.search = QLineEdit()
        self.search.setPlaceholderText("Tìm chứng từ, container, file nguồn...")
        self.status_filter.currentIndexChanged.connect(self._refresh_table)
        self.type_filter.currentIndexChanged.connect(self._refresh_table)
        self.search.textChanged.connect(self._refresh_table)
        filters.addWidget(self.status_filter)
        filters.addWidget(self.type_filter)
        filters.addWidget(self.search, stretch=1)
        layout.addLayout(filters)

        self.table = QTableWidget(0, len(self.COLUMNS))
        self.table.setHorizontalHeaderLabels(self.COLUMNS)
        self.table.setAlternatingRowColors(True)
        self.table.setSelectionBehavior(QTableWidget.SelectRows)
        self.table.setSelectionMode(QTableWidget.SingleSelection)
        self.table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.table.verticalHeader().setVisible(False)
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeToContents)
        self.table.horizontalHeader().setSectionResizeMode(2, QHeaderView.Fixed)
        self.table.setColumnWidth(2, 160)
        self.table.horizontalHeader().setStretchLastSection(True)
        self.table.horizontalHeader().setSortIndicatorShown(True)
        self.table.horizontalHeader().sectionClicked.connect(self._sort_by)
        self.table.cellClicked.connect(self._show_full_source_name)
        self.table.doubleClicked.connect(self.edit_row)
        layout.addWidget(self.table, stretch=1)

        buttons = QHBoxLayout()
        btn_edit = QPushButton("Sửa dòng")
        btn_delete = QPushButton("Xóa dòng")
        btn_save = QPushButton("Lưu")
        btn_save.setProperty("variant", "primary")
        btn_close = QPushButton("Đóng")
        btn_edit.clicked.connect(self.edit_row)
        btn_delete.clicked.connect(self.delete_row)
        btn_save.clicked.connect(self.save)
        btn_close.clicked.connect(self.reject)
        buttons.addWidget(btn_edit)
        buttons.addWidget(btn_delete)
        buttons.addStretch(1)
        buttons.addWidget(btn_save)
        buttons.addWidget(btn_close)
        layout.addLayout(buttons)

        self._refresh()
        if self.rows and self._first_error_row() in self._visible_indices:
            self.table.selectRow(self._visible_indices.index(self._first_error_row()))

    # ------------------------------------------------------------------ #
    # Hiển thị
    # ------------------------------------------------------------------ #
    def _refresh(self) -> None:
        error_indexes = {
            index for index, reason in enumerate(self.row_errors) if reason
        }
        counts = extract_summary(
            self.rows, self.payload.get("canh_bao"), error_indexes
        )
        for key, value in self.summary_values.items():
            value.setText(str(counts[key]))
        for key in ("need_check", "warnings"):
            tone = "warning" if counts[key] else "quiet"
            for widget in (self.summary_metrics[key], self.summary_values[key]):
                if widget.property("tone") != tone:
                    widget.setProperty("tone", tone)
                    repolish(widget)
        for key, value in self.type_values.items():
            value.setText(str(counts[key]))
        self.lbl_error_hint.setVisible(bool(error_indexes))
        if error_indexes:
            self.lbl_error_hint.setText(
                f"{len(error_indexes)} dòng chưa nhập được (tô đỏ bên dưới). Hãy sửa "
                "lại hoặc xóa dòng đó, bấm “Lưu” rồi nhập lại ở Bước 3."
            )

        self._refresh_table()

    def _sort_by(self, column: int) -> None:
        if column == self._sort_column:
            self._sort_ascending = not self._sort_ascending
        else:
            self._sort_column = column
            self._sort_ascending = True
        self._refresh_table()

    def _sort_value(self, index: int):
        row = self.rows[index]
        cells = self._row_cells(index, row, classify_doc_type(row.get("loai_chung_tu")))
        if self._sort_column == 0:
            try:
                return (0, int(cells[0]))
            except ValueError:
                return (1, cells[0].casefold())
        if self._sort_column == 4:
            parsed = parse_date(row.get("ngay_dong"))
            return (0, parsed) if parsed else (1, "")
        if self._sort_column in (7, 8):
            if self._sort_column == 8:
                value = parse_money(row.get("gia_vat_lieu"))
                return (0, value) if value is not None else (1, 0)
            is_tons = classify_doc_type(row.get("loai_chung_tu")) in (DOC_SCALE, DOC_BILL) and row.get("so_tan") not in (None, "")
            raw = row.get("so_tan") if is_tons else row.get("tong_tien")
            value = parse_number(raw) if is_tons else parse_money(raw)
            return (0, value) if value is not None else (1, 0)
        return (0, cells[self._sort_column].casefold())

    def _refresh_table(self, *_args) -> None:
        current = self._visible_indices[self.table.currentRow()] if (
            hasattr(self, "table") and 0 <= self.table.currentRow() < len(self._visible_indices)
        ) else None
        wanted_status = self.status_filter.currentText()
        wanted_type = self.type_filter.currentData()
        query = self.search.text().strip().casefold()
        visible = []
        for index, row in enumerate(self.rows):
            doc_type = classify_doc_type(row.get("loai_chung_tu"))
            warnings = _row_warning_count(row)
            needs_check = bool(self.row_errors[index]) or needs_review(row)
            if wanted_status == "Cần kiểm tra" and not needs_check:
                continue
            if wanted_status == "Sẵn sàng" and needs_check:
                continue
            if wanted_type != "ALL" and (
                (wanted_type == "OTHER" and doc_type in (DOC_EXPENSE, DOC_SCALE, DOC_BILL))
                or (wanted_type != "OTHER" and doc_type != wanted_type)
            ):
                continue
            if query and query not in " ".join(map(str, self._row_cells(index, row, doc_type))).casefold():
                continue
            visible.append(index)
        if self._sort_column in (4, 7, 8):
            filled = [index for index in visible if self._sort_value(index)[0] == 0]
            blank = [index for index in visible if self._sort_value(index)[0] != 0]
            filled.sort(key=self._sort_value, reverse=not self._sort_ascending)
            visible = filled + blank
        else:
            visible.sort(key=self._sort_value, reverse=not self._sort_ascending)
        self._visible_indices = visible
        self.table.blockSignals(True)
        self.table.setRowCount(len(visible))
        for table_row, index in enumerate(visible):
            row = self.rows[index]
            reason = self.row_errors[index]
            doc_type = classify_doc_type(row.get("loai_chung_tu"))
            warnings = _row_warning_count(row)
            needs_check = needs_review(row)
            for col, text in enumerate(self._row_cells(index, row, doc_type)):
                item = QTableWidgetItem(text)
                item.setToolTip(self._row_tooltip(row, reason))
                if col == 2:
                    item.setToolTip(str(row.get("file_nguon") or "") + "\nBấm để xem tên đầy đủ")
                if reason:
                    item.setBackground(ERROR_BG)
                    item.setForeground(ERROR_FG)
                elif needs_check or warnings:
                    item.setBackground(WARN_BG)
                if not reason and col == 9:
                    item.setForeground(QColor("#B45309" if needs_check else "#15803D"))
                self.table.setItem(table_row, col, item)
        self.table.blockSignals(False)
        if visible:
            self.table.selectRow(visible.index(current) if current in visible else 0)
        self.table.horizontalHeader().setSortIndicator(
            self._sort_column, Qt.AscendingOrder if self._sort_ascending else Qt.DescendingOrder
        )

    def _row_cells(
        self, index: int, row: Dict[str, Any], doc_type: str
    ) -> List[str]:
        stt = row.get("stt_hien_thi")
        return [
            str(stt if stt not in (None, "") else index + 1),
            {DOC_EXPENSE: "Khoản chi", DOC_SCALE: "Phiếu cân", DOC_BILL: "Bill"}.get(
                doc_type, str(row.get("loai_chung_tu") or "—")
            ),
            str(row.get("file_nguon") or ""),
            str(row.get("so_container") or ""),
            _display_date(row.get("ngay_dong")),
            str(row.get("loai_hang") or ""),
            _main_info(doc_type, row),
            _amount_info(doc_type, row),
            _display_number(row.get("gia_vat_lieu")),
            self._status_text(index, row),
        ]

    def _status_text(self, index: int, row: Dict[str, Any]) -> str:
        if self.row_errors[index]:
            return self.row_errors[index]
        status = str(row.get("trang_thai") or "").strip() or "OK"
        warnings = _row_warning_count(row)
        if warnings or has_low_confidence(row):
            return "Cần kiểm tra"
        return status

    @staticmethod
    def _row_tooltip(row: Dict[str, Any], reason: str) -> str:
        lines = []
        if reason:
            lines.append(f"Chưa nhập được: {reason}")
        lines.append(f"File nguồn: {row.get('file_nguon') or '—'}")
        for warning in warning_items(row.get("canh_bao")):
            lines.append("Cảnh báo: " + warning_text(warning))
        if has_low_confidence(row):
            lines.append("Độ tin cậy bóc tách thấp; cần đối chiếu chứng từ.")
        return "\n".join(lines)

    @Slot(int, int)
    def _show_full_source_name(self, table_row: int, column: int) -> None:
        if column == 2 and 0 <= table_row < len(self._visible_indices):
            filename = str(self.rows[self._visible_indices[table_row]].get("file_nguon") or "")
            notify.info(self, "Tên file nguồn đầy đủ", filename or "Không có tên file nguồn.")

    def _first_error_row(self) -> int:
        for index, reason in enumerate(self.row_errors):
            if reason:
                return index
        return 0

    def _selected_index(self) -> Optional[int]:
        table_row = self.table.currentRow()
        if table_row < 0 or table_row >= len(self._visible_indices):
            notify.info(
                self, "Chưa chọn dòng", "Hãy chọn một dòng trong bảng trước."
            )
            return None
        return self._visible_indices[table_row]

    # ------------------------------------------------------------------ #
    # Thao tác
    # ------------------------------------------------------------------ #
    @Slot()
    def edit_row(self) -> None:
        index = self._selected_index()
        if index is None:
            return
        dialog = _EditExtractRowDialog(self.rows[index], self.row_errors[index], self)
        if dialog.exec() != QDialog.Accepted:
            return
        self.rows[index] = dialog.updated_row()
        if self.row_errors[index] and not self.row_errors[index].startswith("Đã sửa; cần nhập lại"):
            self.row_errors[index] = "Đã sửa; cần nhập lại để kiểm tra. Lỗi lần trước: " + self.row_errors[index]
        self._dirty = True
        self._refresh()
        if index in self._visible_indices:
            self.table.selectRow(self._visible_indices.index(index))

    @Slot()
    def delete_row(self) -> None:
        index = self._selected_index()
        if index is None:
            return
        row = self.rows[index]
        confirmed = notify.confirm(
            self,
            "Xóa dòng",
            "Xóa dòng này khỏi dữ liệu bóc tách?\n\n"
            f"{row.get('loai_chung_tu') or 'Chứng từ'} • "
            f"{row.get('so_container') or '—'} • {row.get('file_nguon') or '—'}",
            confirm_text="Xóa dòng",
            cancel_text="Giữ lại",
            tone="warning",
            confirm_variant="danger",
        )
        if not confirmed:
            return
        self.rows.pop(index)
        self.row_errors.pop(index)
        self._dirty = True
        self._refresh()
        if self._visible_indices:
            next_row = next((row for row, source_index in enumerate(self._visible_indices) if source_index >= index), len(self._visible_indices) - 1)
            self.table.selectRow(next_row)

    @Slot()
    def save(self) -> None:
        """Lưu xong là đóng luôn.

        Không hiện thêm hộp thoại “đã lưu” bắt người dùng tắt: dòng “Lưu thành công
        lần cuối” ở Bước 2 tự cập nhật, thế là đủ báo thành công.
        """
        # Revenue has its own review screen. Preserve it verbatim in a mixed
        # GPT output when the user saves changes to the other document types.
        self.payload["du_lieu_boc_tach"] = self.rows + self._revenue_rows
        try:
            save_extract_payload(self.path, self.payload)
        except DailyImportError as exc:
            notify.error(self, "Không lưu được JSON", str(exc))
            return
        self.saved = True
        self._dirty = False
        self.accept()

    def reject(self) -> None:  # noqa: N802 - override Qt
        if self._dirty:
            confirmed = notify.confirm(
                self,
                "Chưa lưu thay đổi",
                "Bạn đã sửa dữ liệu nhưng chưa bấm “Lưu”. Đóng và bỏ các thay đổi?",
                confirm_text="Bỏ thay đổi",
                cancel_text="Tiếp tục chỉnh sửa",
                tone="warning",
                confirm_variant="danger",
            )
            if not confirmed:
                return
        super().reject()


class UnmatchedRowsDialog(ResizableDialog):
    """Popup khi có dòng không ghép được: sửa lại ở Bước 2, hoặc bỏ dòng lỗi."""

    BACK_TO_REVIEW = "REVIEW"
    DROP_ERRORS = "DROP"

    def __init__(self, rows: List[UnmatchedRow], parent=None):
        super().__init__(parent)
        # Đóng cửa sổ bằng dấu X = không nhập gì thêm, quay lại kiểm tra.
        self.choice = self.BACK_TO_REVIEW
        self.setWindowTitle("Có dòng chưa nhập được")
        self.resize(880, 420)

        layout = QVBoxLayout(self)
        hint = QLabel(
            f"{len(rows)} dòng không ghép được với dữ liệu quyết toán nên chưa được "
            "nhập. Các dòng còn lại vẫn nhập bình thường."
        )
        hint.setWordWrap(True)
        layout.addWidget(hint)

        table = QTableWidget(len(rows), 5)
        table.setHorizontalHeaderLabels(
            [
                "Loại chứng từ",
                "Container",
                "Ngày dùng để ghép",
                "Số HĐ/Bill/Seal",
                "Lý do",
            ]
        )
        table.setAlternatingRowColors(True)
        table.setSelectionMode(QTableWidget.NoSelection)
        table.setEditTriggers(QTableWidget.NoEditTriggers)
        table.verticalHeader().setVisible(False)
        for index, row in enumerate(rows):
            values = [
                row.doc_label,
                row.container,
                _display_date(row.match_date),
                row.reference,
                row.reason,
            ]
            for col, value in enumerate(values):
                table.setItem(index, col, QTableWidgetItem(str(value)))
        table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeToContents)
        table.horizontalHeader().setStretchLastSection(True)
        layout.addWidget(table, stretch=1)

        buttons = QHBoxLayout()
        btn_drop = QPushButton("Hủy các dòng lỗi và nhập các dòng còn lại")
        btn_review = QPushButton("Quay lại bước 2 để kiểm tra")
        btn_review.setProperty("variant", "primary")
        btn_review.setDefault(True)
        btn_drop.clicked.connect(self._choose_drop)
        btn_review.clicked.connect(self._choose_review)
        buttons.addStretch(1)
        buttons.addWidget(btn_drop)
        buttons.addWidget(btn_review)
        layout.addLayout(buttons)

    @Slot()
    def _choose_review(self) -> None:
        self.choice = self.BACK_TO_REVIEW
        self.accept()

    @Slot()
    def _choose_drop(self) -> None:
        self.choice = self.DROP_ERRORS
        self.accept()


class ConflictDialog(ResizableDialog):
    """Cho người dùng quyết định khi dữ liệu mới khác ô đã nhập."""

    def __init__(self, conflicts: List[FieldConflict], parent=None):
        super().__init__(parent)
        self.conflicts = conflicts
        self.combos: Dict[str, QComboBox] = {}
        self.setWindowTitle("Kiểm tra dữ liệu khác nhau")
        self.resize(900, 440)

        layout = QVBoxLayout(self)
        hint = QLabel(
            "Dữ liệu mới khác thông tin đang có. Mặc định phần mềm giữ dữ liệu "
            "hiện tại; bạn có thể chọn dùng dữ liệu mới cho từng trường."
        )
        hint.setWordWrap(True)
        layout.addWidget(hint)

        table = QTableWidget(len(conflicts), 5)
        table.setHorizontalHeaderLabels(
            ["Dòng Excel", "Trường", "Hiện tại", "Dữ liệu mới", "Lựa chọn"]
        )
        table.setAlternatingRowColors(True)
        table.setEditTriggers(QTableWidget.NoEditTriggers)
        for row_index, conflict in enumerate(conflicts):
            values = [
                str(conflict.target_row),
                conflict.field_name,
                str(conflict.current_value or ""),
                str(conflict.new_value or ""),
            ]
            for col, value in enumerate(values):
                table.setItem(row_index, col, QTableWidgetItem(value))
            combo = QComboBox()
            combo.addItem("Giữ dữ liệu hiện tại", False)
            combo.addItem("Dùng dữ liệu mới", True)
            table.setCellWidget(row_index, 4, combo)
            self.combos[conflict.conflict_id] = combo
        table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeToContents)
        table.horizontalHeader().setStretchLastSection(True)
        layout.addWidget(table)

        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.button(QDialogButtonBox.Ok).setText("Xác nhận lựa chọn")
        buttons.button(QDialogButtonBox.Cancel).setText("Hủy lần nhập")
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)

    def decisions(self) -> Dict[str, bool]:
        return {
            key: bool(combo.currentData()) for key, combo in self.combos.items()
        }


class _PasteTable(QTableWidget):
    """Dán nhiều ô từ Excel vào bảng sửa nhanh."""

    def keyPressEvent(self, event) -> None:  # noqa: N802
        if event.matches(QKeySequence.Paste):
            start_row = max(0, self.currentRow())
            start_col = max(1, self.currentColumn())
            for row_offset, line in enumerate(QApplication.clipboard().text().splitlines()):
                if start_row + row_offset >= self.rowCount():
                    break
                for col_offset, value in enumerate(line.split("\t")):
                    col = start_col + col_offset
                    if col >= self.columnCount():
                        break
                    item = self.item(start_row + row_offset, col)
                    if item is not None:
                        item.setText(value.strip())
            return
        super().keyPressEvent(event)


class BulkQuickEditDialog(ResizableDialog):
    def __init__(self, issues: List[ImportIssue], payload: Dict[str, Any], parent=None):
        super().__init__(parent)
        self.issues = issues
        self.payload = payload
        self.changes: Dict[int, tuple[str, str]] = {}
        self.setWindowTitle("Sửa nhanh nhiều dòng")
        self.resize(760, 570)
        layout = QVBoxLayout(self)
        layout.addWidget(QLabel(f"Sửa thông tin để ghép lại {len(issues)} dòng"))
        note = QLabel("Sửa ngày đóng, container; có thể dán nhiều ô từ Excel. Phần mềm sẽ kiểm tra lại từng dòng sau khi áp dụng.")
        note.setWordWrap(True)
        layout.addWidget(note)
        self.table = _PasteTable(len(issues), 3)
        self.table.setHorizontalHeaderLabels(["Chứng từ", "Ngày đóng", "Container"])
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        rows = payload.get("du_lieu_boc_tach") or []
        for table_row, issue in enumerate(issues):
            index = issue.json_indexes[0]
            source = rows[index]
            label = QTableWidgetItem(issue.reference or issue.doc_label)
            label.setFlags(label.flags() & ~Qt.ItemIsEditable)
            self.table.setItem(table_row, 0, label)
            self.table.setItem(table_row, 1, QTableWidgetItem(_display_date(source.get("ngay_dong"))))
            self.table.setItem(table_row, 2, QTableWidgetItem(str(source.get("so_container") or "")))
        layout.addWidget(self.table, stretch=1)
        actions = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        actions.button(QDialogButtonBox.Ok).setText("Kiểm tra và áp dụng")
        actions.button(QDialogButtonBox.Cancel).setText("Hủy")
        actions.accepted.connect(self._accept)
        actions.rejected.connect(self.reject)
        layout.addWidget(actions)

    def _accept(self) -> None:
        for row, issue in enumerate(self.issues):
            date_text = self.table.item(row, 1).text().strip()
            container = self.table.item(row, 2).text().strip().upper()
            if date_text and parse_date(date_text) is None:
                notify.warning(self, "Ngày chưa hợp lệ", f"Dòng {row + 1}: hãy kiểm tra lại ngày đóng.")
                return
            self.changes[issue.json_indexes[0]] = (date_text, container)
        self.accept()


class BulkCandidateDialog(ResizableDialog):
    def __init__(self, issues: List[ImportIssue], candidates: List[ImportCandidate], parent=None):
        super().__init__(parent)
        self.choice = ""
        self.setWindowTitle("Chọn cùng một kết quả ghép")
        self.resize(800, 550)
        layout = QVBoxLayout(self)
        layout.addWidget(QLabel(f"Áp dụng cho {len(issues)} dòng đã chọn"))
        self.options = QComboBox()
        for candidate in candidates:
            self.options.addItem(candidate.label.replace("SQT", "số quyết toán"), candidate.candidate_id)
        layout.addWidget(self.options)
        table = QTableWidget(len(issues), 3)
        table.setHorizontalHeaderLabels(["Chứng từ", "Container", "Vấn đề"])
        table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        for row, issue in enumerate(issues):
            for col, value in enumerate((issue.reference or issue.doc_label, issue.container, issue.title.replace("SQT", "số quyết toán"))):
                table.setItem(row, col, QTableWidgetItem(value))
        layout.addWidget(table, stretch=1)
        layout.addWidget(QLabel("Chỉ những dòng trong bảng sẽ được thay đổi."))
        actions = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        actions.button(QDialogButtonBox.Ok).setText("Áp dụng cho các dòng này")
        actions.accepted.connect(self._accept)
        actions.rejected.connect(self.reject)
        layout.addWidget(actions)

    def _accept(self) -> None:
        self.choice = str(self.options.currentData() or "")
        if self.choice:
            self.accept()


class BulkFieldDialog(ResizableDialog):
    def __init__(self, issues: List[ImportIssue], parent=None):
        super().__init__(parent)
        self.modes: Dict[str, str] = {}
        self.setWindowTitle("Chọn dữ liệu cho nhiều dòng")
        self.resize(900, 590)
        layout = QVBoxLayout(self)
        layout.addWidget(QLabel(f"Chọn cách xử lý các trường khác nhau trong {len(issues)} dòng"))
        note = QLabel("Mỗi trường cần một lựa chọn. Hãy xem giá trị của từng dòng trước khi áp dụng chung.")
        note.setWordWrap(True)
        layout.addWidget(note)
        names = sorted({conflict.field_name for issue in issues for conflict in issue.conflicts})
        self.combos: Dict[str, QComboBox] = {}
        choices = QTableWidget(len(names), 3)
        choices.setHorizontalHeaderLabels(["Trường", "Số dòng", "Cách xử lý"])
        choices.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        for row, name in enumerate(names):
            choices.setItem(row, 0, QTableWidgetItem(name))
            count = sum(any(c.field_name == name for c in issue.conflicts) for issue in issues)
            choices.setItem(row, 1, QTableWidgetItem(str(count)))
            combo = QComboBox()
            combo.addItem("Chọn cách xử lý", "")
            combo.addItem("Giữ dữ liệu đang có", "CURRENT")
            combo.addItem("Dùng dữ liệu mới", "NEW")
            choices.setCellWidget(row, 2, combo)
            self.combos[name] = combo
        layout.addWidget(choices)
        preview = QTableWidget(sum(len(i.conflicts) for i in issues), 4)
        preview.setHorizontalHeaderLabels(["Chứng từ", "Trường", "Hiện tại", "Dữ liệu mới"])
        preview.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        preview.setEditTriggers(QAbstractItemView.NoEditTriggers)
        row = 0
        for issue in issues:
            for conflict in issue.conflicts:
                for col, value in enumerate((issue.reference or issue.doc_label, conflict.field_name, _display_value(conflict.current_value), _display_value(conflict.new_value))):
                    preview.setItem(row, col, QTableWidgetItem(value))
                row += 1
        layout.addWidget(preview, stretch=1)
        actions = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        actions.button(QDialogButtonBox.Ok).setText("Áp dụng cho các dòng này")
        actions.accepted.connect(self._accept)
        actions.rejected.connect(self.reject)
        layout.addWidget(actions)

    def _accept(self) -> None:
        self.modes = {name: str(combo.currentData() or "") for name, combo in self.combos.items()}
        if not all(self.modes.values()):
            notify.warning(self, "Chưa chọn đủ", "Hãy chọn cách xử lý cho từng trường.")
            return
        self.accept()


class ImportResolutionDialog(ResizableDialog):
    """Cửa sổ duy nhất để xử lý mọi vấn đề trước khi ghi file Hàng ngày."""

    def __init__(
        self,
        service: DailyImportService,
        analysis: ImportAnalysis,
        parent=None,
    ):
        super().__init__(parent)
        self.service = service
        self.analysis = analysis
        self.payload = load_extract_payload(analysis.output_path)
        self.resolutions: Dict[str, IssueResolution] = dict(analysis.resolutions)
        self.final_analysis: Optional[ImportAnalysis] = None
        self.saved_only = False
        self._dirty = False
        self._reviewing = False
        self._visible_issue_ids: List[str] = []
        self._selected_issue_ids: set[str] = set()
        self._bulk_undo: Optional[tuple[Dict[str, IssueResolution], Dict[str, Any], set[str]]] = None
        self._sort_column = 1
        self._sort_ascending = True
        self._candidate_group: Optional[QButtonGroup] = None

        self.setWindowTitle("Xử lý dữ liệu trước khi nhập")
        self.resize(1380, 850)
        self.setMinimumSize(1050, 680)

        root = QVBoxLayout(self)
        title = QLabel("Xử lý dữ liệu trước khi nhập")
        title.setStyleSheet("font-size: 18pt; font-weight: 700; color: #111827;")
        subtitle = QLabel("Chọn một hoặc nhiều dòng để xử lý. File theo dõi chưa được thay đổi.")
        subtitle.setObjectName("hintText")
        root.addWidget(title)
        root.addWidget(subtitle)

        summary_row = QHBoxLayout()
        self.lbl_total = self._summary_card("#DC2626")
        self.lbl_done = self._summary_card("#16A34A")
        self.lbl_skipped = self._summary_card("#6B7280")
        self.lbl_pending = self._summary_card("#2563EB")
        for widget in (
            self.lbl_total,
            self.lbl_done,
            self.lbl_skipped,
            self.lbl_pending,
        ):
            summary_row.addWidget(widget)
        root.addLayout(summary_row)

        body = QHBoxLayout()
        left = QFrame()
        left.setObjectName("card")
        left_layout = QVBoxLayout(left)
        self.search = QLineEdit()
        self.search.setPlaceholderText("Tìm theo container, số quyết toán, chứng từ...")
        self.search.textChanged.connect(lambda _text: self._refresh_issue_table())
        left_layout.addWidget(self.search)
        self.filter_combo = QComboBox()
        self.filter_combo.addItem("Tất cả", "ALL")
        self.filter_combo.addItem("Chưa xử lý", "PENDING")
        self.filter_combo.addItem("Đã xử lý", "DONE")
        self.filter_combo.addItem("Để xử lý sau", "SKIP")
        self.filter_combo.currentIndexChanged.connect(
            lambda _index: self._refresh_issue_table()
        )
        left_layout.addWidget(self.filter_combo)

        self.kind_filter = QComboBox()
        self.kind_filter.addItem("Mọi vấn đề", "ALL")
        for kind, label in (
            (ISSUE_NO_TARGET, "Không tìm thấy số quyết toán"),
            (ISSUE_MISSING_CLOSE_DATE, "Thiếu ngày đóng"),
            (ISSUE_MULTIPLE_TARGETS, "Nhiều số quyết toán phù hợp"),
            (ISSUE_MULTIPLE_BILLS, "Nhiều Bill phù hợp"),
            (ISSUE_FIELD_CONFLICT, "Dữ liệu khác nhau"),
        ):
            self.kind_filter.addItem(label, kind)
        self.kind_filter.currentIndexChanged.connect(self._refresh_issue_table)
        left_layout.addWidget(self.kind_filter)

        selection_row = QHBoxLayout()
        self.btn_select_visible = QPushButton("Chọn tất cả đang hiện")
        self.btn_clear_selection = QPushButton("Bỏ chọn")
        self.btn_select_visible.clicked.connect(self._select_visible_issues)
        self.btn_clear_selection.clicked.connect(self._clear_issue_selection)
        selection_row.addWidget(self.btn_select_visible)
        selection_row.addWidget(self.btn_clear_selection)
        selection_row.addStretch(1)
        left_layout.addLayout(selection_row)

        self.issue_table = QTableWidget(0, 7)
        self.issue_table.setHorizontalHeaderLabels(
            ["Chọn", "STT", "Chứng từ", "Số chứng từ", "Container", "Vấn đề", "Trạng thái"]
        )
        self.issue_table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.issue_table.setSelectionMode(QAbstractItemView.SingleSelection)
        self.issue_table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.issue_table.setAlternatingRowColors(True)
        self.issue_table.verticalHeader().setVisible(False)
        self.issue_table.horizontalHeader().setSectionResizeMode(QHeaderView.Interactive)
        self.issue_table.horizontalHeader().setStretchLastSection(True)
        for column, width in enumerate((47, 48, 90, 105, 125, 180)):
            self.issue_table.setColumnWidth(column, width)
        self.issue_table.horizontalHeader().setSortIndicatorShown(True)
        self.issue_table.horizontalHeader().sectionClicked.connect(self._sort_issues_by)
        self.issue_table.itemChanged.connect(self._check_issue_changed)
        self.issue_table.itemSelectionChanged.connect(self._show_selected_issue)
        left_layout.addWidget(self.issue_table, stretch=1)
        self.lbl_visible = QLabel()
        self.lbl_visible.setObjectName("hintText")
        left_layout.addWidget(self.lbl_visible)

        self.detail_frame = QFrame()
        self.detail_frame.setObjectName("card")
        self.detail_layout = QVBoxLayout(self.detail_frame)
        body.addWidget(left, 59)
        body.addWidget(self.detail_frame, 41)
        root.addLayout(body, stretch=1)

        bulk_row = QHBoxLayout()
        self.lbl_selected = QLabel("Chưa chọn dòng nào")
        self.lbl_selected.setObjectName("hintText")
        self.btn_bulk = QPushButton("Xử lý hàng loạt ▾")
        self.btn_bulk.clicked.connect(self._show_bulk_menu)
        self.btn_defer_selected = QPushButton("Để xử lý sau")
        self.btn_defer_selected.clicked.connect(self._defer_selected_issues)
        self.btn_undo_bulk = QPushButton("Hoàn tác")
        self.btn_undo_bulk.clicked.connect(self._undo_last_bulk)
        bulk_row.addWidget(self.lbl_selected)
        bulk_row.addStretch(1)
        bulk_row.addWidget(self.btn_bulk)
        bulk_row.addWidget(self.btn_defer_selected)
        bulk_row.addWidget(self.btn_undo_bulk)
        root.addLayout(bulk_row)

        footer = QHBoxLayout()
        self.lbl_progress = QLabel()
        self.lbl_progress.setObjectName("hintText")
        footer.addWidget(self.lbl_progress)
        footer.addStretch(1)
        self.btn_cancel = QPushButton("Hủy")
        self.btn_save = QPushButton("Lưu và đóng")
        self.btn_continue = QPushButton("Tiếp tục nhập")
        self.btn_continue.setProperty("variant", "primary")
        self.btn_cancel.clicked.connect(self.reject)
        self.btn_save.clicked.connect(self._save_and_close)
        self.btn_continue.clicked.connect(self._continue)
        footer.addWidget(self.btn_cancel)
        footer.addWidget(self.btn_save)
        footer.addWidget(self.btn_continue)
        root.addLayout(footer)

        self._refresh_all()

    @staticmethod
    def _summary_card(color: str) -> QLabel:
        label = QLabel()
        label.setAlignment(Qt.AlignCenter)
        label.setMinimumHeight(54)
        label.setStyleSheet(
            "background: #FFFFFF; border: 1px solid #E5E7EB; border-radius: 9px; "
            f"font-size: 12pt; font-weight: 700; color: {color}; padding: 8px;"
        )
        return label

    @staticmethod
    def _clear_layout(layout: QVBoxLayout) -> None:
        while layout.count():
            item = layout.takeAt(0)
            widget = item.widget()
            child_layout = item.layout()
            if widget is not None:
                # deleteLater() alone leaves the old widget visible until Qt gets
                # back to the event loop.  The detail panel can be rebuilt more
                # than once in one signal, so detach it immediately first.
                widget.hide()
                widget.setParent(None)
                widget.deleteLater()
            elif child_layout is not None:
                ImportResolutionDialog._clear_layout(child_layout)

    def _issue_status(self, issue: ImportIssue) -> str:
        resolution = self.resolutions.get(issue.issue_id)
        if resolution is None:
            return "PENDING"
        return "SKIP" if resolution.action == RESOLUTION_SKIP else "DONE"

    def _refresh_all(self, preferred_issue_id: str = "") -> None:
        self._reviewing = False
        self.btn_continue.setText("Tiếp tục nhập")
        issues = self.analysis.issues
        pending = sum(self._issue_status(issue) == "PENDING" for issue in issues)
        done = sum(self._issue_status(issue) == "DONE" for issue in issues)
        skipped = sum(self._issue_status(issue) == "SKIP" for issue in issues)
        self.lbl_total.setText(f"{len(issues)}  vấn đề trong lô")
        self.lbl_done.setText(f"{done}  đã xử lý")
        self.lbl_skipped.setText(f"{skipped}  để sau")
        self.lbl_pending.setText(f"{pending}  chưa xử lý")
        self.lbl_progress.setText(
            f"Đã quyết định {done + skipped}/{len(issues)} dòng • "
            f"Còn {pending} dòng cần xử lý"
        )
        self.btn_continue.setEnabled(pending == 0)
        self.btn_continue.setToolTip(
            "" if pending == 0 else "Cần xử lý hoặc để sau tất cả các dòng trước."
        )
        self._selected_issue_ids.intersection_update(issue.issue_id for issue in issues)
        self._refresh_issue_table(preferred_issue_id)

    def _sort_issues_by(self, column: int) -> None:
        if column == 0:
            return
        if column == self._sort_column:
            self._sort_ascending = not self._sort_ascending
        else:
            self._sort_column = column
            self._sort_ascending = True
        self._refresh_issue_table()

    def _issue_sort_value(self, issue: ImportIssue, original_order: int):
        status = self._issue_status(issue)
        values = {
            1: original_order,
            2: issue.doc_label.casefold(),
            3: issue.reference.casefold(),
            4: issue.container.casefold(),
            5: issue.title.casefold(),
            6: {"PENDING": 0, "DONE": 1, "SKIP": 2}[status],
        }
        return values.get(self._sort_column, original_order)

    @Slot()
    def _refresh_issue_table(self, preferred_issue_id: str = "") -> None:
        if not isinstance(preferred_issue_id, str):
            preferred_issue_id = ""
        current = preferred_issue_id or self._selected_issue_id()
        query = self.search.text().strip().lower()
        wanted = self.filter_combo.currentData()
        wanted_kind = self.kind_filter.currentData()
        visible: List[ImportIssue] = []
        for issue in self.analysis.issues:
            status = self._issue_status(issue)
            if wanted != "ALL" and status != wanted:
                continue
            if wanted_kind != "ALL" and issue.kind != wanted_kind:
                continue
            haystack = " ".join(
                [issue.doc_label, issue.container, issue.title, issue.reference]
            ).lower()
            if query and query not in haystack:
                continue
            visible.append(issue)

        order = {issue.issue_id: index for index, issue in enumerate(self.analysis.issues)}
        visible.sort(
            key=lambda issue: self._issue_sort_value(issue, order[issue.issue_id]),
            reverse=not self._sort_ascending,
        )

        self._visible_issue_ids = [issue.issue_id for issue in visible]
        self.issue_table.blockSignals(True)
        self.issue_table.setRowCount(len(visible))
        selected_row = 0
        for row_index, issue in enumerate(visible):
            status = self._issue_status(issue)
            status_text = {
                "PENDING": "Chưa xử lý",
                "DONE": "Đã xử lý",
                "SKIP": "Để sau",
            }[status]
            values = [
                str(order[issue.issue_id] + 1),
                issue.doc_label,
                issue.reference or "—",
                issue.container or "—",
                issue.title.replace("SQT", "số quyết toán"),
                status_text,
            ]
            check = QTableWidgetItem()
            check.setFlags(Qt.ItemIsEnabled | Qt.ItemIsUserCheckable)
            check.setCheckState(Qt.Checked if issue.issue_id in self._selected_issue_ids else Qt.Unchecked)
            self.issue_table.setItem(row_index, 0, check)
            for column, value in enumerate(values, start=1):
                item = QTableWidgetItem(value)
                if status == "PENDING":
                    item.setBackground(ERROR_BG if row_index == 0 else WARN_BG)
                elif status == "SKIP":
                    item.setForeground(QColor("#6B7280"))
                else:
                    item.setBackground(QColor("#DCFCE7"))
                self.issue_table.setItem(row_index, column, item)
            if issue.issue_id == current:
                selected_row = row_index
        if visible:
            self.issue_table.selectRow(selected_row)
        else:
            self.issue_table.clearSelection()
        self.issue_table.blockSignals(False)
        self.issue_table.horizontalHeader().setSortIndicator(
            self._sort_column, Qt.AscendingOrder if self._sort_ascending else Qt.DescendingOrder
        )
        self.lbl_visible.setText(f"Đang hiện {len(visible)}/{len(self.analysis.issues)} dòng")
        self._refresh_bulk_state()
        if visible:
            self._show_issue(visible[selected_row])
        else:
            self._show_empty_detail()

    def _selected_issue_id(self) -> str:
        row = self.issue_table.currentRow()
        if 0 <= row < len(self._visible_issue_ids):
            return self._visible_issue_ids[row]
        return ""

    def _check_issue_changed(self, item: QTableWidgetItem) -> None:
        if item.column() != 0 or not (0 <= item.row() < len(self._visible_issue_ids)):
            return
        issue_id = self._visible_issue_ids[item.row()]
        if item.checkState() == Qt.Checked:
            self._selected_issue_ids.add(issue_id)
        else:
            self._selected_issue_ids.discard(issue_id)
        self._refresh_bulk_state()

    def _select_visible_issues(self) -> None:
        self._selected_issue_ids.update(self._visible_issue_ids)
        self._refresh_issue_table()

    def _clear_issue_selection(self) -> None:
        self._selected_issue_ids.clear()
        self._refresh_issue_table()

    def _selected_pending_issues(self) -> List[ImportIssue]:
        return [issue for issue in self.analysis.issues if issue.issue_id in self._selected_issue_ids and self._issue_status(issue) == "PENDING"]

    def _refresh_bulk_state(self) -> None:
        selected = len(self._selected_issue_ids)
        hidden = len(self._selected_issue_ids - set(self._visible_issue_ids))
        self.lbl_selected.setText(
            f"Đã chọn {selected} dòng" + (f" · {hidden} dòng đang ẩn bởi bộ lọc" if hidden else "")
            if selected else "Chưa chọn dòng nào"
        )
        enabled = bool(self._selected_pending_issues())
        self.btn_bulk.setEnabled(enabled)
        self.btn_defer_selected.setEnabled(enabled)
        self.btn_undo_bulk.setEnabled(self._bulk_undo is not None)

    def _snapshot_bulk(self) -> None:
        self._bulk_undo = (
            copy.deepcopy(self.resolutions), copy.deepcopy(self.payload),
            set(self._selected_issue_ids),
        )

    def _undo_last_bulk(self) -> None:
        if self._bulk_undo is None:
            return
        self.resolutions, self.payload, self._selected_issue_ids = self._bulk_undo
        self._bulk_undo = None
        self._dirty = True
        self._queue_reanalysis()

    def _build_bulk_menu(self) -> QMenu:
        menu = QMenu(self)
        pending = self._selected_pending_issues()
        edit_ids = tuple(i.issue_id for i in pending if i.kind in (ISSUE_NO_TARGET, ISSUE_MISSING_CLOSE_DATE))
        if edit_ids:
            action = menu.addAction(f"Sửa nhanh thông tin để ghép lại · {len(edit_ids)} dòng")
            action.triggered.connect(lambda _checked=False, ids=edit_ids: self._bulk_quick_edit(ids))
        for kind, label in (
            (ISSUE_MULTIPLE_TARGETS, "Chọn số quyết toán phù hợp"),
            (ISSUE_MULTIPLE_BILLS, "Chọn Bill phù hợp"),
            (ISSUE_FIELD_CONFLICT, "Chọn cách xử lý dữ liệu khác nhau"),
        ):
            ids = tuple(i.issue_id for i in pending if i.kind == kind)
            if ids:
                action = menu.addAction(f"{label} · {len(ids)} dòng")
                handler = self._bulk_fields if kind == ISSUE_FIELD_CONFLICT else self._bulk_candidates
                action.triggered.connect(lambda _checked=False, ids=ids, handler=handler: handler(ids))
        if pending:
            menu.addSeparator()
            action = menu.addAction(f"Để cả {len(pending)} dòng chưa xử lý đã chọn lại sau")
            action.triggered.connect(self._defer_selected_issues)
        return menu

    def _show_bulk_menu(self) -> None:
        menu = self._build_bulk_menu()
        menu.exec(self.btn_bulk.mapToGlobal(self.btn_bulk.rect().bottomLeft()))

    def _issues_for_ids(self, ids: tuple[str, ...]) -> List[ImportIssue]:
        return [issue for issue in self._selected_pending_issues() if issue.issue_id in ids]

    def _bulk_quick_edit(self, ids: tuple[str, ...]) -> None:
        issues = [issue for issue in self._issues_for_ids(ids) if issue.json_indexes]
        if not issues:
            return
        dialog = BulkQuickEditDialog(issues, self.payload, self)
        if dialog.exec() != QDialog.Accepted:
            return
        self._snapshot_bulk()
        rows = self.payload.get("du_lieu_boc_tach") or []
        for index, (date_text, container) in dialog.changes.items():
            rows[index]["ngay_dong"] = date_text
            rows[index]["so_container"] = container
        self._dirty = True
        self._queue_reanalysis()

    def _bulk_candidates(self, ids: tuple[str, ...]) -> None:
        issues = self._issues_for_ids(ids)
        if not issues:
            return
        common_ids = set(candidate.candidate_id for candidate in issues[0].candidates)
        for issue in issues[1:]:
            common_ids.intersection_update(candidate.candidate_id for candidate in issue.candidates)
        candidates = [candidate for candidate in issues[0].candidates if candidate.candidate_id in common_ids]
        if not candidates:
            notify.info(self, "Không có lựa chọn chung", "Các dòng này không có cùng một kết quả ghép hợp lệ. Hãy chọn từng dòng hoặc chọn một nhóm nhỏ hơn.")
            return
        dialog = BulkCandidateDialog(issues, candidates, self)
        if dialog.exec() != QDialog.Accepted:
            return
        self._snapshot_bulk()
        for issue in issues:
            self.resolutions[issue.issue_id] = IssueResolution(issue.issue_id, RESOLUTION_SELECT, candidate_id=dialog.choice)
        self._selected_issue_ids.difference_update(ids)
        self._dirty = True
        self._queue_reanalysis()

    def _bulk_fields(self, ids: tuple[str, ...]) -> None:
        issues = self._issues_for_ids(ids)
        if not issues:
            return
        dialog = BulkFieldDialog(issues, self)
        if dialog.exec() != QDialog.Accepted:
            return
        self._snapshot_bulk()
        for issue in issues:
            choices = {conflict.conflict_id: {"mode": dialog.modes[conflict.field_name], "value": ""} for conflict in issue.conflicts}
            self.resolutions[issue.issue_id] = IssueResolution(issue.issue_id, RESOLUTION_FIELDS, field_choices=choices)
        self._selected_issue_ids.difference_update(ids)
        self._dirty = True
        self._queue_reanalysis()

    def _defer_selected_issues(self) -> None:
        issues = self._selected_pending_issues()
        if not issues:
            return
        confirmed = notify.confirm(
            self, "Để xử lý sau",
            f"Để {len(issues)} dòng lại xử lý sau? Các dòng này sẽ không được nhập trong lần này và sẽ xuất hiện ở danh sách chờ.",
            confirm_text="Để xử lý sau", cancel_text="Quay lại",
        )
        if not confirmed:
            return
        self._snapshot_bulk()
        for issue in issues:
            self.resolutions[issue.issue_id] = IssueResolution(issue.issue_id, RESOLUTION_SKIP)
        self._selected_issue_ids.difference_update(issue.issue_id for issue in issues)
        self._dirty = True
        self._queue_reanalysis()

    @Slot()
    def _show_selected_issue(self) -> None:
        issue_id = self._selected_issue_id()
        issue = next(
            (item for item in self.analysis.issues if item.issue_id == issue_id),
            None,
        )
        if issue:
            self._show_issue(issue)

    def _show_empty_detail(self) -> None:
        self._clear_layout(self.detail_layout)
        label = QLabel("Không có dòng phù hợp với bộ lọc.")
        label.setAlignment(Qt.AlignCenter)
        label.setObjectName("hintText")
        self.detail_layout.addWidget(label, stretch=1)

    def _source_row(
        self, issue: ImportIssue, json_index: Optional[int] = None
    ) -> Dict[str, Any]:
        rows = self.payload.get("du_lieu_boc_tach") or []
        index = json_index
        if index is None and issue.json_indexes:
            index = issue.json_indexes[0]
        if index is not None and 0 <= index < len(rows):
            row = rows[index]
            return row if isinstance(row, dict) else {}
        return {}

    def _show_issue(self, issue: ImportIssue) -> None:
        self._clear_layout(self.detail_layout)
        header = QLabel(issue.title.replace("SQT", "số quyết toán"))
        header.setStyleSheet("font-size: 14pt; font-weight: 700; color: #111827;")
        reason = QLabel(issue.reason)
        reason.setObjectName("noteText")
        reason.setWordWrap(True)
        self.detail_layout.addWidget(header)
        self.detail_layout.addWidget(reason)

        source = self._source_row(issue)
        source_card = QFrame()
        source_card.setObjectName("guideBox")
        source_layout = QVBoxLayout(source_card)
        source_layout.addWidget(QLabel("Dữ liệu bóc tách"))
        source_text = QLabel(
            "\n".join(
                [
                    f"Container: {source.get('so_container') or '—'}",
                    f"Ngày đóng: {_display_date(source.get('ngay_dong')) or '—'}",
                    f"Loại hàng: {source.get('loai_hang') or '—'}",
                    f"File nguồn: {source.get('file_nguon') or '—'}",
                ]
            )
        )
        source_text.setTextInteractionFlags(Qt.TextSelectableByMouse)
        source_layout.addWidget(source_text)
        action_row = QHBoxLayout()
        action_row.addStretch(1)
        if issue.issue_id in self.resolutions:
            btn_reset = QPushButton("Đặt lại lựa chọn")
            btn_reset.clicked.connect(lambda: self._reset_issue(issue))
            action_row.addWidget(btn_reset)
        btn_edit = QPushButton("Chỉnh sửa dòng")
        btn_skip = QPushButton("Để xử lý sau")
        btn_edit.clicked.connect(lambda: self._show_editor(issue))
        btn_skip.clicked.connect(lambda: self._skip_issue(issue))
        action_row.addWidget(btn_edit)
        action_row.addWidget(btn_skip)
        source_layout.addLayout(action_row)
        self.detail_layout.addWidget(source_card)

        if issue.kind == ISSUE_FIELD_CONFLICT:
            self._show_conflict_choices(issue)
        elif issue.candidates:
            self._show_candidates(issue)
        else:
            note = QLabel(
                "Dòng này cần được chỉnh sửa để có thể ghép dữ liệu, hoặc chọn "
                "“Để xử lý sau” nếu chưa đủ thông tin."
            )
            note.setWordWrap(True)
            note.setObjectName("hintText")
            self.detail_layout.addWidget(note)
        self.detail_layout.addStretch(1)

    def _show_candidates(self, issue: ImportIssue) -> None:
        title = QLabel(
            "Chọn số quyết toán phù hợp"
            if any(item.kind == "INFO_ROW" for item in issue.candidates)
            else "Chọn Bill phù hợp"
        )
        title.setStyleSheet("font-size: 11.5pt; font-weight: 700;")
        self.detail_layout.addWidget(title)
        group = QButtonGroup(self)
        self._candidate_group = group
        selected_id = (self.resolutions.get(issue.issue_id) or IssueResolution("", "")).candidate_id
        for candidate in issue.candidates:
            card = QFrame()
            card.setObjectName("guideBox")
            layout = QVBoxLayout(card)
            radio = QRadioButton(candidate.label.replace("SQT", "số quyết toán"))
            radio.setChecked(candidate.candidate_id == selected_id)
            details = " • ".join(
                f"{key.replace('SQT PM', 'Số quyết toán')}: {_display_value(value)}"
                for key, value in candidate.values.items()
                if value not in (None, "")
            )
            detail_label = QLabel(details or "Không có thông tin bổ sung")
            detail_label.setWordWrap(True)
            detail_label.setObjectName("hintText")
            choice_row = QHBoxLayout()
            choice_row.addWidget(radio)
            choice_row.addStretch(1)
            if candidate.kind == "BILL" and candidate.json_index >= 0:
                btn_edit_bill = QPushButton("Sửa Bill này")
                btn_edit_bill.clicked.connect(
                    lambda _checked=False, index=candidate.json_index: self._show_editor(
                        issue, index
                    )
                )
                choice_row.addWidget(btn_edit_bill)
            layout.addLayout(choice_row)
            layout.addWidget(detail_label)
            group.addButton(radio)
            radio.toggled.connect(
                lambda checked, candidate_id=candidate.candidate_id: (
                    self._select_candidate(issue, candidate_id) if checked else None
                )
            )
            self.detail_layout.addWidget(card)

    def _show_conflict_choices(self, issue: ImportIssue) -> None:
        title = QLabel("Chọn dữ liệu sử dụng cho từng trường")
        title.setStyleSheet("font-size: 11.5pt; font-weight: 700;")
        self.detail_layout.addWidget(title)
        table = QTableWidget(len(issue.conflicts), 5)
        table.setHorizontalHeaderLabels(
            ["Trường", "Hiện tại", "Dữ liệu mới", "Lựa chọn", "Giá trị nhập"]
        )
        table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        combos: Dict[str, QComboBox] = {}
        editors: Dict[str, QLineEdit] = {}
        existing = self.resolutions.get(issue.issue_id)
        for row_index, conflict in enumerate(issue.conflicts):
            for column, value in enumerate(
                [conflict.field_name, conflict.current_value, conflict.new_value]
            ):
                table.setItem(row_index, column, QTableWidgetItem(_display_value(value)))
            combo = QComboBox()
            combo.addItem("Giữ dữ liệu Excel", "CURRENT")
            combo.addItem("Dùng dữ liệu mới", "NEW")
            combo.addItem("Nhập thủ công", "CUSTOM")
            editor = QLineEdit()
            choice = existing.field_choices.get(conflict.conflict_id, {}) if existing else {}
            mode = str(choice.get("mode") or "CURRENT").upper()
            combo.setCurrentIndex(max(0, combo.findData(mode)))
            editor.setText(_display_value(choice.get("value")))
            editor.setEnabled(mode == "CUSTOM")
            combo.currentIndexChanged.connect(
                lambda _index, box=combo, field=editor: field.setEnabled(
                    box.currentData() == "CUSTOM"
                )
            )
            table.setCellWidget(row_index, 3, combo)
            table.setCellWidget(row_index, 4, editor)
            combos[conflict.conflict_id] = combo
            editors[conflict.conflict_id] = editor
        table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeToContents)
        table.horizontalHeader().setStretchLastSection(True)
        self.detail_layout.addWidget(table)
        btn_apply = QPushButton("Áp dụng lựa chọn")
        btn_apply.setProperty("variant", "primary")
        btn_apply.clicked.connect(
            lambda: self._apply_field_choices(issue, combos, editors)
        )
        self.detail_layout.addWidget(btn_apply, alignment=Qt.AlignRight)

    def _show_editor(
        self, issue: ImportIssue, json_index: Optional[int] = None
    ) -> None:
        self._clear_layout(self.detail_layout)
        title = QLabel("Chỉnh sửa dòng bóc tách")
        title.setStyleSheet("font-size: 14pt; font-weight: 700;")
        self.detail_layout.addWidget(title)
        row = self._source_row(issue, json_index)
        editor_form = ExtractRowEditor(row, self)
        self.detail_layout.addWidget(editor_form, stretch=1)
        buttons = QHBoxLayout()
        buttons.addStretch(1)
        btn_back = QPushButton("Quay lại")
        btn_apply = QPushButton("Áp dụng chỉnh sửa")
        btn_apply.setProperty("variant", "primary")
        btn_back.clicked.connect(lambda: self._show_issue(issue))
        btn_apply.clicked.connect(
            lambda: self._apply_edit(issue, editor_form, json_index=json_index)
        )
        buttons.addWidget(btn_back)
        buttons.addWidget(btn_apply)
        self.detail_layout.addLayout(buttons)

    def _apply_edit(
        self,
        issue: ImportIssue,
        editor_form: ExtractRowEditor,
        *,
        json_index: Optional[int] = None,
    ) -> None:
        if json_index is None and not issue.json_indexes:
            return
        rows = self.payload.get("du_lieu_boc_tach") or []
        index = json_index if json_index is not None else issue.json_indexes[0]
        if not (0 <= index < len(rows)):
            return
        try:
            updated = editor_form.updated_row()
        except ValueError as exc:
            notify.warning(self, "Dữ liệu chưa hợp lệ", str(exc))
            return
        rows[index] = updated
        self._dirty = True
        self._queue_reanalysis(issue.issue_id)

    def _select_candidate(self, issue: ImportIssue, candidate_id: str) -> None:
        self.resolutions[issue.issue_id] = IssueResolution(
            issue_id=issue.issue_id,
            action=RESOLUTION_SELECT,
            candidate_id=candidate_id,
        )
        self._dirty = True
        self._queue_reanalysis(issue.issue_id)

    def _skip_issue(self, issue: ImportIssue) -> None:
        self.resolutions[issue.issue_id] = IssueResolution(
            issue_id=issue.issue_id,
            action=RESOLUTION_SKIP,
        )
        self._dirty = True
        self._queue_reanalysis(issue.issue_id)

    def _reset_issue(self, issue: ImportIssue) -> None:
        self.resolutions.pop(issue.issue_id, None)
        self._dirty = True
        self._queue_reanalysis(issue.issue_id)

    def _apply_field_choices(
        self,
        issue: ImportIssue,
        combos: Dict[str, QComboBox],
        editors: Dict[str, QLineEdit],
    ) -> None:
        choices: Dict[str, Dict[str, Any]] = {}
        for conflict in issue.conflicts:
            mode = str(combos[conflict.conflict_id].currentData())
            value: Any = editors[conflict.conflict_id].text().strip()
            if mode == "CUSTOM" and not value:
                notify.warning(
                    self,
                    "Chưa nhập giá trị",
                    f"Hãy nhập giá trị thủ công cho trường “{conflict.field_name}”.",
                )
                return
            reference = (
                conflict.new_value
                if conflict.new_value not in (None, "")
                else conflict.current_value
            )
            if (
                mode == "CUSTOM"
                and isinstance(reference, (int, float))
                and not isinstance(reference, bool)
                and parse_number(value) is None
            ):
                notify.warning(
                    self,
                    "Giá trị chưa hợp lệ",
                    f"Trường “{conflict.field_name}” phải là số.",
                )
                return
            if (
                mode == "CUSTOM"
                and isinstance(reference, (date, datetime))
                and parse_date(value) is None
            ):
                notify.warning(
                    self,
                    "Giá trị chưa hợp lệ",
                    f"Trường “{conflict.field_name}” phải là ngày hợp lệ.",
                )
                return
            choices[conflict.conflict_id] = {"mode": mode, "value": value}
        self.resolutions[issue.issue_id] = IssueResolution(
            issue_id=issue.issue_id,
            action=RESOLUTION_FIELDS,
            field_choices=choices,
        )
        self._dirty = True
        self._queue_reanalysis(issue.issue_id)

    def _queue_reanalysis(self, preferred_issue_id: str = "") -> None:
        """Rebuild after the current Qt signal finishes handling its sender."""
        QTimer.singleShot(
            0,
            lambda issue_id=preferred_issue_id: self._reanalyze(
                preferred_issue_id=issue_id
            ),
        )

    def _reanalyze(self, preferred_issue_id: str = "") -> bool:
        try:
            self.analysis = self.service.analyze(
                self.analysis.output_path,
                self.analysis.daily_path,
                processed_file_id=self.analysis.processed_file_id,
                payload_override=self.payload,
                resolutions=self.resolutions,
            )
        except DailyImportError as exc:
            notify.error(self, "Không phân tích được dữ liệu", str(exc))
            return False
        self.resolutions = dict(self.analysis.resolutions)
        self._refresh_all(preferred_issue_id)
        return True

    def _show_review(self) -> None:
        self._reviewing = True
        self._clear_layout(self.detail_layout)
        title = QLabel("Xác nhận nhập dữ liệu")
        title.setStyleSheet("font-size: 15pt; font-weight: 700;")
        self.detail_layout.addWidget(title)
        text = QLabel(
            "Phần mềm sẽ cập nhật file Hàng ngày với kết quả sau:\n\n"
            f"• Dòng quyết toán mới: {self.analysis.new_info_count}\n"
            f"• Dòng quyết toán cập nhật: {self.analysis.updated_info_count}\n"
            f"• Dòng khoản chi: {len(self.analysis.expense_changes)}\n"
            f"• Dòng để xử lý sau: {self.analysis.skipped_count}\n\n"
            "File Excel chỉ được thay đổi sau khi bấm “Xác nhận nhập”."
        )
        text.setWordWrap(True)
        self.detail_layout.addWidget(text)
        self.detail_layout.addStretch(1)
        self.btn_continue.setText("Xác nhận nhập")

    @Slot()
    def _continue(self) -> None:
        if self.analysis.unresolved_issues:
            return
        if not self._reviewing:
            self._show_review()
            return
        try:
            save_extract_payload(self.analysis.output_path, self.payload)
            final_analysis = self.service.analyze(
                self.analysis.output_path,
                self.analysis.daily_path,
                processed_file_id=self.analysis.processed_file_id,
                resolutions=self.resolutions,
            )
        except DailyImportError as exc:
            notify.error(self, "Không chuẩn bị được lần nhập", str(exc))
            return
        if final_analysis.unresolved_issues:
            self.analysis = final_analysis
            self.resolutions = dict(final_analysis.resolutions)
            self._dirty = True
            self._refresh_all()
            notify.warning(
                self,
                "Dữ liệu đã thay đổi",
                "Một số lựa chọn không còn hợp lệ. Hãy kiểm tra lại các dòng được đánh dấu.",
            )
            return
        self.final_analysis = final_analysis
        self._dirty = False
        self.accept()

    @Slot()
    def _save_and_close(self) -> None:
        try:
            save_extract_payload(self.analysis.output_path, self.payload)
            saved_analysis = self.service.analyze(
                self.analysis.output_path,
                self.analysis.daily_path,
                processed_file_id=self.analysis.processed_file_id,
                resolutions=self.resolutions,
            )
            self.service.save_resolution_draft(
                saved_analysis.output_path,
                saved_analysis.daily_path,
                saved_analysis.resolutions,
            )
            self.service.save_work_in_progress(
                saved_analysis, saved_analysis.resolutions
            )
        except DailyImportError as exc:
            notify.error(self, "Không lưu được xử lý", str(exc))
            return
        self.analysis = saved_analysis
        self.resolutions = dict(saved_analysis.resolutions)
        self.saved_only = True
        self._dirty = False
        self.accept()

    def reject(self) -> None:  # noqa: N802 - override Qt
        if self._dirty:
            confirmed = notify.confirm(
                self,
                "Chưa lưu xử lý",
                "Đóng cửa sổ và bỏ các chỉnh sửa/lựa chọn chưa lưu?",
                confirm_text="Bỏ thay đổi",
                cancel_text="Tiếp tục xử lý",
                tone="warning",
                confirm_variant="danger",
            )
            if not confirmed:
                return
        super().reject()
