"""Biểu mẫu sửa chứng từ và cách đọc cảnh báo từ JSON bóc tách."""

from __future__ import annotations

import re
import unicodedata
from typing import Any, Dict, List

from PySide6.QtWidgets import (
    QComboBox, QFrame, QGridLayout, QHBoxLayout, QLabel,
    QLineEdit, QPushButton, QScrollArea, QVBoxLayout, QWidget,
)

from .daily_import import (
    DOC_BILL, DOC_EXPENSE, DOC_SCALE, classify_doc_type, is_valid_container,
    match_date_of, normalize_container, parse_date, parse_money, parse_number,
)
from .dialogs.notifications import notify


COMMON_FIELDS = (("ngay_dong", "Ngày đóng"), ("so_container", "Số Container"))
TYPE_FIELDS = {
    DOC_SCALE: (
        ("bien_so_xe", "Biển số xe"), ("so_chi_seal", "Số chì/Seal"),
        ("so_tan", "Số tấn"), ("loai_hang", "Loại hàng"),
        ("noi_dong", "Nơi đóng"), ("nguoi_nhan", "Người nhận"),
        ("ten_tau", "Tên tàu"), ("ngay_chay", "Ngày chạy"),
        ("vt_bien", "VT biển"),
    ),
    DOC_BILL: (
        ("so_bill", "Số B/L"), ("ten_tau", "Tên tàu"),
        ("ngay_chay", "Ngày chạy"), ("so_chi_seal", "Số chì/Seal"),
        ("so_tan", "Số tấn"), ("loai_hang", "Loại hàng"),
        ("noi_dong", "Nơi đóng"), ("nguoi_nhan", "Người nhận"),
        ("vt_bien", "VT biển"), ("bien_so_xe", "Biển số xe"),
    ),
    DOC_EXPENSE: (
        ("ngay_chay", "Ngày chạy"), ("loai_hang", "Loại hàng"),
        ("so_hd", "Số HĐ"), ("gia_vat_lieu", "Giá vật liệu"),
        ("don_gia", "Đơn giá"), ("thanh_tien", "Thành tiền"),
        ("vat", "VAT"), ("tong_tien", "Tổng tiền"),
    ),
}
ALL_BUSINESS_FIELDS = tuple(dict((*COMMON_FIELDS, *(field for fields in TYPE_FIELDS.values() for field in fields))).items())
NUMBER_FIELDS = {"so_tan", "gia_vat_lieu", "don_gia", "thanh_tien", "vat", "tong_tien"}
DATE_FIELDS = {"ngay_dong", "ngay_chay"}


def warning_items(value: Any) -> List[Any]:
    if isinstance(value, list):
        return [item for item in value if item not in (None, "", {})]
    return [] if value in (None, "", {}) else [value]


def warning_text(value: Any) -> str:
    if isinstance(value, dict):
        for key in ("noi_dung", "nội_dung", "message", "mo_ta", "ly_do", "canh_bao", "detail"):
            if value.get(key):
                return warning_text(value[key])
        return "; ".join(f"{key.replace('_', ' ')}: {warning_text(item)}" for key, item in value.items() if item not in (None, "")) or "Cần kiểm tra chứng từ"
    if isinstance(value, list):
        return "; ".join(warning_text(item) for item in value if item not in (None, ""))
    return str(value).strip()


def has_low_confidence(row: Dict[str, Any]) -> bool:
    return str(row.get("do_tin_cay") or "").strip().casefold() in {"thấp", "thap"}


def needs_review(row: Dict[str, Any]) -> bool:
    return bool(warning_items(row.get("canh_bao"))) or has_low_confidence(row) or str(row.get("trang_thai") or "OK").strip().upper() != "OK"


WARNING_FIELD_PATTERNS = (
    (r"container|\bcont\b", "so_container"),
    (r"so chi|\bseal\b", "so_chi_seal"),
    (r"ngay dong", "ngay_dong"),
    (r"ngay chay", "ngay_chay"),
    (r"so tan|trong luong", "so_tan"),
    (r"loai hang", "loai_hang"),
    (r"bien so|\bxe\b", "bien_so_xe"),
    (r"noi dong", "noi_dong"),
    (r"nguoi nhan", "nguoi_nhan"),
    (r"ten tau", "ten_tau"),
    (r"van tai bien|vt bien", "vt_bien"),
    (r"so bill|\bbill\b|\bb/l\b", "so_bill"),
    (r"so hoa don|hoa don|so hd", "so_hd"),
    (r"gia vat lieu", "gia_vat_lieu"),
    (r"don gia", "don_gia"),
    (r"thanh tien", "thanh_tien"),
    (r"tong tien", "tong_tien"),
    (r"\bvat\b", "vat"),
)


def _plain_text(value: Any) -> str:
    """Chuẩn hóa dấu để nhận diện trường được nhắc trong lời cảnh báo."""
    folded = unicodedata.normalize("NFD", str(value or "").casefold())
    return "".join(char for char in folded if unicodedata.category(char) != "Mn").replace("đ", "d")


def _changed_fields(before: Dict[str, Any], after: Dict[str, Any]) -> set[str]:
    changed = set()
    for key, _ in ALL_BUSINESS_FIELDS:
        old, new = before.get(key), after.get(key)
        if key in NUMBER_FIELDS:
            parser = parse_number if key == "so_tan" else parse_money
            different = parser(old) != parser(new)
        elif key in DATE_FIELDS:
            different = parse_date(old) != parse_date(new)
        else:
            different = str(old or "").strip() != str(new or "").strip()
        if different:
            changed.add(key)
    if classify_doc_type(before.get("loai_chung_tu")) != classify_doc_type(after.get("loai_chung_tu")):
        changed.add("loai_chung_tu")
    return changed


def _warning_resolved(item: Any, row: Dict[str, Any], changed: set[str]) -> bool:
    text = _plain_text(warning_text(item))
    fields = {field for pattern, field in WARNING_FIELD_PATTERNS if re.search(pattern, text)}
    if not fields and re.search(r"\btien\b|\bchi phi\b", text):
        money_fields = {"gia_vat_lieu", "don_gia", "thanh_tien", "vat", "tong_tien"}
        return bool(money_fields.intersection(changed)) and any(row.get(key) not in (None, "") for key in money_fields)
    if fields:
        if not fields.intersection(changed):
            return False
        for field in fields:
            value = row.get(field)
            if value in (None, ""):
                return False
            if field == "so_container" and not is_valid_container(normalize_container(value)):
                return False
            if field in NUMBER_FIELDS and (parse_number(value) if field == "so_tan" else parse_money(value)) is None:
                return False
        return True
    # Cảnh báo chung không nêu trường cụ thể chỉ được xử lý sau khi sửa dữ liệu.
    return bool(changed)


def _valid_for_ok(row: Dict[str, Any]) -> bool:
    doc_type = classify_doc_type(row.get("loai_chung_tu"))
    if not doc_type or not is_valid_container(normalize_container(row.get("so_container"))):
        return False
    if not match_date_of(row):
        return False
    if doc_type == DOC_EXPENSE:
        total, amount, vat = (parse_money(row.get(key)) for key in ("tong_tien", "thanh_tien", "vat"))
        if all(value is not None for value in (total, amount, vat)) and abs(total - amount - vat) > 1:
            return False
    return True


class ExtractRowEditor(QWidget):
    """Form dùng chung cho Bước 2 và cửa sổ xử lý lỗi nhập."""

    def __init__(self, row: Dict[str, Any], parent=None):
        super().__init__(parent)
        self.original = dict(row)
        self.draft = dict(row)
        self.editors: Dict[str, QLineEdit] = {}

        root = QVBoxLayout(self)
        root.setContentsMargins(0, 0, 0, 0)
        root.setSpacing(10)

        source = QFrame()
        source.setObjectName("extractEditCard")
        source_layout = QVBoxLayout(source)
        source_layout.addWidget(self._section("CHỨNG TỪ"))
        number = row.get("stt_hien_thi")
        source_layout.addWidget(QLabel(f"Dòng {number if number not in (None, '') else '—'}"))
        filename = str(row.get("file_nguon") or "—")
        short_name = filename if len(filename) <= 62 else filename[:38] + "…" + filename[-20:]
        file_button = QPushButton(short_name)
        file_button.setObjectName("extractSourceFile")
        file_button.setToolTip(filename)
        file_button.clicked.connect(lambda: notify.info(self, "Tên file nguồn đầy đủ", filename))
        source_layout.addWidget(file_button)
        type_layout = QHBoxLayout()
        type_layout.addWidget(QLabel("Loại chứng từ:"))
        self.type_combo = QComboBox()
        for title in ("Phiếu cân", "Bill", "Khoản chi"):
            self.type_combo.addItem(title)
        raw_type = str(row.get("loai_chung_tu") or "").strip()
        known = {DOC_SCALE: "Phiếu cân", DOC_BILL: "Bill", DOC_EXPENSE: "Khoản chi"}.get(classify_doc_type(raw_type))
        if known:
            self.type_combo.setCurrentText(known)
        else:
            self.type_combo.insertItem(0, raw_type or "Loại khác")
            self.type_combo.setCurrentIndex(0)
        type_layout.addWidget(self.type_combo, 1)
        source_layout.addLayout(type_layout)
        root.addWidget(source)

        warning_card = QFrame()
        warning_card.setObjectName("extractEditCard")
        warning_layout = QVBoxLayout(warning_card)
        warning_layout.addWidget(self._section("TRẠNG THÁI VÀ CẢNH BÁO"))
        warnings = warning_items(row.get("canh_bao"))
        for item in warnings:
            message = QLabel("• " + warning_text(item))
            message.setWordWrap(True)
            message.setObjectName("extractWarningText")
            warning_layout.addWidget(message)
        if has_low_confidence(row):
            message = QLabel("• Độ tin cậy bóc tách thấp — cần đối chiếu chứng từ")
            message.setWordWrap(True)
            message.setObjectName("extractWarningText")
            warning_layout.addWidget(message)
        status = str(row.get("trang_thai") or "OK").strip() or "OK"
        if status.upper() != "OK":
            message = QLabel("• Trạng thái bóc tách: " + status)
            message.setWordWrap(True)
            message.setObjectName("extractWarningText")
            warning_layout.addWidget(message)
        if not warnings and not has_low_confidence(row) and status.upper() == "OK":
            ok = QLabel("OK · Không có cảnh báo")
            ok.setObjectName("extractWarningOk")
            warning_layout.addWidget(ok)
        root.addWidget(warning_card)

        self.scroll = QScrollArea()
        self.scroll.setWidgetResizable(True)
        self.scroll.setFrameShape(QFrame.NoFrame)
        self.form_host = QWidget()
        self.scroll.setWidget(self.form_host)
        root.addWidget(self.scroll, 1)
        self.type_combo.currentIndexChanged.connect(self._rebuild_fields)
        self._rebuild_fields()

    @staticmethod
    def _section(text: str) -> QLabel:
        heading = QLabel(text)
        heading.setObjectName("extractEditSection")
        return heading

    def _rebuild_fields(self, *_args) -> None:
        for key, editor in self.editors.items():
            self.draft[key] = editor.text()
        self.editors = {}
        host = QWidget()
        layout = QVBoxLayout(host)
        layout.setContentsMargins(0, 0, 4, 4)
        doc_type = classify_doc_type(self.type_combo.currentText())
        self._add_fields(layout, "THÔNG TIN CHUNG", COMMON_FIELDS)
        fields = TYPE_FIELDS.get(doc_type, ALL_BUSINESS_FIELDS)
        self._add_fields(layout, f"THÔNG TIN {self.type_combo.currentText().upper()}", fields)
        layout.addStretch()
        previous = self.scroll.takeWidget()
        self.scroll.setWidget(host)
        if previous is not None:
            previous.deleteLater()
        self.form_host = host

    def _add_fields(self, layout: QVBoxLayout, title: str, fields) -> None:
        card = QFrame()
        card.setObjectName("extractEditCard")
        card_layout = QVBoxLayout(card)
        card_layout.addWidget(self._section(title))
        grid = QGridLayout()
        grid.setHorizontalSpacing(14)
        grid.setVerticalSpacing(9)
        for index, (key, name) in enumerate(fields):
            field = QVBoxLayout()
            field.setSpacing(3)
            field.addWidget(QLabel(name))
            editor = QLineEdit(str(self.draft.get(key) or ""))
            editor.setPlaceholderText("Chưa có dữ liệu")
            field.addWidget(editor)
            grid.addLayout(field, index // 2, index % 2)
            self.editors[key] = editor
        grid.setColumnStretch(0, 1)
        grid.setColumnStretch(1, 1)
        card_layout.addLayout(grid)
        layout.addWidget(card)

    def updated_row(self) -> Dict[str, Any]:
        updated = dict(self.original)
        doc_type = classify_doc_type(self.type_combo.currentText())
        updated["loai_chung_tu"] = self.type_combo.currentText()
        fields = (*COMMON_FIELDS, *TYPE_FIELDS.get(doc_type, ALL_BUSINESS_FIELDS))
        for key, name in fields:
            raw = self.editors[key].text().strip()
            if key in NUMBER_FIELDS:
                parsed = parse_number(raw) if key == "so_tan" else parse_money(raw)
                if raw and parsed is None:
                    raise ValueError(f"{name} phải là số hợp lệ.")
                updated[key] = parsed
            elif key in DATE_FIELDS:
                if raw and parse_date(raw) is None:
                    raise ValueError(f"{name} phải là ngày hợp lệ (ví dụ 26/06/2026).")
                updated[key] = raw or None
            else:
                updated[key] = raw or None

        changed = _changed_fields(self.original, updated)
        valid = _valid_for_ok(updated)
        original_warnings = warning_items(self.original.get("canh_bao"))
        remaining = [item for item in original_warnings if not (valid and _warning_resolved(item, updated, changed))]
        if len(remaining) != len(original_warnings):
            updated["canh_bao"] = remaining
        if changed and valid and not remaining and has_low_confidence(updated):
            updated["do_tin_cay"] = "Đã kiểm tra"
        if changed and valid and not remaining and not has_low_confidence(updated):
            updated["trang_thai"] = "OK"
        elif remaining or has_low_confidence(updated) or not valid:
            if str(updated.get("trang_thai") or "").strip().upper() == "OK":
                updated["trang_thai"] = "Cần kiểm tra"
        return updated
