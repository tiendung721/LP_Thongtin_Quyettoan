"""Phân tích và nhập output GPT vào file quyết toán hằng ngày.

Module này không phụ thuộc giao diện Qt. Giao diện chỉ gọi ``analyze`` để lấy
kế hoạch, thu thập lựa chọn của người dùng rồi gọi ``commit`` để ghi an toàn.

File JSON bóc tách được coi là bộ nhớ tạm của lô. Các dòng chưa rõ được trả về
dưới dạng ``ImportIssue`` kèm ứng viên để giao diện cho người dùng sửa, chọn hoặc
bỏ qua trong lần nhập hiện tại. Excel chỉ được ghi khi mọi issue đã có quyết định.
"""

from __future__ import annotations

import hashlib
import json
import os
import re
import unicodedata
from copy import copy
from dataclasses import asdict, dataclass, field
from datetime import date, datetime
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Sequence, Tuple

from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

from . import file_utils
from .database import Database


INFO_SHEET = "Thong_Tin_Quyet_Toan"
EXPENSE_SHEET = "Khoan_Chi"

INFO_HEADERS = [
    "SQT PM",
    "Ngày Đóng",
    "Số Container",
    "Biển số xe",
    "Số B/L",
    "Số chì/Seal",
    "Lô SX",
    "Số tấn",
    "Loại hàng",
    "Nơi đóng",
    "Tên tàu",
    "Ngày chạy",
    "Dự kiến giao",
    "Người nhận",
    "VT biển",
    "Vận chuyển",
    "HD HP",
    "HD HCM",
    "Ngày Giao",
    "Hóa Đơn quyết toán",
    "Ghi chú",
    "Trạng thái nhập",
    "Ngày nhập cuối",
    "Ngày cập nhật",
    "MD5",
]

INFO_TECHNICAL_HEADERS = {
    "Trạng thái nhập",
    "Ngày nhập cuối",
    "Ngày cập nhật",
    "MD5",
}

INFO_HEADER_ALIASES = {
    "B/L": "Số B/L",
    "Số Bill": "Số B/L",
    "Bill No": "Số B/L",
    "Bill No.": "Số B/L",
    "Số chì": "Số chì/Seal",
    "Hóa Đơn": "Hóa Đơn quyết toán",
    "Mã MD5": "MD5",
}

EXPENSE_HEADERS = [
    "SQT PM",
    "Ngày tháng",
    "Số Container",
    "Nơi đóng",
    "Số HĐ",
    "Giá vật liệu",
    "Đơn giá",
    "Thành tiền",
    "VAT",
    "Tổng tiền",
    "Trạng thái nhập",
    "Ngày cập nhật",
    "MD5",
]

MD5_HEADER_ALIASES = ("MD5", "Mã MD5")

DOC_SCALE = "PHIEU_CAN"
DOC_BILL = "BILL"
DOC_EXPENSE = "KHOAN_CHI"

DOC_LABELS = {
    DOC_SCALE: "Phiếu cân",
    DOC_BILL: "Bill",
    DOC_EXPENSE: "Khoản chi",
}

STATE_READY = "READY"
STATE_COMPLETED = "COMPLETED"
STATE_IGNORED = "IGNORED"

ISSUE_MULTIPLE_TARGETS = "MULTIPLE_TARGETS"
ISSUE_MULTIPLE_BILLS = "MULTIPLE_BILLS"
ISSUE_NO_TARGET = "NO_TARGET"
ISSUE_MISSING_CLOSE_DATE = "MISSING_CLOSE_DATE"
ISSUE_FIELD_CONFLICT = "FIELD_CONFLICT"

RESOLUTION_SELECT = "SELECT"
RESOLUTION_FIELDS = "FIELDS"
RESOLUTION_SKIP = "SKIP"

# Lý do một dòng bóc tách không ghép được với dữ liệu quyết toán.
REASON_NO_TARGET = "Không tìm thấy dòng quyết toán tương ứng"
REASON_MANY_TARGETS = "Khớp nhiều dòng quyết toán, chưa rõ SQT nào"
REASON_MANY_BILLS = "Container có nhiều Bill, chưa rõ Bill nào đúng"
REASON_NO_MONTH = "Không xác định được tháng ghi file do thiếu Ngày Đóng"

_ACCENTED_VIETNAMESE = re.compile(
    r"[ÀÁẢÃẠĂẰẮẲẴẶÂẦẤẨẪẬÈÉẺẼẸÊỀẾỂỄỆÌÍỈĨỊ"
    r"ÒÓỎÕỌÔỒỐỔỖỘƠỜỚỞỠỢÙÚỦŨỤƯỪỨỬỮỰỲÝỶỸỴĐ]",
    re.IGNORECASE,
)

DEFAULT_CARGO_MAPPINGS = {
    "VOI": "Vôi",
    "VOI ROI": "Vôi rời",
    "VOI BOT": "Vôi bột",
    "VOI BOT NONG NGHIEP": "Vôi bột nông nghiệp",
    "BOT NN BAO 25 KG": "Bột NN bao 25 kg",
}


class DailyImportError(RuntimeError):
    """Lỗi nghiệp vụ có nội dung phù hợp để hiển thị cho người dùng."""


@dataclass
class ExtractedRow:
    staged_id: int = 0
    json_index: int = -1
    source_row: int = 0
    row_key: str = ""
    source_name: str = ""
    md5: str = ""
    doc_type: str = ""
    source_status: str = ""
    close_date: Optional[str] = None
    container: str = ""
    tons: Optional[float] = None
    cargo: str = ""
    cargo_recognized: bool = True
    place: str = ""
    recipient: str = ""
    truck_no: str = ""
    vessel: str = ""
    sail_date: Optional[str] = None
    carrier: str = ""
    material_price: Optional[float] = None
    invoice_no: str = ""
    bill_no: str = ""
    seal: str = ""
    unit_price: Optional[float] = None
    amount: Optional[float] = None
    vat: Optional[float] = None
    total: Optional[float] = None
    other: str = ""
    confidence: str = ""
    warning: str = ""
    md5_is_synthetic: bool = False

    def to_dict(self) -> Dict[str, Any]:
        return asdict(self)


@dataclass
class UnmatchedRow:
    """Một dòng bóc tách không ghép được với dữ liệu quyết toán.

    ``json_index`` là vị trí của dòng trong mảng ``du_lieu_boc_tach`` của file
    JSON tạm, để Bước 2 tô sáng đúng dòng và để xóa đúng dòng khi người dùng
    chọn bỏ các dòng lỗi.
    """

    json_index: int
    staged_id: int
    doc_type: str
    container: str
    match_date: Optional[str]
    reference: str
    reason: str

    @property
    def doc_label(self) -> str:
        return DOC_LABELS.get(self.doc_type, self.doc_type or "Chứng từ")


@dataclass
class FieldConflict:
    conflict_id: str
    target_row: int
    field_name: str
    current_value: Any
    new_value: Any
    target_sheet: str = ""


@dataclass
class ImportCandidate:
    """Một dòng Excel hoặc Bill mà người dùng có thể chọn để ghép."""

    candidate_id: str
    kind: str
    label: str
    sheet_name: str = ""
    row_number: int = 0
    sqt: int = 0
    json_index: int = -1
    staged_id: int = 0
    values: Dict[str, Any] = field(default_factory=dict)


@dataclass
class ImportIssue:
    """Một vấn đề cần quyết định trước khi được phép ghi Excel."""

    issue_id: str
    kind: str
    title: str
    reason: str
    doc_type: str
    container: str
    match_date: Optional[str]
    reference: str
    json_indexes: List[int] = field(default_factory=list)
    staged_ids: List[int] = field(default_factory=list)
    candidates: List[ImportCandidate] = field(default_factory=list)
    conflicts: List[FieldConflict] = field(default_factory=list)

    @property
    def doc_label(self) -> str:
        return DOC_LABELS.get(self.doc_type, self.doc_type or "Chứng từ")


@dataclass
class IssueResolution:
    """Quyết định của người dùng cho một :class:`ImportIssue`."""

    issue_id: str
    action: str
    candidate_id: str = ""
    field_choices: Dict[str, Dict[str, Any]] = field(default_factory=dict)

    def to_dict(self) -> Dict[str, Any]:
        return asdict(self)

    @classmethod
    def from_dict(cls, issue_id: str, value: Any) -> Optional["IssueResolution"]:
        if not isinstance(value, dict):
            return None
        action = str(value.get("action") or "").upper()
        if action not in {RESOLUTION_SELECT, RESOLUTION_FIELDS, RESOLUTION_SKIP}:
            return None
        choices = value.get("field_choices")
        return cls(
            issue_id=issue_id,
            action=action,
            candidate_id=str(value.get("candidate_id") or ""),
            field_choices=choices if isinstance(choices, dict) else {},
        )


@dataclass
class InfoChange:
    action: str
    target_row: Optional[int]
    sqt: int
    values: Dict[str, Any]
    staged_ids: List[int]
    document_md5s: List[str]
    check_status: str = "OK"
    conflicts: List[FieldConflict] = field(default_factory=list)
    target_sheet: str = ""


@dataclass
class ExpenseChange:
    action: str
    target_row: Optional[int]
    values: Dict[str, Any]
    staged_ids: List[int]
    document_md5s: List[str]
    check_status: str = "OK"
    conflicts: List[FieldConflict] = field(default_factory=list)


@dataclass
class ImportAnalysis:
    output_path: str
    daily_path: str
    processed_file_id: Optional[int]
    extracted_rows: int = 0
    duplicate_documents: int = 0
    info_changes: List[InfoChange] = field(default_factory=list)
    expense_changes: List[ExpenseChange] = field(default_factory=list)
    unmatched_rows: List[UnmatchedRow] = field(default_factory=list)
    conflicts: List[FieldConflict] = field(default_factory=list)
    pending_count: int = 0
    pending_states: Dict[int, Tuple[str, str]] = field(default_factory=dict)
    completed_staged_ids: List[int] = field(default_factory=list)
    source_hash: str = ""
    daily_hash: str = ""
    issues: List[ImportIssue] = field(default_factory=list)
    resolutions: Dict[str, IssueResolution] = field(default_factory=dict)
    conflict_decisions: Dict[str, bool] = field(default_factory=dict)
    skipped_staged_ids: List[int] = field(default_factory=list)
    skipped_json_indexes: List[int] = field(default_factory=list)

    @property
    def new_info_count(self) -> int:
        return sum(change.action == "CREATE" for change in self.info_changes)

    @property
    def updated_info_count(self) -> int:
        return sum(change.action == "UPDATE" for change in self.info_changes)

    @property
    def has_changes(self) -> bool:
        return bool(self.info_changes or self.expense_changes)

    @property
    def unresolved_issues(self) -> List[ImportIssue]:
        return [issue for issue in self.issues if issue.issue_id not in self.resolutions]

    @property
    def skipped_count(self) -> int:
        return len(set(self.skipped_json_indexes))


@dataclass
class ImportSummary:
    new_info: int = 0
    updated_info: int = 0
    new_expenses: int = 0
    updated_expenses: int = 0
    duplicate_documents: int = 0
    pending_count: int = 0
    backup_path: str = ""
    status: str = "DAILY_IMPORTED"

    def to_dict(self) -> Dict[str, Any]:
        return asdict(self)


@dataclass
class _TargetInfoRow:
    sheet_name: str
    row_number: int
    sqt: int
    close_date: Optional[str]
    container: str
    cargo: str
    values: Dict[str, Any]


@dataclass
class _WorkbookSnapshot:
    info_rows: List[_TargetInfoRow]
    expense_rows: List[Tuple[int, Dict[str, Any]]]
    target_md5: set[str]
    max_sqt: int
    info_sheet_names: List[str] = field(default_factory=list)
    monthly_info: bool = False


def _strip_accents(value: str) -> str:
    normalized = unicodedata.normalize("NFD", value or "")
    return "".join(ch for ch in normalized if unicodedata.category(ch) != "Mn")


def _key_text(value: Any) -> str:
    text = _strip_accents(str(value or "")).upper().replace("Đ", "D")
    return re.sub(r"\s+", " ", text).strip()


def _canonical_header(value: Any) -> str:
    text = re.sub(r"\s+", " ", str(value or "")).strip()
    if not text:
        return ""
    alias_by_key = {_key_text(key): target for key, target in INFO_HEADER_ALIASES.items()}
    return alias_by_key.get(_key_text(text), text)


def month_sheet_number(sheet_name: str) -> Optional[int]:
    match = re.fullmatch(r"THANG\s*(\d{1,2})", _key_text(sheet_name))
    if not match:
        return None
    month = int(match.group(1))
    return month if 1 <= month <= 12 else None


def is_month_info_sheet(sheet_name: str) -> bool:
    return month_sheet_number(sheet_name) is not None


def month_sheet_name(month: int) -> str:
    return f"Tháng {month}"


def month_sheet_name_for_date(value: Optional[str]) -> Optional[str]:
    if not value:
        return None
    try:
        parsed = date.fromisoformat(value)
    except ValueError:
        return None
    return month_sheet_name(parsed.month)


def normalize_container(value: Any) -> str:
    return re.sub(r"[^A-Z0-9]", "", str(value or "").upper())


def is_valid_container(value: str) -> bool:
    return bool(re.fullmatch(r"[A-Z]{4}\d{7}", value or ""))


def normalize_md5(value: Any) -> str:
    return re.sub(r"\s+", "", str(value or "").lower())


def parse_date(value: Any) -> Optional[str]:
    if value in (None, ""):
        return None
    if isinstance(value, datetime):
        return value.date().isoformat()
    if isinstance(value, date):
        return value.isoformat()
    text = str(value).strip()
    for fmt in ("%d/%m/%Y", "%d-%m-%Y", "%Y-%m-%d", "%d/%m/%y"):
        try:
            return datetime.strptime(text, fmt).date().isoformat()
        except ValueError:
            pass
    return None


def excel_date(value: Optional[str]) -> Optional[datetime]:
    if not value:
        return None
    try:
        parsed = date.fromisoformat(value)
    except ValueError:
        return None
    return datetime(parsed.year, parsed.month, parsed.day)


def parse_number(value: Any) -> Optional[float]:
    if value in (None, ""):
        return None
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        return float(value)
    text = str(value).strip().replace(" ", "")
    text = re.sub(r"[^0-9,.-]", "", text)
    if not text:
        return None
    if "," in text and "." in text:
        if text.rfind(",") > text.rfind("."):
            text = text.replace(".", "").replace(",", ".")
        else:
            text = text.replace(",", "")
    elif "," in text:
        parts = text.split(",")
        text = "".join(parts) if len(parts[-1]) == 3 else text.replace(",", ".")
    try:
        return float(text)
    except ValueError:
        return None


def _sqt_key(value: Any) -> str:
    number = parse_number(value)
    if number is not None:
        return str(int(number)) if float(number).is_integer() else format(number, "g")
    return re.sub(r"\s+", " ", str(value or "")).strip()


def normalize_cargo(
    value: Any, mappings: Optional[Dict[str, str]] = None
) -> Tuple[str, bool]:
    original = re.sub(r"\s+", " ", str(value or "")).strip()
    if not original:
        return "", False
    mapping = dict(DEFAULT_CARGO_MAPPINGS)
    mapping.update(mappings or {})
    key = _key_text(original)
    normalized_mapping = {_key_text(k): v for k, v in mapping.items()}
    if key in normalized_mapping:
        return normalized_mapping[key], True
    if _ACCENTED_VIETNAMESE.search(original):
        return original, True
    replacements = [
        (r"^VOI BOT NONG NGHIEP\b", "Vôi bột nông nghiệp"),
        (r"^VOI CUC\b", "Vôi cục"),
        (r"^VOI CANXI\b", "Vôi canxi"),
        (r"^VOI ROI\b", "Vôi rời"),
        (r"^VOI BOT\b", "Vôi bột"),
        (r"^VOI\b", "Vôi"),
    ]
    for pattern, replacement in replacements:
        if re.search(pattern, key):
            tail = re.sub(pattern, "", key).strip()
            tail = tail.replace(" XA ROI", " xá rời").replace(" CO CHAN", " có chặn")
            return (replacement + (" " + tail.lower() if tail else "")).strip(), not bool(tail)
    return original, False


def classify_doc_type(value: Any) -> str:
    """Nhận diện loại chứng từ từ trường ``loai_chung_tu`` của JSON."""
    key = _key_text(value)
    if "PHIEU CAN" in key:
        return DOC_SCALE
    if key == "BILL" or "BILL OF LADING" in key:
        return DOC_BILL
    if "KHOAN CHI" in key or "CHI PHI" in key:
        return DOC_EXPENSE
    return ""


def match_date_of(item: Dict[str, Any]) -> Optional[str]:
    """Ngày dùng để ghép của một dòng JSON bóc tách (dạng ISO).

    Phiếu cân ghép theo ngày đóng; Khoản chi và Bill ưu tiên ngày chạy rồi mới
    tới ngày đóng. Bước 2 hiển thị đúng ngày này để người dùng biết phần mềm
    đang dùng ngày nào khi ghép dữ liệu.
    """
    close_date = parse_date(item.get("ngay_dong"))
    sail_date = parse_date(item.get("ngay_chay"))
    if classify_doc_type(item.get("loai_chung_tu")) == DOC_SCALE:
        return close_date
    return sail_date or close_date


# ---------------------------------------------------------------------------
# File JSON tạm (bộ nhớ tạm của một lô bóc tách)
# ---------------------------------------------------------------------------
def load_extract_payload(path: str) -> Dict[str, Any]:
    """Đọc file JSON bóc tách và kiểm tra cấu trúc tối thiểu."""
    try:
        with open(path, "r", encoding="utf-8-sig") as f:
            payload = json.load(f)
    except json.JSONDecodeError as exc:
        raise DailyImportError(f"File JSON bóc tách không hợp lệ: {exc}") from exc
    except UnicodeDecodeError as exc:
        raise DailyImportError(
            "File JSON bóc tách không đọc được bằng UTF-8. Vui lòng tải lại file."
        ) from exc
    except OSError as exc:
        raise DailyImportError(f"Không đọc được file JSON bóc tách: {exc}") from exc

    if not isinstance(payload, dict):
        raise DailyImportError("File JSON bóc tách phải có object gốc.")
    rows = payload.get("du_lieu_boc_tach")
    if not isinstance(rows, list):
        raise DailyImportError("File JSON thiếu mảng 'du_lieu_boc_tach'.")
    if any(not isinstance(row, dict) for row in rows):
        raise DailyImportError("Mỗi dòng trong 'du_lieu_boc_tach' phải là object JSON.")
    return payload


def save_extract_payload(path: str, payload: Dict[str, Any]) -> None:
    """Ghi nguyên tử file JSON tạm sau khi người dùng chỉnh sửa."""
    refresh_extract_metadata(payload)
    destination = Path(path)
    temp_path = destination.with_name(f".{destination.name}.tmp-{os.getpid()}")
    try:
        with open(temp_path, "w", encoding="utf-8") as f:
            json.dump(payload, f, ensure_ascii=False, indent=2)
            f.write("\n")
        os.replace(temp_path, destination)
    except OSError as exc:
        try:
            if temp_path.exists():
                temp_path.unlink()
        except OSError:
            pass
        raise DailyImportError(f"Không lưu được file JSON bóc tách: {exc}") from exc


def extract_payload_hash(payload: Dict[str, Any]) -> str:
    raw = json.dumps(payload, ensure_ascii=False, sort_keys=True, default=str)
    return hashlib.sha256(raw.encode("utf-8")).hexdigest()


def refresh_extract_metadata(payload: Dict[str, Any]) -> Dict[str, Any]:
    """Tính lại các con số thống kê trong ``metadata`` theo dữ liệu hiện có."""
    rows = [row for row in payload.get("du_lieu_boc_tach") or [] if isinstance(row, dict)]
    metadata = payload.get("metadata")
    if not isinstance(metadata, dict):
        metadata = {}
        payload["metadata"] = metadata

    counts = extract_summary(rows, payload.get("canh_bao"))
    metadata["tong_so_file"] = counts["files"]
    metadata["tong_so_dong_boc_tach"] = counts["total"]
    metadata["so_dong_ok"] = counts["ok"]
    metadata["so_dong_can_kiem_tra"] = counts["need_check"]
    metadata["tong_so_canh_bao"] = counts["warnings"]
    return metadata


def extract_summary(
    rows: Sequence[Dict[str, Any]],
    global_warnings: Any = None,
    error_indexes: Optional[Iterable[int]] = None,
) -> Dict[str, int]:
    """Thống kê nhanh một lô bóc tách để hiển thị ở đầu màn hình Bước 2.

    ``error_indexes`` là các dòng Bước 3 không nhập được; chúng luôn được tính
    vào nhóm “cần kiểm tra” dù trạng thái trong JSON vẫn là OK.
    """
    errors = set(error_indexes or ())
    warnings = len(global_warnings) if isinstance(global_warnings, list) else 0
    counts = {
        "files": len(
            {str(row.get("file_nguon") or "") for row in rows if row.get("file_nguon")}
        ),
        "total": len(rows),
        "ok": 0,
        "need_check": 0,
        "warnings": warnings,
        DOC_SCALE: 0,
        DOC_BILL: 0,
        DOC_EXPENSE: 0,
        "other": 0,
    }
    for index, row in enumerate(rows):
        counts[classify_doc_type(row.get("loai_chung_tu")) or "other"] += 1
        row_warnings = row.get("canh_bao")
        if isinstance(row_warnings, list):
            counts["warnings"] += len(row_warnings)
        elif row_warnings not in (None, ""):
            counts["warnings"] += 1
        if index not in errors and _key_text(row.get("trang_thai")) in ("", "OK"):
            counts["ok"] += 1
        else:
            counts["need_check"] += 1
    return counts


def remove_extract_rows(path: str, indexes: Iterable[int]) -> int:
    """Xóa các dòng theo vị trí khỏi ``du_lieu_boc_tach`` của file JSON tạm.

    Dùng khi người dùng chọn “Hủy các dòng lỗi và nhập các dòng còn lại”: các
    dòng lỗi biến mất hẳn khỏi bộ nhớ tạm thay vì nằm lại hàng chờ.
    Trả về số dòng đã xóa.
    """
    payload = load_extract_payload(path)
    rows = payload.get("du_lieu_boc_tach") or []
    drop = {index for index in indexes if 0 <= index < len(rows)}
    if not drop:
        return 0
    payload["du_lieu_boc_tach"] = [
        row for index, row in enumerate(rows) if index not in drop
    ]
    save_extract_payload(path, payload)
    return len(drop)


def _join_md5(*values: Any) -> str:
    result: List[str] = []
    for value in values:
        for part in re.split(r"[;,\n]+", str(value or "")):
            md5 = normalize_md5(part)
            if md5 and md5 not in result:
                result.append(md5)
    return ";".join(result)


def _md5_cell_value(values: Dict[str, Any]) -> Any:
    for header in MD5_HEADER_ALIASES:
        if values.get(header) not in (None, ""):
            return values.get(header)
    return ""


def _normalize_import_status(value: Any, default: str = "Chưa nhập") -> str:
    key = _key_text(value)
    if key in ("DA NHAP", "HOAN THANH"):
        return "Đã nhập"
    if key == "CHUA NHAP":
        return "Chưa nhập"
    return default


def _stable_row_key(row: ExtractedRow, occurrence: int = 0) -> str:
    identity = {
        "type": row.doc_type,
        "container": row.container,
        "close_date": row.close_date,
        "sail_date": row.sail_date,
        "bill_no": row.bill_no,
        "invoice_no": row.invoice_no,
        "seal": row.seal,
        "cargo": row.cargo,
        "tons": row.tons,
        "amount": row.amount,
        "total": row.total,
        "occurrence": occurrence,
    }
    raw = json.dumps(identity, ensure_ascii=False, sort_keys=True)
    return hashlib.sha1(raw.encode("utf-8")).hexdigest()


class DailyImportService:
    def __init__(
        self,
        database: Database,
        logger,
        cargo_mappings: Optional[Dict[str, str]] = None,
    ):
        self.database = database
        self.logger = logger
        self.cargo_mappings = cargo_mappings or {}

    @staticmethod
    def _normalize_resolutions(values: Optional[Dict[str, Any]]) -> Dict[str, IssueResolution]:
        result: Dict[str, IssueResolution] = {}
        for issue_id, value in (values or {}).items():
            if isinstance(value, IssueResolution):
                result[str(issue_id)] = value
                continue
            resolution = IssueResolution.from_dict(str(issue_id), value)
            if resolution:
                result[str(issue_id)] = resolution
        return result

    @staticmethod
    def _issue_id(kind: str, *parts: Any) -> str:
        raw = "|".join([kind, *(str(part) for part in parts)])
        return hashlib.sha1(raw.encode("utf-8")).hexdigest()

    @staticmethod
    def _target_candidate(row: _TargetInfoRow) -> ImportCandidate:
        raw = (
            f"INFO|{row.sheet_name}|{row.row_number}|{row.sqt}|"
            f"{row.close_date}|{row.container}"
        )
        candidate_id = hashlib.sha1(raw.encode("utf-8")).hexdigest()
        return ImportCandidate(
            candidate_id=candidate_id,
            kind="INFO_ROW",
            label=f"{row.sheet_name} • Dòng {row.row_number} • SQT {row.sqt}",
            sheet_name=row.sheet_name,
            row_number=row.row_number,
            sqt=row.sqt,
            values={
                "Vị trí": f"{row.sheet_name} • Dòng {row.row_number}",
                "SQT PM": row.sqt,
                "Ngày đóng": row.close_date,
                "Container": row.container,
                "Loại hàng": row.cargo,
                "Tên tàu": row.values.get("Tên tàu"),
                "Người nhận": row.values.get("Người nhận"),
            },
        )

    @staticmethod
    def _bill_candidate(row: ExtractedRow) -> ImportCandidate:
        raw = f"BILL|{row.json_index}|{row.md5}|{row.row_key}"
        candidate_id = hashlib.sha1(raw.encode("utf-8")).hexdigest()
        return ImportCandidate(
            candidate_id=candidate_id,
            kind="BILL",
            label=(
                f"Dòng JSON {row.json_index + 1} • "
                f"{row.bill_no or row.vessel or row.source_name or 'Bill'}"
            ),
            json_index=row.json_index,
            staged_id=row.staged_id,
            values={
                "Số B/L": row.bill_no,
                "Container": row.container,
                "Ngày đóng": row.close_date,
                "Ngày chạy": row.sail_date,
                "Tên tàu": row.vessel,
                "VT biển": row.carrier,
                "File nguồn": row.source_name,
            },
        )

    @staticmethod
    def _valid_field_resolution(
        issue: ImportIssue, resolution: IssueResolution
    ) -> bool:
        if resolution.action != RESOLUTION_FIELDS:
            return False
        for conflict in issue.conflicts:
            choice = resolution.field_choices.get(conflict.conflict_id)
            if not isinstance(choice, dict):
                return False
            if str(choice.get("mode") or "").upper() not in {
                "CURRENT",
                "NEW",
                "CUSTOM",
            }:
                return False
            if (
                str(choice.get("mode") or "").upper() == "CUSTOM"
                and choice.get("value") in (None, "")
            ):
                return False
        return True

    @staticmethod
    def _custom_conflict_value(conflict: FieldConflict, value: Any) -> Any:
        reference = (
            conflict.new_value
            if conflict.new_value not in (None, "")
            else conflict.current_value
        )
        if isinstance(reference, (datetime, date)):
            parsed = parse_date(value)
            return excel_date(parsed) if parsed else value
        if isinstance(reference, (int, float)) and not isinstance(reference, bool):
            parsed_number = parse_number(value)
            if parsed_number is None:
                return value
            return int(parsed_number) if isinstance(reference, int) else parsed_number
        return value

    def _register_issue(
        self,
        analysis: ImportAnalysis,
        issue: ImportIssue,
        requested: Dict[str, IssueResolution],
    ) -> Optional[IssueResolution]:
        analysis.issues.append(issue)
        resolution = requested.get(issue.issue_id)
        if resolution is None:
            return None
        if resolution.action == RESOLUTION_SKIP:
            analysis.resolutions[issue.issue_id] = resolution
            analysis.skipped_staged_ids.extend(issue.staged_ids)
            analysis.skipped_json_indexes.extend(issue.json_indexes)
            return resolution
        if resolution.action == RESOLUTION_SELECT and any(
            item.candidate_id == resolution.candidate_id for item in issue.candidates
        ):
            analysis.resolutions[issue.issue_id] = resolution
            return resolution
        if self._valid_field_resolution(issue, resolution):
            analysis.resolutions[issue.issue_id] = resolution
            return resolution
        return None

    def _row_issue(
        self,
        kind: str,
        title: str,
        reason: str,
        primary: ExtractedRow,
        *,
        related: Sequence[ExtractedRow] = (),
        candidates: Sequence[ImportCandidate] = (),
        conflicts: Sequence[FieldConflict] = (),
    ) -> ImportIssue:
        rows = [primary, *related]
        json_indexes = sorted({row.json_index for row in rows if row.json_index >= 0})
        staged_ids = sorted({row.staged_id for row in rows if row.staged_id})
        issue_id = self._issue_id(kind, primary.md5, *json_indexes)
        return ImportIssue(
            issue_id=issue_id,
            kind=kind,
            title=title,
            reason=reason,
            doc_type=primary.doc_type,
            container=primary.container,
            match_date=self._match_date(primary),
            reference=self._reference_text(primary),
            json_indexes=json_indexes,
            staged_ids=staged_ids,
            candidates=list(candidates),
            conflicts=list(conflicts),
        )

    def _apply_conflict_issue(
        self,
        analysis: ImportAnalysis,
        requested: Dict[str, IssueResolution],
        primary: ExtractedRow,
        related: Sequence[ExtractedRow],
        conflicts: Sequence[FieldConflict],
        values: Dict[str, Any],
    ) -> bool:
        if not conflicts:
            return True
        issue = self._row_issue(
            ISSUE_FIELD_CONFLICT,
            "Dữ liệu mới khác Excel",
            "Chọn giá trị sử dụng cho từng trường khác nhau.",
            primary,
            related=related,
            conflicts=conflicts,
        )
        resolution = self._register_issue(analysis, issue, requested)
        if resolution and resolution.action == RESOLUTION_SKIP:
            return False
        if resolution and resolution.action == RESOLUTION_FIELDS:
            for conflict in conflicts:
                choice = resolution.field_choices[conflict.conflict_id]
                mode = str(choice.get("mode") or "CURRENT").upper()
                if mode == "CURRENT":
                    analysis.conflict_decisions[conflict.conflict_id] = False
                else:
                    if mode == "CUSTOM":
                        values[conflict.field_name] = self._custom_conflict_value(
                            conflict, choice.get("value")
                        )
                    analysis.conflict_decisions[conflict.conflict_id] = True
        return True

    # ------------------------------------------------------------------ #
    # Phân tích
    # ------------------------------------------------------------------ #
    def analyze(
        self,
        output_path: str,
        daily_path: str,
        processed_file_id: Optional[int] = None,
        *,
        payload_override: Optional[Dict[str, Any]] = None,
        resolutions: Optional[Dict[str, Any]] = None,
    ) -> ImportAnalysis:
        if not os.path.isfile(output_path):
            raise DailyImportError("Không tìm thấy file JSON bóc tách đang dùng.")
        if not os.path.isfile(daily_path):
            raise DailyImportError(
                "Không tìm thấy file theo dõi hàng ngày. Vui lòng kiểm tra Cài đặt."
            )
        if file_utils.is_file_locked(daily_path):
            raise DailyImportError(
                "File theo dõi hàng ngày đang mở. Vui lòng lưu và đóng Excel rồi thử lại."
            )

        payload = payload_override or load_extract_payload(output_path)
        source_hash = extract_payload_hash(payload)
        daily_hash = file_utils.sha256_file(daily_path)
        requested_resolutions = self._normalize_resolutions(resolutions)
        if resolutions is None:
            draft = self.database.get_import_resolution_draft(output_path)
            if (
                draft
                and draft.get("output_hash") == source_hash
                and os.path.normcase(str(draft.get("daily_path") or ""))
                == os.path.normcase(daily_path)
                and draft.get("daily_hash") == daily_hash
            ):
                requested_resolutions = self._normalize_resolutions(
                    draft.get("resolutions")
                )
            elif draft:
                self.database.delete_import_resolution_draft(output_path)

        parsed = self._read_output(output_path, payload=payload)
        snapshot = self._read_daily_snapshot(daily_path)
        analysis = ImportAnalysis(
            output_path=output_path,
            daily_path=daily_path,
            processed_file_id=processed_file_id,
            extracted_rows=len(parsed),
            source_hash=source_hash,
            daily_hash=daily_hash,
        )

        # JSON tạm là nguồn dữ liệu duy nhất của lần nhập này: dọn hàng chờ cũ
        # rồi chỉ ghi nhận đúng các dòng đang có trong file.
        self.database.clear_pending_staged_rows()
        occurrences: Dict[Tuple[str, str], int] = {}
        for row in parsed:
            if not row.md5:
                continue
            key = (row.md5, _stable_row_key(row))
            occurrences[key] = occurrences.get(key, 0) + 1
            row.row_key = _stable_row_key(row, occurrences[key] - 1)
            self.database.upsert_source_document(
                row.md5,
                row.doc_type,
                row.source_name,
                processed_file_id,
            )
            row.staged_id = self.database.upsert_staged_row(
                row.md5,
                row.row_key,
                row.source_row,
                row.doc_type,
                row.container,
                row.to_dict(),
            )

        # Luôn cho phép xử lý lại toàn bộ dòng trong file JSON hiện tại, kể cả khi
        # DB đã ghi nhận hoàn tất và MD5 vẫn còn trong file theo dõi. Các bước ghép
        # bên dưới sẽ quyết định CREATE hay UPDATE theo dữ liệu đích trong Excel.
        active = [row for row in parsed if row.staged_id]

        bills_by_container: Dict[str, List[ExtractedRow]] = {}
        all_bills_by_container: Dict[str, List[ExtractedRow]] = {}
        scales: List[ExtractedRow] = []
        expenses: List[ExtractedRow] = []
        for row in parsed:
            if row.doc_type == DOC_BILL:
                all_bills_by_container.setdefault(row.container, []).append(row)
        for row in active:
            if row.doc_type == DOC_BILL:
                bills_by_container.setdefault(row.container, []).append(row)
            elif row.doc_type == DOC_SCALE:
                scales.append(row)
            elif row.doc_type == DOC_EXPENSE:
                expenses.append(row)

        next_sqt = snapshot.max_sqt + 1
        used_bills: set[int] = set()
        resolved_bill_containers: set[str] = set()
        blocked_containers: set[str] = set()
        virtual_rows = list(snapshot.info_rows)

        for scale in scales:
            raw_bills = bills_by_container.get(scale.container, [])
            candidates = self._distinct_bills(raw_bills)
            if len(candidates) > 1:
                issue = self._row_issue(
                    ISSUE_MULTIPLE_BILLS,
                    "Có nhiều Bill phù hợp",
                    REASON_MANY_BILLS,
                    scale,
                    related=raw_bills,
                    candidates=[self._bill_candidate(row) for row in candidates],
                )
                resolution = self._register_issue(
                    analysis, issue, requested_resolutions
                )
                if resolution is None:
                    self._mark_unmatched(analysis, scale, REASON_MANY_BILLS)
                    for bill in raw_bills:
                        self._mark_unmatched(analysis, bill, REASON_MANY_BILLS)
                    blocked_containers.add(scale.container)
                    continue
                if resolution.action == RESOLUTION_SKIP:
                    blocked_containers.add(scale.container)
                    continue
                selected_candidate = next(
                    item
                    for item in issue.candidates
                    if item.candidate_id == resolution.candidate_id
                )
                candidates = [
                    row for row in candidates if row.json_index == selected_candidate.json_index
                ]
                for bill in raw_bills:
                    if bill.json_index != selected_candidate.json_index:
                        analysis.skipped_staged_ids.append(bill.staged_id)
                        analysis.skipped_json_indexes.append(bill.json_index)
                        used_bills.add(bill.staged_id)

            # Container khớp nhiều dòng quyết toán: Bill của container cũng sẽ khớp
            # nhiều dòng như vậy nên cứ để vòng Bill bên dưới báo lỗi cho nó.
            target_options = self._scale_target_options(virtual_rows, scale)
            if len(target_options) > 1:
                related = [candidates[0]] if candidates else []
                issue = self._row_issue(
                    ISSUE_MULTIPLE_TARGETS,
                    "Khớp nhiều dòng quyết toán",
                    REASON_MANY_TARGETS,
                    scale,
                    related=related,
                    candidates=[self._target_candidate(row) for row in target_options],
                )
                resolution = self._register_issue(
                    analysis, issue, requested_resolutions
                )
                if resolution is None:
                    self._mark_unmatched(analysis, scale, REASON_MANY_TARGETS)
                    for bill in related:
                        self._mark_unmatched(analysis, bill, REASON_MANY_TARGETS)
                        used_bills.add(bill.staged_id)
                    continue
                for bill in related:
                    used_bills.add(bill.staged_id)
                if resolution.action == RESOLUTION_SKIP:
                    continue
                selected_candidate = next(
                    item
                    for item in issue.candidates
                    if item.candidate_id == resolution.candidate_id
                )
                target_options = [
                    row
                    for row in target_options
                    if row.sheet_name == selected_candidate.sheet_name
                    and row.row_number == selected_candidate.row_number
                    and row.sqt == selected_candidate.sqt
                ]

            selected_bill = candidates[0] if candidates else None
            target = target_options[0] if target_options else None
            fallback = target.values if target else {}
            target_sheet = (
                target.sheet_name
                if target
                else self._target_info_sheet(snapshot, scale.close_date)
            )
            if not target_sheet:
                related = [selected_bill] if selected_bill else []
                issue = self._row_issue(
                    ISSUE_MISSING_CLOSE_DATE,
                    "Thiếu Ngày Đóng",
                    REASON_NO_MONTH,
                    scale,
                    related=related,
                )
                resolution = self._register_issue(
                    analysis, issue, requested_resolutions
                )
                if resolution is None:
                    self._mark_unmatched(analysis, scale, REASON_NO_MONTH)
                    if selected_bill:
                        self._mark_unmatched(analysis, selected_bill, REASON_NO_MONTH)
                continue
            sqt = target.sqt if target else next_sqt
            action = "UPDATE" if target else "CREATE"
            if not target:
                next_sqt += 1
            previous_same_container = any(
                item.container == scale.container
                and item.close_date != scale.close_date
                for item in snapshot.info_rows
            )
            check_status = (
                "Cần kiểm tra"
                if previous_same_container
                or not scale.cargo_recognized
                # Tạo QT mới mà không có Bill (thông tin tàu nhập tay) -> đánh dấu
                # để người dùng soát lại cho an toàn.
                or (not selected_bill and not target)
                or self._source_needs_review(scale, selected_bill)
                else "OK"
            )
            values = self._info_values(scale, selected_bill, sqt, fallback)
            conflicts = self._find_conflicts(target, values) if target else []
            if conflicts:
                check_status = "Cần kiểm tra"
                analysis.conflicts.extend(conflicts)
                if not self._apply_conflict_issue(
                    analysis,
                    requested_resolutions,
                    scale,
                    [selected_bill] if selected_bill else [],
                    conflicts,
                    values,
                ):
                    continue
            staged_ids = [scale.staged_id]
            md5_values = [self._real_md5(scale)]
            if selected_bill:
                staged_ids.append(selected_bill.staged_id)
                md5_values.append(self._real_md5(selected_bill))
                used_bills.add(selected_bill.staged_id)
                resolved_bill_containers.add(scale.container)
                analysis.pending_states[selected_bill.staged_id] = (STATE_READY, "")
            analysis.info_changes.append(
                InfoChange(
                    action=action,
                    target_row=target.row_number if target else None,
                    sqt=sqt,
                    values=values,
                    staged_ids=staged_ids,
                    document_md5s=md5_values,
                    check_status=check_status,
                    conflicts=conflicts,
                    target_sheet=target_sheet,
                )
            )
            analysis.pending_states[scale.staged_id] = (STATE_READY, "")
            analysis.completed_staged_ids.extend(staged_ids)
            if action == "CREATE":
                virtual_rows.append(
                    _TargetInfoRow(
                        sheet_name=target_sheet,
                        row_number=0,
                        sqt=sqt,
                        close_date=scale.close_date,
                        container=scale.container,
                        cargo=scale.cargo,
                        values=values,
                    )
                )

        # Bill có thể bổ sung trực tiếp cho một dòng quyết toán đã tồn tại.
        for container, raw_candidates in bills_by_container.items():
            if container in blocked_containers:
                continue
            remaining = [
                row for row in raw_candidates if row.staged_id not in used_bills
            ]
            if not remaining:
                continue
            if container in resolved_bill_containers:
                # Cùng một Bill (trùng MD5) đã ghép với Phiếu cân ở trên.
                for bill in remaining:
                    analysis.pending_states[bill.staged_id] = (STATE_READY, "")
                    analysis.completed_staged_ids.append(bill.staged_id)
                continue

            candidates = self._distinct_bills(remaining)
            all_candidates = self._distinct_bills(
                all_bills_by_container.get(container, remaining)
            )
            if len(all_candidates) > 1:
                issue = self._row_issue(
                    ISSUE_MULTIPLE_BILLS,
                    "Có nhiều Bill phù hợp",
                    REASON_MANY_BILLS,
                    candidates[0],
                    related=candidates[1:],
                    candidates=[self._bill_candidate(row) for row in all_candidates],
                )
                resolution = self._register_issue(
                    analysis, issue, requested_resolutions
                )
                if resolution is None:
                    for bill in remaining:
                        self._mark_unmatched(analysis, bill, REASON_MANY_BILLS)
                    continue
                if resolution.action == RESOLUTION_SKIP:
                    continue
                selected_candidate = next(
                    item
                    for item in issue.candidates
                    if item.candidate_id == resolution.candidate_id
                )
                selected_bill = next(
                    (
                        row
                        for row in candidates
                        if row.json_index == selected_candidate.json_index
                    ),
                    None,
                )
                if selected_bill is None:
                    analysis.skipped_staged_ids.extend(
                        row.staged_id for row in candidates if row.staged_id
                    )
                    analysis.skipped_json_indexes.extend(
                        row.json_index for row in candidates if row.json_index >= 0
                    )
                    continue
                for bill in candidates:
                    if bill is not selected_bill:
                        analysis.skipped_staged_ids.append(bill.staged_id)
                        analysis.skipped_json_indexes.append(bill.json_index)
            else:
                selected_bill = candidates[0]

            target_rows = [row for row in virtual_rows if row.container == container]
            if len(target_rows) > 1:
                issue = self._row_issue(
                    ISSUE_MULTIPLE_TARGETS,
                    "Khớp nhiều dòng quyết toán",
                    REASON_MANY_TARGETS,
                    selected_bill,
                    candidates=[self._target_candidate(row) for row in target_rows],
                )
                resolution = self._register_issue(
                    analysis, issue, requested_resolutions
                )
                if resolution is None:
                    self._mark_unmatched(analysis, selected_bill, REASON_MANY_TARGETS)
                    continue
                if resolution.action == RESOLUTION_SKIP:
                    continue
                selected_candidate = next(
                    item
                    for item in issue.candidates
                    if item.candidate_id == resolution.candidate_id
                )
                target_rows = [
                    row
                    for row in target_rows
                    if row.sheet_name == selected_candidate.sheet_name
                    and row.row_number == selected_candidate.row_number
                    and row.sqt == selected_candidate.sqt
                ]
            target = target_rows[0] if target_rows else None
            target_sheet = (
                target.sheet_name
                if target
                else self._target_info_sheet(snapshot, selected_bill.close_date)
            )
            if not target_sheet:
                issue = self._row_issue(
                    ISSUE_MISSING_CLOSE_DATE,
                    "Thiếu Ngày Đóng",
                    REASON_NO_MONTH,
                    selected_bill,
                )
                resolution = self._register_issue(
                    analysis, issue, requested_resolutions
                )
                if resolution is None:
                    self._mark_unmatched(analysis, selected_bill, REASON_NO_MONTH)
                continue
            action = "UPDATE" if target else "CREATE"
            sqt = target.sqt if target else next_sqt
            if not target:
                next_sqt += 1
            fallback = target.values if target else {}
            values = self._bill_info_values(selected_bill, sqt, fallback)
            conflicts = self._find_conflicts(target, values) if target else []
            analysis.conflicts.extend(conflicts)
            if conflicts and not self._apply_conflict_issue(
                analysis,
                requested_resolutions,
                selected_bill,
                [],
                conflicts,
                values,
            ):
                continue
            analysis.info_changes.append(
                InfoChange(
                    action=action,
                    target_row=target.row_number if target else None,
                    sqt=sqt,
                    values=values,
                    staged_ids=[selected_bill.staged_id],
                    document_md5s=[self._real_md5(selected_bill)],
                    check_status="OK",
                    conflicts=conflicts,
                    target_sheet=target_sheet,
                )
            )
            if action == "CREATE":
                virtual_rows.append(
                    _TargetInfoRow(
                        sheet_name=target_sheet,
                        row_number=0,
                        sqt=sqt,
                        close_date=selected_bill.close_date,
                        container=selected_bill.container,
                        cargo=selected_bill.cargo,
                        values=values,
                    )
                )
            used_bills.add(selected_bill.staged_id)
            analysis.pending_states[selected_bill.staged_id] = (STATE_READY, "")
            analysis.completed_staged_ids.append(selected_bill.staged_id)

        self._analyze_expenses(
            expenses,
            virtual_rows,
            snapshot,
            analysis,
            requested_resolutions,
        )

        for row_id, (state, note) in analysis.pending_states.items():
            self.database.update_staged_row(row_id, state=state, note=note)
        # Dòng chưa xử lý hoặc được bỏ qua chỉ trong lần này phải nằm lại hàng chờ.
        # Nhờ vậy lô được đánh dấu nhập một phần và người dùng có thể xử lý lại sau.
        unresolved_staged = {
            staged_id
            for issue in analysis.unresolved_issues
            for staged_id in issue.staged_ids
        } | set(analysis.skipped_staged_ids)
        for row in active:
            if row.staged_id in unresolved_staged or row.staged_id not in analysis.pending_states:
                self.database.update_staged_row(
                    row.staged_id,
                    state="PENDING",
                    note="Còn chờ xử lý ở Bước 3",
                )
        for md5 in {row.md5 for row in parsed if row.md5}:
            self.database.refresh_document_status(md5)
        analysis.pending_count = len(
            {
                staged_id
                for issue in analysis.unresolved_issues
                for staged_id in issue.staged_ids
            }
            | set(analysis.skipped_staged_ids)
        )
        return analysis

    # ------------------------------------------------------------------ #
    # Ghi workbook
    # ------------------------------------------------------------------ #
    def commit(
        self,
        analysis: ImportAnalysis,
        conflict_decisions: Optional[Dict[str, bool]] = None,
    ) -> ImportSummary:
        supplied_conflict_decisions = conflict_decisions or {}
        blocking_issues = [
            issue
            for issue in analysis.unresolved_issues
            if issue.kind != ISSUE_FIELD_CONFLICT
            or not all(
                conflict.conflict_id in supplied_conflict_decisions
                for conflict in issue.conflicts
            )
        ]
        if blocking_issues:
            raise DailyImportError(
                "Còn dòng chưa được xử lý. Vui lòng xử lý hoặc bỏ qua tất cả trước khi nhập."
            )
        conflict_decisions = {
            **analysis.conflict_decisions,
            **supplied_conflict_decisions,
        }
        daily_path = analysis.daily_path
        if file_utils.is_file_locked(daily_path):
            raise DailyImportError(
                "File theo dõi hàng ngày đang mở. Vui lòng đóng Excel rồi thử lại."
            )
        current_daily_hash = file_utils.sha256_file(daily_path)
        if analysis.daily_hash and current_daily_hash != analysis.daily_hash:
            raise DailyImportError(
                "File theo dõi hàng ngày đã thay đổi sau khi phân tích. "
                "Vui lòng mở lại cửa sổ xử lý để chọn lại dòng đích."
            )
        current_source_hash = extract_payload_hash(load_extract_payload(analysis.output_path))
        if analysis.source_hash and current_source_hash != analysis.source_hash:
            raise DailyImportError(
                "Dữ liệu bóc tách đã thay đổi sau khi phân tích. "
                "Vui lòng mở lại cửa sổ xử lý trước khi nhập."
            )
        run_id = self.database.create_import_run(
            analysis.processed_file_id, analysis.output_path, daily_path
        )
        backup_path = ""
        temp_path = ""
        try:
            wb = load_workbook(daily_path)
            self._upgrade_workbook(wb)
            expense_ws = wb[EXPENSE_SHEET]
            info_sheet_names = self._info_sheet_names(wb)
            default_info_sheet = info_sheet_names[0] if info_sheet_names else INFO_SHEET
            expense_map = self._header_map(expense_ws)

            for change in analysis.info_changes:
                sheet_name = change.target_sheet or default_info_sheet
                if sheet_name not in wb.sheetnames and is_month_info_sheet(sheet_name):
                    self._ensure_month_sheet(wb, sheet_name)
                    self._upgrade_workbook(wb)
                    expense_ws = wb[EXPENSE_SHEET]
                    expense_map = self._header_map(expense_ws)
                if sheet_name not in wb.sheetnames:
                    raise DailyImportError(f"File theo dõi thiếu sheet '{sheet_name}'.")
                info_ws = wb[sheet_name]
                info_map = self._header_map(info_ws)
                row_number = (
                    change.target_row
                    if change.action == "UPDATE" and change.target_row
                    else self._last_data_row(info_ws) + 1
                )
                if change.action == "CREATE":
                    self._copy_row_style(
                        info_ws, max(2, self._last_data_row(info_ws)), row_number
                    )
                conflict_by_field = {
                    conflict.field_name: conflict for conflict in change.conflicts
                }
                for name, value in change.values.items():
                    if name not in info_map:
                        continue
                    cell = info_ws.cell(row_number, info_map[name])
                    conflict = conflict_by_field.get(name)
                    if conflict:
                        if conflict_decisions.get(conflict.conflict_id, False):
                            cell.value = value
                    elif change.action == "CREATE" or cell.value in (None, "") or name in (
                        "MD5",
                        "Ngày cập nhật",
                    ):
                        cell.value = value

            for change in analysis.expense_changes:
                row_number = (
                    change.target_row
                    if change.action == "UPDATE" and change.target_row
                    else self._last_data_row(expense_ws) + 1
                )
                if change.action == "CREATE":
                    self._copy_row_style(
                        expense_ws, max(2, self._last_data_row(expense_ws)), row_number
                    )
                conflict_by_field = {
                    conflict.field_name: conflict for conflict in change.conflicts
                }
                for name, value in change.values.items():
                    if name not in expense_map:
                        continue
                    cell = expense_ws.cell(row_number, expense_map[name])
                    conflict = conflict_by_field.get(name)
                    if conflict:
                        if conflict_decisions.get(conflict.conflict_id, False):
                            cell.value = value
                    elif change.action == "CREATE" or cell.value in (None, "") or name in (
                        "MD5",
                        "Nơi đóng",
                        "Ngày cập nhật",
                    ):
                        cell.value = value

            self._sync_expense_places(wb, expense_ws)
            self._format_workbook(wb)
            suffix = Path(daily_path).suffix or ".xlsx"
            temp_path = str(Path(daily_path).with_name(f".__daily_import_{os.getpid()}{suffix}"))
            wb.save(temp_path)
            wb.close()
            self._validate_saved_workbook(temp_path)
            os.replace(temp_path, daily_path)

            completed_ids = sorted(set(analysis.completed_staged_ids))
            affected_md5: set[str] = set()
            for row_id in completed_ids:
                item = self.database.get_staged_row(row_id)
                if not item:
                    continue
                affected_md5.add(item.get("document_md5") or "")
                self.database.update_staged_row(row_id, state=STATE_COMPLETED)
            for md5 in affected_md5:
                if md5:
                    self.database.refresh_document_status(md5)

            pending = self.database.count_pending_rows()
            status = "DAILY_IMPORT_PARTIAL" if pending else "DAILY_IMPORTED"
            summary = ImportSummary(
                new_info=analysis.new_info_count,
                updated_info=analysis.updated_info_count,
                new_expenses=sum(c.action == "CREATE" for c in analysis.expense_changes),
                updated_expenses=sum(c.action == "UPDATE" for c in analysis.expense_changes),
                duplicate_documents=analysis.duplicate_documents,
                pending_count=pending,
                backup_path=backup_path,
                status=status,
            )
            self.database.finish_import_run(
                run_id, "COMPLETED", backup_path=backup_path, summary=summary.to_dict()
            )
            if analysis.processed_file_id is not None:
                self.database.update_status(
                    analysis.processed_file_id,
                    status,
                    note=self._summary_note(summary),
                )
            self.database.delete_import_resolution_draft(analysis.output_path)
            return summary
        except Exception as exc:
            # Ghi theo cơ chế file tạm + os.replace (nguyên tử) nên file theo dõi
            # gốc không bị hỏng dở; chỉ cần dọn file tạm nếu lỗi.
            if temp_path and os.path.exists(temp_path):
                try:
                    os.remove(temp_path)
                except OSError:
                    pass
            self.database.finish_import_run(
                run_id,
                "FAILED",
                backup_path=backup_path or None,
                error_message=str(exc),
            )
            if analysis.processed_file_id is not None:
                self.database.update_status(
                    analysis.processed_file_id,
                    "DAILY_IMPORT_ERROR",
                    note=f"Lỗi nhập file theo dõi: {exc}",
                )
            if isinstance(exc, DailyImportError):
                raise
            raise DailyImportError(f"Không cập nhật được file theo dõi: {exc}") from exc

    def save_resolution_draft(
        self,
        output_path: str,
        daily_path: str,
        resolutions: Dict[str, IssueResolution],
    ) -> None:
        payload = load_extract_payload(output_path)
        self.database.save_import_resolution_draft(
            output_path,
            extract_payload_hash(payload),
            daily_path,
            file_utils.sha256_file(daily_path),
            {
                issue_id: resolution.to_dict()
                for issue_id, resolution in resolutions.items()
            },
        )

    # ------------------------------------------------------------------ #
    # Đọc dữ liệu
    # ------------------------------------------------------------------ #
    def _read_output(
        self, path: str, *, payload: Optional[Dict[str, Any]] = None
    ) -> List[ExtractedRow]:
        if Path(path).suffix.lower() != ".json":
            raise DailyImportError("File bóc tách phải là định dạng JSON (.json).")

        payload = payload or load_extract_payload(path)
        rows = payload.get("du_lieu_boc_tach") or []

        result: List[ExtractedRow] = []
        for index, item in enumerate(rows, start=1):
            if not any(value not in (None, "", [], {}) for value in item.values()):
                continue

            doc_type = classify_doc_type(item.get("loai_chung_tu"))
            if not doc_type:
                continue

            cargo, recognized = normalize_cargo(
                item.get("loai_hang"),
                self.cargo_mappings,
            )
            close_date = parse_date(item.get("ngay_dong"))
            sail_date = parse_date(item.get("ngay_chay"))
            source_row_number = parse_number(item.get("stt_hien_thi"))
            row = ExtractedRow(
                json_index=index - 1,
                source_row=int(source_row_number or index),
                source_name=str(item.get("file_nguon") or ""),
                md5=normalize_md5(item.get("ma_md5_file")),
                doc_type=doc_type,
                source_status=str(item.get("trang_thai") or ""),
                close_date=close_date,
                container=normalize_container(item.get("so_container")),
                tons=parse_number(item.get("so_tan")),
                cargo=cargo,
                cargo_recognized=recognized,
                place=str(item.get("noi_dong") or "").strip(),
                recipient=str(item.get("nguoi_nhan") or "").strip(),
                truck_no=str(item.get("bien_so_xe") or "").strip(),
                vessel=str(item.get("ten_tau") or "").strip(),
                sail_date=sail_date,
                carrier=str(item.get("vt_bien") or "").strip(),
                material_price=parse_number(item.get("gia_vat_lieu")),
                invoice_no=str(item.get("so_hd") or "").strip(),
                bill_no=str(item.get("so_bill") or "").strip(),
                seal=str(item.get("so_chi_seal") or "").strip(),
                unit_price=parse_number(item.get("don_gia")),
                amount=parse_number(item.get("thanh_tien")),
                vat=parse_number(item.get("vat")),
                total=parse_number(item.get("tong_tien")),
                other=self._json_cell_text(item.get("truong_khac")),
                confidence=str(item.get("do_tin_cay") or "").strip(),
                warning=self._json_warning_text(item.get("canh_bao")),
            )
            if doc_type == DOC_EXPENSE:
                row.close_date = sail_date or close_date
            if not row.md5:
                # Không có MD5 thì vẫn theo dõi được trong danh sách chờ, nhưng
                # dùng khóa cục bộ để không làm mất dòng.
                row.md5 = "missing-" + hashlib.sha1(
                    f"{path}|{index}|{row.container}".encode("utf-8")
                ).hexdigest()
                row.md5_is_synthetic = True
            result.append(row)
        return result

    @staticmethod
    def _json_cell_text(value: Any) -> str:
        if value in (None, ""):
            return ""
        if isinstance(value, (dict, list)):
            return json.dumps(value, ensure_ascii=False, sort_keys=True)
        return str(value).strip()

    @classmethod
    def _json_warning_text(cls, value: Any) -> str:
        if value in (None, ""):
            return ""
        if isinstance(value, list):
            return "; ".join(
                cls._json_cell_text(item) for item in value if item not in (None, "")
            )
        return cls._json_cell_text(value)

    def _read_daily_snapshot(self, path: str) -> _WorkbookSnapshot:
        try:
            wb = load_workbook(path, read_only=True, data_only=False)
        except Exception as exc:
            raise DailyImportError(f"Không đọc được file theo dõi hàng ngày: {exc}") from exc
        try:
            info_sheet_names = self._info_sheet_names(wb)
            if not info_sheet_names:
                raise DailyImportError(
                    f"File theo dõi thiếu sheet thông tin '{INFO_SHEET}' hoặc các sheet 'Tháng N'."
                )
            monthly_info = any(is_month_info_sheet(name) for name in info_sheet_names)
            info_rows: List[_TargetInfoRow] = []
            target_md5: set[str] = set()
            max_sqt = 0
            for sheet_name in info_sheet_names:
                info_ws = wb[sheet_name]
                info_map = self._validate_daily_headers(
                    info_ws, INFO_HEADERS, info_sheet=True
                )
                for row_number, row_values in enumerate(
                    info_ws.iter_rows(min_row=2, values_only=True),
                    start=2,
                ):
                    values = {
                        name: row_values[col - 1] if col - 1 < len(row_values) else None
                        for name, col in info_map.items()
                    }
                    if not any(value not in (None, "") for value in values.values()):
                        continue
                    sqt_value = parse_number(values.get("SQT PM"))
                    sqt = int(sqt_value or 0)
                    max_sqt = max(max_sqt, sqt)
                    container = normalize_container(values.get("Số Container"))
                    cargo, _ = normalize_cargo(
                        values.get("Loại hàng"), self.cargo_mappings
                    )
                    info_rows.append(
                        _TargetInfoRow(
                            sheet_name=sheet_name,
                            row_number=row_number,
                            sqt=sqt,
                            close_date=parse_date(values.get("Ngày Đóng")),
                            container=container,
                            cargo=cargo,
                            values=values,
                        )
                    )
                    target_md5.update(
                        normalize_md5(part)
                        for part in re.split(
                            r"[;,\n]+", str(_md5_cell_value(values) or "")
                        )
                        if normalize_md5(part)
                    )

            expense_rows: List[Tuple[int, Dict[str, Any]]] = []
            if EXPENSE_SHEET in wb.sheetnames:
                expense_ws = wb[EXPENSE_SHEET]
                expense_map = self._validate_daily_headers(
                    expense_ws, EXPENSE_HEADERS, info_sheet=False
                )
                for row_number, row_values in enumerate(
                    expense_ws.iter_rows(min_row=2, values_only=True),
                    start=2,
                ):
                    values = {
                        name: row_values[col - 1] if col - 1 < len(row_values) else None
                        for name, col in expense_map.items()
                    }
                    if not any(value not in (None, "") for value in values.values()):
                        continue
                    expense_rows.append((row_number, values))
                    target_md5.update(
                        normalize_md5(part)
                        for part in re.split(
                            r"[;,\n]+", str(_md5_cell_value(values) or "")
                        )
                        if normalize_md5(part)
                    )
            return _WorkbookSnapshot(
                info_rows,
                expense_rows,
                target_md5,
                max_sqt,
                info_sheet_names=info_sheet_names,
                monthly_info=monthly_info,
            )
        finally:
            wb.close()

    # ------------------------------------------------------------------ #
    # Matching helpers
    # ------------------------------------------------------------------ #
    @staticmethod
    def _target_group(
        rows: Sequence[_TargetInfoRow], close_date: Optional[str], container: str
    ) -> List[_TargetInfoRow]:
        return [
            row
            for row in rows
            if row.close_date == close_date and row.container == container
        ]

    @staticmethod
    def _real_md5(row: ExtractedRow) -> str:
        return "" if row.md5_is_synthetic or row.md5.startswith("missing-") else row.md5

    def _scale_target_options(
        self, rows: Sequence[_TargetInfoRow], scale: ExtractedRow
    ) -> List[_TargetInfoRow]:
        return self._target_group(rows, scale.close_date, scale.container)

    @staticmethod
    def _distinct_bills(rows: Iterable[ExtractedRow]) -> List[ExtractedRow]:
        result: List[ExtractedRow] = []
        seen = set()
        for row in rows:
            if row.md5 not in seen:
                seen.add(row.md5)
                result.append(row)
        return result

    @staticmethod
    def _match_date(row: ExtractedRow) -> Optional[str]:
        """Ngày dùng để ghép của một dòng đã bóc tách (Bill ưu tiên ngày chạy)."""
        if row.doc_type == DOC_BILL:
            return row.sail_date or row.close_date
        return row.close_date

    @staticmethod
    def _reference_text(row: ExtractedRow) -> str:
        """Số HĐ/Bill/Seal để người dùng nhận ra dòng lỗi trong file gốc."""
        if row.doc_type == DOC_EXPENSE and row.invoice_no:
            return f"HĐ {row.invoice_no}"
        if row.doc_type == DOC_BILL and row.bill_no:
            return f"Bill {row.bill_no}"
        if row.doc_type == DOC_SCALE and row.seal:
            return f"Seal {row.seal}"
        return row.invoice_no or row.bill_no or row.seal or ""

    def _mark_unmatched(
        self, analysis: ImportAnalysis, row: ExtractedRow, reason: str
    ) -> None:
        analysis.unmatched_rows.append(
            UnmatchedRow(
                json_index=row.json_index,
                staged_id=row.staged_id,
                doc_type=row.doc_type,
                container=row.container,
                match_date=self._match_date(row),
                reference=self._reference_text(row),
                reason=reason,
            )
        )

    @staticmethod
    def _target_info_sheet(
        snapshot: _WorkbookSnapshot, close_date: Optional[str]
    ) -> Optional[str]:
        if snapshot.monthly_info:
            return month_sheet_name_for_date(close_date)
        return snapshot.info_sheet_names[0] if snapshot.info_sheet_names else INFO_SHEET

    @staticmethod
    def _source_needs_review(*rows: Optional[ExtractedRow]) -> bool:
        for row in rows:
            if row is None:
                continue
            if row.warning:
                return True
            if _key_text(row.source_status) not in ("", "OK"):
                return True
            if _key_text(row.confidence) == "THAP":
                return True
        return False

    def _info_values(
        self,
        scale: ExtractedRow,
        bill: Optional[ExtractedRow],
        sqt: int,
        fallback: Dict[str, Any],
    ) -> Dict[str, Any]:
        vessel = (
            (bill.vessel if bill else "")
            or str(fallback.get("Tên tàu") or "")
            or (scale.vessel or "")
        )
        sail_date = (
            (bill.sail_date if bill else None)
            or parse_date(fallback.get("Ngày chạy"))
            or scale.sail_date
        )
        carrier = (
            (bill.carrier if bill else "")
            or str(fallback.get("VT biển") or "")
            or (scale.carrier or "")
        )
        seal = scale.seal or (bill.seal if bill else "") or fallback.get("Số chì/Seal")
        return {
            "SQT PM": sqt,
            "Ngày Đóng": excel_date(scale.close_date),
            "Số Container": scale.container,
            "Biển số xe": scale.truck_no or fallback.get("Biển số xe"),
            "Số B/L": (bill.bill_no if bill else "") or fallback.get("Số B/L"),
            "Số chì/Seal": seal,
            "Số tấn": scale.tons,
            "Loại hàng": scale.cargo,
            "Nơi đóng": scale.place or fallback.get("Nơi đóng"),
            "Tên tàu": vessel,
            "Ngày chạy": excel_date(sail_date),
            "Người nhận": scale.recipient or fallback.get("Người nhận"),
            "VT biển": carrier,
            "MD5": _join_md5(
                _md5_cell_value(fallback),
                self._real_md5(scale),
                self._real_md5(bill) if bill else "",
            ),
            "Trạng thái nhập": _normalize_import_status(
                fallback.get("Trạng thái nhập"), "Chưa nhập"
            ),
            "Ngày cập nhật": datetime.now(),
        }

    def _bill_info_values(
        self, bill: ExtractedRow, sqt: int, fallback: Dict[str, Any]
    ) -> Dict[str, Any]:
        return {
            "SQT PM": sqt,
            "Ngày Đóng": excel_date(bill.close_date) or fallback.get("Ngày Đóng"),
            "Số Container": bill.container or fallback.get("Số Container"),
            "Biển số xe": bill.truck_no or fallback.get("Biển số xe"),
            "Số B/L": bill.bill_no or fallback.get("Số B/L"),
            "Số chì/Seal": bill.seal or fallback.get("Số chì/Seal"),
            "Số tấn": bill.tons if bill.tons is not None else fallback.get("Số tấn"),
            "Loại hàng": bill.cargo or fallback.get("Loại hàng"),
            "Nơi đóng": bill.place or fallback.get("Nơi đóng"),
            "Tên tàu": bill.vessel or fallback.get("Tên tàu"),
            "Ngày chạy": excel_date(bill.sail_date) or fallback.get("Ngày chạy"),
            "Người nhận": bill.recipient or fallback.get("Người nhận"),
            "VT biển": bill.carrier or fallback.get("VT biển"),
            "MD5": _join_md5(_md5_cell_value(fallback), self._real_md5(bill)),
            "Trạng thái nhập": _normalize_import_status(
                fallback.get("Trạng thái nhập"), "Chưa nhập"
            ),
            "Ngày cập nhật": datetime.now(),
        }

    @staticmethod
    def _find_conflicts(
        target: Optional[_TargetInfoRow], values: Dict[str, Any]
    ) -> List[FieldConflict]:
        if target is None:
            return []
        ignored = {
            "MD5",
            "Trạng thái nhập",
            "Ngày nhập cuối",
            "Ngày cập nhật",
        }
        conflicts = []
        for name, new_value in values.items():
            current = target.values.get(name)
            if name in ignored or current in (None, "") or new_value in (None, ""):
                continue
            both_numbers = all(
                isinstance(value, (int, float)) and not isinstance(value, bool)
                for value in (current, new_value)
            )
            if both_numbers:
                current_compare = float(current)
                new_compare = float(new_value)
            else:
                current_compare = parse_date(current) or _key_text(current)
                new_compare = parse_date(new_value) or _key_text(new_value)
            if current_compare != new_compare:
                raw_id = (
                    f"{target.sheet_name}|{target.row_number}|"
                    f"{name}|{current}|{new_value}"
                )
                conflicts.append(
                    FieldConflict(
                        conflict_id=hashlib.sha1(raw_id.encode("utf-8")).hexdigest(),
                        target_row=target.row_number,
                        field_name=name,
                        current_value=current,
                        new_value=new_value,
                        target_sheet=target.sheet_name,
                    )
                )
        return conflicts

    def _analyze_expenses(
        self,
        expenses: Sequence[ExtractedRow],
        virtual_rows: Sequence[_TargetInfoRow],
        snapshot: _WorkbookSnapshot,
        analysis: ImportAnalysis,
        requested_resolutions: Dict[str, IssueResolution],
    ) -> None:
        for expense in expenses:
            matches = [
                row
                for row in virtual_rows
                if row.close_date == expense.close_date
                and row.container == expense.container
            ]
            sqt_values = {row.sqt for row in matches if row.sqt}
            if len(sqt_values) != 1:
                many = bool(sqt_values)
                issue = self._row_issue(
                    ISSUE_MULTIPLE_TARGETS if many else ISSUE_NO_TARGET,
                    "Khớp nhiều dòng quyết toán" if many else "Không tìm thấy SQT",
                    REASON_MANY_TARGETS if many else REASON_NO_TARGET,
                    expense,
                    candidates=(
                        [self._target_candidate(row) for row in matches if row.sqt]
                        if many
                        else []
                    ),
                )
                resolution = self._register_issue(
                    analysis, issue, requested_resolutions
                )
                if resolution is None:
                    self._mark_unmatched(
                        analysis,
                        expense,
                        REASON_MANY_TARGETS if many else REASON_NO_TARGET,
                    )
                    continue
                if resolution.action == RESOLUTION_SKIP:
                    continue
                selected_candidate = next(
                    item
                    for item in issue.candidates
                    if item.candidate_id == resolution.candidate_id
                )
                matches = [
                    row
                    for row in matches
                    if row.sheet_name == selected_candidate.sheet_name
                    and row.row_number == selected_candidate.row_number
                    and row.sqt == selected_candidate.sqt
                ]
                sqt_values = {selected_candidate.sqt}
            sqt = next(iter(sqt_values))
            target = next((row for row in matches if row.sqt == sqt), None)
            target_values = target.values if target else {}
            check_status = (
                "Cần kiểm tra" if self._source_needs_review(expense) else "OK"
            )
            if (
                expense.total is not None
                and expense.amount is not None
                and expense.vat is not None
                and abs(float(expense.total) - float(expense.amount) - float(expense.vat)) > 1
            ):
                check_status = "Cần kiểm tra"
            existing = next(
                (
                    (row_number, values)
                    for row_number, values in snapshot.expense_rows
                    if parse_date(values.get("Ngày tháng")) == expense.close_date
                    and normalize_container(values.get("Số Container")) == expense.container
                ),
                None,
            )
            values = {
                "SQT PM": sqt,
                "Ngày tháng": excel_date(expense.close_date),
                "Số Container": expense.container,
                "Nơi đóng": target_values.get("Nơi đóng"),
                "Số HĐ": expense.invoice_no,
                "Giá vật liệu": expense.material_price,
                "Đơn giá": expense.unit_price,
                "Thành tiền": expense.amount,
                "VAT": expense.vat,
                "Tổng tiền": expense.total,
                "MD5": _join_md5(
                    _md5_cell_value(existing[1]) if existing else "",
                    self._real_md5(expense),
                ),
                "Trạng thái nhập": _normalize_import_status(
                    existing[1].get("Trạng thái nhập") if existing else None,
                    "Chưa nhập",
                ),
                "Ngày cập nhật": datetime.now(),
            }
            conflicts: List[FieldConflict] = []
            if existing:
                expense_target = _TargetInfoRow(
                    sheet_name=EXPENSE_SHEET,
                    row_number=existing[0],
                    sqt=sqt,
                    close_date=expense.close_date,
                    container=expense.container,
                    cargo="",
                    values=existing[1],
                )
                conflicts = self._find_conflicts(expense_target, values)
                analysis.conflicts.extend(conflicts)
                if conflicts and not self._apply_conflict_issue(
                    analysis,
                    requested_resolutions,
                    expense,
                    [],
                    conflicts,
                    values,
                ):
                    continue
            analysis.expense_changes.append(
                ExpenseChange(
                    action="UPDATE" if existing else "CREATE",
                    target_row=existing[0] if existing else None,
                    values=values,
                    staged_ids=[expense.staged_id],
                    document_md5s=[self._real_md5(expense)],
                    check_status=check_status,
                    conflicts=conflicts,
                )
            )
            analysis.pending_states[expense.staged_id] = (STATE_READY, "")
            analysis.completed_staged_ids.append(expense.staged_id)

    # ------------------------------------------------------------------ #
    # Workbook helpers
    # ------------------------------------------------------------------ #
    @staticmethod
    def _info_sheet_names(wb) -> List[str]:
        monthly = [name for name in wb.sheetnames if is_month_info_sheet(name)]
        if monthly:
            return monthly
        return [INFO_SHEET] if INFO_SHEET in wb.sheetnames else []

    @staticmethod
    def _header_map(ws) -> Dict[str, int]:
        mapping: Dict[str, int] = {}
        for col in range(1, ws.max_column + 1):
            value = ws.cell(1, col).value
            if value in (None, ""):
                continue
            raw = re.sub(r"\s+", " ", str(value)).strip()
            canonical = _canonical_header(raw)
            mapping[raw] = col
            mapping[canonical] = col
        if "MD5" not in mapping and "Mã MD5" in mapping:
            mapping["MD5"] = mapping["Mã MD5"]
        if "Mã MD5" not in mapping and "MD5" in mapping:
            mapping["Mã MD5"] = mapping["MD5"]
        return mapping

    def _validate_daily_headers(
        self, ws, expected: Sequence[str], *, info_sheet: bool
    ) -> Dict[str, int]:
        mapping = self._header_map(ws)
        auto_addable = {"Ngày nhập cuối", "Ngày cập nhật", "MD5", "Trạng thái nhập"}
        if info_sheet:
            auto_addable.update({"Biển số xe", "Số B/L", "Lô SX", "Dự kiến giao"})
        else:
            auto_addable.add("Nơi đóng")
        missing = [
            name
            for name in expected
            if name not in mapping and name not in auto_addable
        ]
        if missing:
            raise DailyImportError(
                f"Sheet '{ws.title}' thiếu các cột: " + ", ".join(missing)
            )
        return mapping

    def _upgrade_workbook(self, wb) -> None:
        info_sheet_names = self._info_sheet_names(wb)
        if not info_sheet_names:
            raise DailyImportError(
                f"File theo dõi thiếu sheet thông tin '{INFO_SHEET}' hoặc các sheet 'Tháng N'."
            )
        for sheet_name in info_sheet_names:
            self._upgrade_info_sheet(wb[sheet_name])

        if EXPENSE_SHEET not in wb.sheetnames:
            expense_ws = wb.create_sheet(EXPENSE_SHEET)
            expense_ws.append(EXPENSE_HEADERS + ["Ngày nhập cuối"])
        self._upgrade_expense_sheet(wb[EXPENSE_SHEET])

    def _upgrade_info_sheet(self, ws) -> None:
        self._canonicalize_headers(ws)
        self._validate_daily_headers(ws, INFO_HEADERS, info_sheet=True)
        if "Biển số xe" not in self._header_map(ws):
            if not self._use_blank_header(ws, "Biển số xe", before_header="Số chì/Seal"):
                self._ensure_column(ws, "Biển số xe", before_header="Số chì/Seal")
        if "Số B/L" not in self._header_map(ws):
            if not self._use_blank_header(ws, "Số B/L", before_header="Số chì/Seal"):
                if not self._use_empty_placeholder_header(
                    ws,
                    "Column1",
                    "Số B/L",
                    before_header="Số chì/Seal",
                ):
                    self._ensure_column(ws, "Số B/L", before_header="Số chì/Seal")
        if "Lô SX" not in self._header_map(ws):
            self._ensure_column(ws, "Lô SX", before_header="Số tấn")
        if "Dự kiến giao" not in self._header_map(ws):
            if not self._use_blank_header(ws, "Dự kiến giao", before_header="Người nhận"):
                if not self._use_blank_header(ws, "Dự kiến giao", after_header="Ngày chạy"):
                    self._ensure_column(ws, "Dự kiến giao", before_header="Người nhận")
        self._ensure_column(ws, "Trạng thái nhập")
        self._ensure_column(ws, "Ngày cập nhật", before_header="MD5")
        self._ensure_column(ws, "Ngày nhập cuối", before_header="Ngày cập nhật")
        if "MD5" not in self._header_map(ws):
            self._ensure_column(ws, "MD5")
        self._validate_upgraded_headers(ws, INFO_HEADERS)
        if is_month_info_sheet(ws.title):
            self._hide_columns(ws, INFO_TECHNICAL_HEADERS)

    def _upgrade_expense_sheet(self, ws) -> None:
        self._canonicalize_headers(ws)
        self._validate_daily_headers(ws, EXPENSE_HEADERS, info_sheet=False)
        self._ensure_column(ws, "Nơi đóng", before_header="Số HĐ")
        self._ensure_column(ws, "Trạng thái nhập")
        self._ensure_column(ws, "Ngày cập nhật", before_header="MD5")
        if "MD5" not in self._header_map(ws):
            self._ensure_column(ws, "MD5")
        self._validate_upgraded_headers(ws, EXPENSE_HEADERS)

    def _validate_upgraded_headers(self, ws, expected: Sequence[str]) -> None:
        mapping = self._header_map(ws)
        missing_after = [name for name in expected if name not in mapping]
        if missing_after:
            raise DailyImportError(
                f"Không nâng cấp được sheet '{ws.title}': "
                + ", ".join(missing_after)
            )

    def _sync_expense_places(self, wb, expense_ws) -> None:
        expense_map = self._header_map(expense_ws)
        if not all(name in expense_map for name in ("SQT PM", "Nơi đóng")):
            return

        place_by_sqt: Dict[str, Any] = {}
        duplicate_sqt: set[str] = set()
        for sheet_name in self._info_sheet_names(wb):
            info_ws = wb[sheet_name]
            info_map = self._header_map(info_ws)
            if not all(name in info_map for name in ("SQT PM", "Nơi đóng")):
                continue
            for row_number in range(2, self._last_data_row(info_ws) + 1):
                sqt = _sqt_key(info_ws.cell(row_number, info_map["SQT PM"]).value)
                if not sqt:
                    continue
                place = info_ws.cell(row_number, info_map["Nơi đóng"]).value
                if sqt in place_by_sqt and place_by_sqt[sqt] != place:
                    duplicate_sqt.add(sqt)
                    continue
                place_by_sqt[sqt] = place

        for sqt in duplicate_sqt:
            place_by_sqt.pop(sqt, None)

        for row_number in range(2, self._last_data_row(expense_ws) + 1):
            sqt = _sqt_key(expense_ws.cell(row_number, expense_map["SQT PM"]).value)
            if not sqt or sqt not in place_by_sqt:
                continue
            expense_ws.cell(row_number, expense_map["Nơi đóng"]).value = place_by_sqt[sqt]

    def _canonicalize_headers(self, ws) -> None:
        for col in range(1, ws.max_column + 1):
            value = ws.cell(1, col).value
            if value in (None, ""):
                continue
            canonical = _canonical_header(value)
            if canonical and canonical != str(value).strip():
                ws.cell(1, col).value = canonical

    def _use_blank_header(
        self,
        ws,
        header: str,
        *,
        before_header: Optional[str] = None,
        after_header: Optional[str] = None,
    ) -> bool:
        mapping = self._header_map(ws)
        candidates: List[int] = []
        if before_header and before_header in mapping:
            candidates.append(mapping[before_header] - 1)
        if after_header and after_header in mapping:
            candidates.append(mapping[after_header] + 1)
        for col in candidates:
            if col < 1 or col > ws.max_column:
                continue
            if ws.cell(1, col).value in (None, ""):
                ws.cell(1, col).value = header
                return True
        return False

    def _use_empty_placeholder_header(
        self,
        ws,
        placeholder: str,
        header: str,
        *,
        before_header: str,
    ) -> bool:
        """Đổi tên cột giữ chỗ chỉ khi cột đó hoàn toàn chưa có dữ liệu."""
        mapping = self._header_map(ws)
        if before_header not in mapping:
            return False
        col = mapping[before_header] - 1
        if col < 1 or _key_text(ws.cell(1, col).value) != _key_text(placeholder):
            return False
        if any(
            ws.cell(row, col).value not in (None, "")
            for row in range(2, ws.max_row + 1)
        ):
            return False
        ws.cell(1, col).value = header
        return True

    @staticmethod
    def _hide_columns(ws, headers: Iterable[str]) -> None:
        mapping = DailyImportService._header_map(ws)
        for header in headers:
            if header not in mapping:
                continue
            ws.column_dimensions[get_column_letter(mapping[header])].hidden = True

    def _ensure_month_sheet(self, wb, sheet_name: str) -> None:
        if sheet_name in wb.sheetnames:
            return
        target_month = month_sheet_number(sheet_name)
        if target_month is None:
            raise DailyImportError(f"Tên sheet tháng không hợp lệ: {sheet_name}")
        existing = [
            (month_sheet_number(name) or 0, name)
            for name in wb.sheetnames
            if is_month_info_sheet(name)
        ]
        if not existing:
            raise DailyImportError("Không có sheet tháng nào để làm mẫu tạo tháng mới.")
        _, source_name = min(
            existing,
            key=lambda item: (abs(item[0] - target_month), item[0] > target_month),
        )
        ws = wb.copy_worksheet(wb[source_name])
        ws.title = sheet_name
        if ws.max_row > 1:
            ws.delete_rows(2, ws.max_row - 1)
        self._upgrade_info_sheet(ws)

    @staticmethod
    def _last_data_row(ws) -> int:
        for row_number in range(ws.max_row, 1, -1):
            if any(
                ws.cell(row_number, col).value not in (None, "")
                for col in range(1, ws.max_column + 1)
            ):
                return row_number
        return 1

    def _ensure_column(
        self,
        ws,
        header: str,
        *,
        before_header: Optional[str] = None,
    ) -> None:
        mapping = self._header_map(ws)
        if header in mapping:
            return
        insert_col = ws.max_column + 1
        if before_header and before_header in mapping:
            insert_col = mapping[before_header]
            ws.insert_cols(insert_col, 1)
        ws.cell(1, insert_col).value = header
        style_source_col = min(insert_col + 1, ws.max_column)
        if style_source_col != insert_col:
            ws.cell(1, insert_col)._style = copy(ws.cell(1, style_source_col)._style)
        for row_number in range(2, ws.max_row + 1):
            if style_source_col != insert_col:
                ws.cell(row_number, insert_col)._style = copy(
                    ws.cell(row_number, style_source_col)._style
                )

    @staticmethod
    def _copy_row_style(ws, source_row: int, target_row: int) -> None:
        if source_row < 2 or source_row > ws.max_row:
            return
        for col in range(1, ws.max_column + 1):
            source = ws.cell(source_row, col)
            target = ws.cell(target_row, col)
            if source.has_style:
                target._style = copy(source._style)
            target.number_format = source.number_format
            target.alignment = copy(source.alignment)

    def _format_workbook(self, wb) -> None:
        sheet_names = self._info_sheet_names(wb)
        if EXPENSE_SHEET in wb.sheetnames:
            sheet_names.append(EXPENSE_SHEET)
        for sheet_name in sheet_names:
            ws = wb[sheet_name]
            mapping = self._header_map(ws)
            last_row = max(1, self._last_data_row(ws))
            ws.freeze_panes = "A2"
            ws.auto_filter.ref = f"A1:{get_column_letter(ws.max_column)}{last_row}"
            # Form mới chỉ còn hai trạng thái nhập phục vụ PAD.
            ws.data_validations.dataValidation = []
            import_col = get_column_letter(mapping["Trạng thái nhập"])
            import_validation = DataValidation(
                type="list", formula1='"Chưa nhập,Đã nhập"', allow_blank=False
            )
            ws.add_data_validation(import_validation)
            import_validation.add(f"{import_col}2:{import_col}10000")
            status_col = mapping["Trạng thái nhập"]
            for row_number in range(2, last_row + 1):
                if any(
                    ws.cell(row_number, col).value not in (None, "")
                    for col in range(1, ws.max_column + 1)
                    if col != status_col
                ):
                    ws.cell(row_number, status_col).value = _normalize_import_status(
                        ws.cell(row_number, status_col).value,
                        "Chưa nhập",
                    )
            for header in ("Ngày Đóng", "Ngày chạy", "Ngày Giao", "Ngày tháng"):
                if header in mapping:
                    for row_number in range(2, last_row + 1):
                        ws.cell(row_number, mapping[header]).number_format = "dd/mm/yyyy"
            for header in ("Ngày nhập cuối", "Ngày cập nhật"):
                if header not in mapping:
                    continue
                for row_number in range(2, last_row + 1):
                    ws.cell(row_number, mapping[header]).number_format = (
                        "dd/mm/yyyy hh:mm:ss"
                    )
            for header in ("Giá vật liệu", "Đơn giá", "Thành tiền", "VAT", "Tổng tiền"):
                if header in mapping:
                    for row_number in range(2, last_row + 1):
                        ws.cell(row_number, mapping[header]).number_format = "#,##0"
            if "Số tấn" in mapping:
                for row_number in range(2, last_row + 1):
                    ws.cell(row_number, mapping["Số tấn"]).number_format = "0.00"
            if is_month_info_sheet(sheet_name):
                self._hide_columns(ws, INFO_TECHNICAL_HEADERS)
            # Cập nhật bảng Excel nếu workbook cũ có Table.
            for table in ws.tables.values():
                table.ref = f"A1:{get_column_letter(ws.max_column)}{max(2, last_row)}"

    def _validate_saved_workbook(self, path: str) -> None:
        wb = load_workbook(path, read_only=True, data_only=False)
        try:
            info_sheet_names = self._info_sheet_names(wb)
            if not info_sheet_names:
                raise DailyImportError("File tạm thiếu sheet thông tin.")
            for sheet_name in info_sheet_names:
                mapping = self._header_map(wb[sheet_name])
                missing = [name for name in INFO_HEADERS if name not in mapping]
                if missing:
                    raise DailyImportError(
                        f"File tạm thiếu cột ở sheet '{sheet_name}': "
                        + ", ".join(missing)
                    )
            if EXPENSE_SHEET not in wb.sheetnames:
                raise DailyImportError(f"File tạm thiếu sheet '{EXPENSE_SHEET}'.")
            mapping = self._header_map(wb[EXPENSE_SHEET])
            missing = [name for name in EXPENSE_HEADERS if name not in mapping]
            if missing:
                raise DailyImportError(
                    f"File tạm thiếu cột ở sheet '{EXPENSE_SHEET}': "
                    + ", ".join(missing)
                )
        finally:
            wb.close()

    @staticmethod
    def _summary_note(summary: ImportSummary) -> str:
        return (
            f"Đã nhập file theo dõi: {summary.new_info} dòng quyết toán mới, "
            f"{summary.updated_info} dòng cập nhật, {summary.new_expenses} khoản chi mới; "
            f"còn {summary.pending_count} dữ liệu chờ."
        )
