"""Conservative inspection and manual recovery of Excel owner files on Windows/SMB."""

from __future__ import annotations

import ctypes
import os
import sys
import time
import uuid
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path


AVAILABLE = "AVAILABLE"
IN_USE = "IN_USE"
POSSIBLE_STALE = "POSSIBLE_STALE"
UNKNOWN = "UNKNOWN"
MISSING = "MISSING"


@dataclass(frozen=True)
class ExcelLockState:
    status: str
    lock_path: str = ""
    reason: str = ""


def owner_file_candidates(excel_path: str) -> tuple[str, ...]:
    """Office versions can use either observed owner-file naming convention."""
    path = Path(excel_path)
    names = ("~$" + path.name, "~$" + path.name[2:])
    return tuple(str(path.with_name(name)) for name in dict.fromkeys(names))


def _exclusive_probe(path: str) -> str:
    """Check sharing on the file itself; access and network errors stay unknown."""
    if sys.platform != "win32":
        return UNKNOWN

    kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)
    kernel32.CreateFileW.argtypes = (
        ctypes.c_wchar_p, ctypes.c_uint32, ctypes.c_uint32,
        ctypes.c_void_p, ctypes.c_uint32, ctypes.c_uint32, ctypes.c_void_p,
    )
    kernel32.CreateFileW.restype = ctypes.c_void_p
    kernel32.CloseHandle.argtypes = (ctypes.c_void_p,)
    kernel32.CloseHandle.restype = ctypes.c_int
    handle = kernel32.CreateFileW(path, 0x80000000 | 0x40000000, 0, None, 3, 0, None)
    if handle == ctypes.c_void_p(-1).value:
        error = ctypes.get_last_error()
        if error in (32, 33):  # ERROR_SHARING_VIOLATION / ERROR_LOCK_VIOLATION
            return IN_USE
        if error in (2, 3):
            return MISSING
        return UNKNOWN
    kernel32.CloseHandle(handle)
    return AVAILABLE


def _owner_name_is_ambiguous(excel_path: str, lock_path: str) -> bool:
    """Reject a ~$ name that could also belong to a different workbook."""
    target = os.path.normcase(os.path.abspath(excel_path))
    owner = os.path.normcase(os.path.abspath(lock_path))
    with os.scandir(Path(excel_path).parent) as entries:
        for entry in entries:
            if entry.name.startswith("~$"):
                continue
            if Path(entry.name).suffix.lower() not in {".xlsx", ".xlsm", ".xlsb", ".xls"}:
                continue
            other = os.path.normcase(os.path.abspath(entry.path))
            if other == target:
                continue
            if owner in (
                os.path.normcase(os.path.abspath(candidate))
                for candidate in owner_file_candidates(entry.path)
            ):
                return True
    return False


def inspect_excel_lock(excel_path: str) -> ExcelLockState:
    if not excel_path:
        return ExcelLockState(MISSING, reason="Chưa cấu hình file Excel.")
    try:
        os.stat(excel_path)
    except FileNotFoundError:
        return ExcelLockState(MISSING, reason="Không tìm thấy file Excel.")
    except OSError:
        return ExcelLockState(UNKNOWN, reason="Không truy cập được file Excel hoặc mạng LAN.")

    present = []
    for candidate in owner_file_candidates(excel_path):
        try:
            os.stat(candidate)
            present.append(candidate)
        except FileNotFoundError:
            continue
        except OSError:
            return ExcelLockState(UNKNOWN, reason="Không kiểm tra được file khóa trên LAN.")
    if len(present) > 1:
        return ExcelLockState(UNKNOWN, reason="Có nhiều file khóa có thể thuộc workbook này.")

    lock_path = present[0] if present else ""
    if lock_path:
        try:
            if _owner_name_is_ambiguous(excel_path, lock_path):
                return ExcelLockState(
                    UNKNOWN, lock_path, "Tên file khóa cũng khớp với workbook khác trong thư mục."
                )
        except OSError:
            return ExcelLockState(UNKNOWN, lock_path, "Không kiểm tra được các file trong thư mục LAN.")
    workbook_probe = _exclusive_probe(excel_path)
    if workbook_probe == IN_USE:
        return ExcelLockState(IN_USE, lock_path, "File Excel đang được sử dụng.")
    if workbook_probe != AVAILABLE:
        return ExcelLockState(UNKNOWN, lock_path, "Không xác định được trạng thái file Excel.")
    if not lock_path:
        return ExcelLockState(AVAILABLE)

    owner_probe = _exclusive_probe(lock_path)
    if owner_probe == IN_USE:
        return ExcelLockState(IN_USE, lock_path, "File khóa đang được sử dụng.")
    if owner_probe != AVAILABLE:
        return ExcelLockState(UNKNOWN, lock_path, "Không xác định được trạng thái file khóa.")
    return ExcelLockState(POSSIBLE_STALE, lock_path, "File khóa có thể đã bị bỏ sót.")


def _identity(path: str) -> tuple[int, int, int]:
    stat = os.stat(path)
    return stat.st_size, stat.st_mtime_ns, stat.st_ino


def recover_stale_lock(excel_path: str, wait_seconds: float = 2.0) -> str:
    """Rename one owner file after a user action and two matching inspections.

    This is a best-effort check, not an atomic guarantee against another LAN user
    opening Excel after the check. The caller must ask for confirmation first.
    """
    first = inspect_excel_lock(excel_path)
    if first.status != POSSIBLE_STALE:
        raise RuntimeError(first.reason or "Không có file khóa đủ điều kiện khôi phục.")
    try:
        identity = (_identity(excel_path), _identity(first.lock_path))
        time.sleep(max(0.0, wait_seconds))
        second = inspect_excel_lock(excel_path)
        if second.status != POSSIBLE_STALE or second.lock_path != first.lock_path:
            raise RuntimeError("Trạng thái file khóa đã thay đổi; chưa đổi tên file.")
        if identity != (_identity(excel_path), _identity(first.lock_path)):
            raise RuntimeError("File Excel hoặc file khóa vừa thay đổi; chưa đổi tên file.")
        stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        backup = str(Path(first.lock_path).with_name(
            f"lock_backup_{stamp}_{uuid.uuid4().hex[:8]}.bak"
        ))
        os.rename(first.lock_path, backup)
    except OSError as exc:
        raise RuntimeError(f"Không đổi tên được file khóa trên LAN: {exc}") from exc

    final = inspect_excel_lock(excel_path)
    if final.status != AVAILABLE:
        raise RuntimeError(
            "Đã lưu giữ file khóa cũ nhưng trạng thái Excel lại thay đổi. "
            f"File lưu giữ: {backup}. Vui lòng kiểm tra trước khi tiếp tục."
        )
    return backup
