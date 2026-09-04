"""Quản lý vòng đời cửa sổ Chrome dành riêng cho trợ lý ChatGPT."""

from __future__ import annotations

import os
import subprocess
from typing import Optional


class AssistantBrowserError(RuntimeError):
    """Lỗi điều khiển trình duyệt trợ lý có thể ghi log cho người dùng."""


def output_is_from_session(
    path: str,
    launched_at: Optional[float],
    *,
    clock_skew_seconds: float = 2.0,
) -> bool:
    """Kiểm tra file có được tạo/sửa sau khi phiên trợ lý bắt đầu hay không."""
    if not path or launched_at is None:
        return False
    try:
        return os.path.getmtime(path) >= launched_at - clock_skew_seconds
    except OSError:
        return False


def close_script_for_bat(bat_path: str) -> str:
    """Trả về script đóng Chrome nằm cạnh BAT mở trợ lý."""
    if not bat_path:
        return ""
    return os.path.join(os.path.dirname(os.path.abspath(bat_path)), "close_chatgpt_app.ps1")


def request_assistant_browser_close(bat_path: str, *, logger) -> subprocess.Popen:
    """Gọi bất đồng bộ script đóng đúng Chrome dùng profile của trợ lý.

    Script PowerShell tự xác minh ``--user-data-dir`` trước khi đóng nên không
    tác động tới các cửa sổ Chrome thông thường của người dùng.
    """
    script_path = close_script_for_bat(bat_path)
    if not script_path or not os.path.isfile(script_path):
        raise AssistantBrowserError(
            "Không tìm thấy close_chatgpt_app.ps1 cạnh file BAT mở trợ lý."
        )

    command = [
        "powershell.exe",
        "-NoProfile",
        "-ExecutionPolicy",
        "Bypass",
        "-File",
        script_path,
    ]
    creationflags = getattr(subprocess, "CREATE_NO_WINDOW", 0)
    try:
        process = subprocess.Popen(
            command,
            cwd=os.path.dirname(script_path) or None,
            stdin=subprocess.DEVNULL,
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
            creationflags=creationflags,
        )
    except Exception as exc:  # noqa: BLE001 - chuyển thành lỗi nghiệp vụ để UI ghi log
        raise AssistantBrowserError(
            f"Không gọi được lệnh đóng trình duyệt trợ lý: {exc}"
        ) from exc

    logger.info("Đã gửi yêu cầu đóng trình duyệt trợ lý: %s", script_path)
    return process
