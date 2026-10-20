from __future__ import annotations

import tempfile
import unittest
import ctypes
import sys
from pathlib import Path
from unittest.mock import patch

from app import excel_lock
from app import file_utils


class ExcelLockTests(unittest.TestCase):
    def setUp(self) -> None:
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.book = Path(self.temp.name) / "Hang ngay PM nhap lieu.xlsx"
        self.book.write_bytes(b"workbook")

    def test_both_exact_owner_names_are_recognized(self) -> None:
        candidates = excel_lock.owner_file_candidates(str(self.book))
        self.assertEqual(len(candidates), 2)
        self.assertEqual(Path(candidates[0]).name, "~$Hang ngay PM nhap lieu.xlsx")
        self.assertEqual(Path(candidates[1]).name, "~$ng ngay PM nhap lieu.xlsx")

    @patch.object(excel_lock, "_exclusive_probe", return_value=excel_lock.AVAILABLE)
    def test_multiple_candidates_are_not_recovered(self, _probe) -> None:
        for candidate in excel_lock.owner_file_candidates(str(self.book)):
            Path(candidate).write_bytes(b"owner")
        state = excel_lock.inspect_excel_lock(str(self.book))
        self.assertEqual(state.status, excel_lock.UNKNOWN)
        with self.assertRaises(RuntimeError):
            excel_lock.recover_stale_lock(str(self.book), wait_seconds=0)
        for candidate in excel_lock.owner_file_candidates(str(self.book)):
            self.assertTrue(Path(candidate).exists())

    @patch.object(excel_lock, "_exclusive_probe", return_value=excel_lock.AVAILABLE)
    def test_owner_name_matching_another_workbook_is_not_recovered(self, _probe) -> None:
        owner = Path(excel_lock.owner_file_candidates(str(self.book))[1])
        owner.write_bytes(b"owner")
        (owner.parent / owner.name[2:]).write_bytes(b"other workbook")
        self.assertEqual(excel_lock.inspect_excel_lock(str(self.book)).status, excel_lock.UNKNOWN)
        with self.assertRaises(RuntimeError):
            excel_lock.recover_stale_lock(str(self.book), wait_seconds=0)
        self.assertTrue(owner.exists())

    @patch.object(excel_lock, "_exclusive_probe", return_value=excel_lock.IN_USE)
    def test_active_workbook_is_never_renamed(self, _probe) -> None:
        owner = Path(excel_lock.owner_file_candidates(str(self.book))[0])
        owner.write_bytes(b"owner")
        self.assertEqual(excel_lock.inspect_excel_lock(str(self.book)).status, excel_lock.IN_USE)
        with self.assertRaises(RuntimeError):
            excel_lock.recover_stale_lock(str(self.book), wait_seconds=0)
        self.assertTrue(owner.exists())

    @patch.object(excel_lock, "_exclusive_probe", return_value=excel_lock.AVAILABLE)
    def test_recovery_renames_only_matching_owner_file(self, _probe) -> None:
        owner = Path(excel_lock.owner_file_candidates(str(self.book))[1])
        owner.write_bytes(b"owner")
        unrelated = Path(self.temp.name) / "~$Other.xlsx"
        unrelated.write_bytes(b"leave alone")
        backup = Path(excel_lock.recover_stale_lock(str(self.book), wait_seconds=0))
        self.assertFalse(owner.exists())
        self.assertEqual(backup.read_bytes(), b"owner")
        self.assertTrue(unrelated.exists())
        self.assertEqual(excel_lock.inspect_excel_lock(str(self.book)).status, excel_lock.AVAILABLE)

    @patch.object(excel_lock, "_exclusive_probe", return_value=excel_lock.AVAILABLE)
    def test_existing_owner_file_blocks_normal_read_until_recovery(self, _probe) -> None:
        owner = Path(excel_lock.owner_file_candidates(str(self.book))[0])
        owner.write_bytes(b"owner")
        self.assertTrue(file_utils.is_file_locked(str(self.book)))

    @unittest.skipUnless(sys.platform == "win32", "Windows sharing probe")
    def test_windows_exclusive_probe_detects_open_handle(self) -> None:
        kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)
        kernel32.CreateFileW.argtypes = (
            ctypes.c_wchar_p, ctypes.c_uint32, ctypes.c_uint32,
            ctypes.c_void_p, ctypes.c_uint32, ctypes.c_uint32, ctypes.c_void_p,
        )
        kernel32.CreateFileW.restype = ctypes.c_void_p
        handle = kernel32.CreateFileW(str(self.book), 0x80000000, 0, None, 3, 0, None)
        self.assertNotEqual(handle, ctypes.c_void_p(-1).value)
        try:
            self.assertEqual(excel_lock.inspect_excel_lock(str(self.book)).status, excel_lock.IN_USE)
        finally:
            kernel32.CloseHandle(ctypes.c_void_p(handle))

    @patch.object(excel_lock, "_exclusive_probe", return_value=excel_lock.AVAILABLE)
    def test_change_between_checks_prevents_rename(self, _probe) -> None:
        owner = Path(excel_lock.owner_file_candidates(str(self.book))[0])
        owner.write_bytes(b"owner")

        def change_file(_seconds: float) -> None:
            owner.write_bytes(b"owner updated")

        with patch.object(excel_lock.time, "sleep", side_effect=change_file):
            with self.assertRaisesRegex(RuntimeError, "vừa thay đổi"):
                excel_lock.recover_stale_lock(str(self.book))
        self.assertTrue(owner.exists())


if __name__ == "__main__":
    unittest.main()
