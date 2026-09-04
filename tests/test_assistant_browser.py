from __future__ import annotations

import logging
import os
import subprocess
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

from app.services.assistant_browser import (
    AssistantBrowserError,
    close_script_for_bat,
    output_is_from_session,
    request_assistant_browser_close,
)


class AssistantBrowserTests(unittest.TestCase):
    def setUp(self) -> None:
        self.temp = tempfile.TemporaryDirectory()
        self.root = Path(self.temp.name)
        self.launcher = self.root / "launcher with spaces"
        self.launcher.mkdir()
        self.bat = self.launcher / "Mo_Tro_Ly_RPA.bat"
        self.bat.write_text("@echo off\n", encoding="utf-8")
        self.close_script = self.launcher / "close_chatgpt_app.ps1"
        self.close_script.write_text("exit 0\n", encoding="utf-8")
        self.logger = logging.getLogger("assistant-browser-test")

    def tearDown(self) -> None:
        self.temp.cleanup()

    def test_close_script_is_resolved_next_to_launcher_bat(self) -> None:
        self.assertEqual(
            close_script_for_bat(str(self.bat)),
            str(self.close_script),
        )

    def test_output_must_be_newer_than_active_session(self) -> None:
        output = self.root / "boc_tach.json"
        output.write_text("{}", encoding="utf-8")
        os.utime(output, (100.0, 100.0))

        self.assertTrue(output_is_from_session(str(output), 101.0))
        self.assertFalse(output_is_from_session(str(output), 103.1))
        self.assertFalse(output_is_from_session(str(output), None))

    @patch("app.services.assistant_browser.subprocess.Popen")
    def test_close_request_runs_powershell_without_waiting(self, popen) -> None:
        result = request_assistant_browser_close(str(self.bat), logger=self.logger)

        self.assertIs(result, popen.return_value)
        args, kwargs = popen.call_args
        self.assertEqual(
            args[0],
            [
                "powershell.exe",
                "-NoProfile",
                "-ExecutionPolicy",
                "Bypass",
                "-File",
                str(self.close_script),
            ],
        )
        self.assertEqual(kwargs["cwd"], str(self.launcher))
        self.assertEqual(kwargs["stdin"], subprocess.DEVNULL)
        self.assertEqual(kwargs["stdout"], subprocess.DEVNULL)
        self.assertEqual(kwargs["stderr"], subprocess.DEVNULL)

    def test_missing_close_script_is_reported(self) -> None:
        self.close_script.unlink()

        with self.assertRaises(AssistantBrowserError):
            request_assistant_browser_close(str(self.bat), logger=self.logger)


if __name__ == "__main__":
    unittest.main()
