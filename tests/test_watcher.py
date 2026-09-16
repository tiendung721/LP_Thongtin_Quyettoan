from __future__ import annotations

import logging
import os
import tempfile
import time
import unittest
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import Mock, patch

from watchdog.events import FileDeletedEvent, FileMovedEvent

from app.watcher import OutputWatcher, _OutputEventHandler


def _watcher(folder: str) -> OutputWatcher:
    config = SimpleNamespace(
        output_folder=folder,
        allowed_extensions=[".json"],
        output_file_patterns=["boc_tach_chung_tu_container_*.json"],
        file_stable_seconds=0,
    )
    return OutputWatcher(config, Mock(), logging.getLogger("test-watcher"))


class OutputWatcherTests(unittest.TestCase):
    def test_atomic_save_does_not_turn_handled_file_into_new_output(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            destination = Path(tmp) / "boc_tach_chung_tu_container_test.json"
            temporary = Path(tmp) / f".{destination.name}.tmp-{os.getpid()}"
            destination.write_text('{"version": 1}', encoding="utf-8")
            watcher = _watcher(tmp)
            handler = _OutputEventHandler(watcher)
            watcher.mark_handled(str(destination))

            temporary.write_text('{"version": 2}', encoding="utf-8")
            destination.unlink()
            with patch("app.watcher.DELETE_GRACE_SECONDS", 0.01):
                handler.on_deleted(FileDeletedEvent(str(destination)))
                os.replace(temporary, destination)
                with patch("app.watcher.threading.Thread") as thread:
                    handler.on_moved(FileMovedEvent(str(temporary), str(destination)))
                    thread.assert_not_called()
                time.sleep(0.03)

            self.assertIn(watcher._path_key(str(destination)), watcher._handled_paths)

    def test_real_delete_is_forgotten_after_grace_period(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            destination = Path(tmp) / "boc_tach_chung_tu_container_test.json"
            destination.write_text("{}", encoding="utf-8")
            watcher = _watcher(tmp)
            handler = _OutputEventHandler(watcher)
            watcher.mark_handled(str(destination))

            destination.unlink()
            with patch("app.watcher.DELETE_GRACE_SECONDS", 0.01):
                handler.on_deleted(FileDeletedEvent(str(destination)))
                time.sleep(0.03)

            self.assertNotIn(watcher._path_key(str(destination)), watcher._handled_paths)

    def test_path_keys_ignore_relative_form_and_case_rules_of_platform(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            watcher = _watcher(tmp)
            absolute = str(Path(tmp) / "boc_tach_chung_tu_container_test.json")
            relative = os.path.relpath(absolute)

            watcher.mark_handled(relative)

            self.assertIn(watcher._path_key(absolute), watcher._handled_paths)


if __name__ == "__main__":
    unittest.main()
