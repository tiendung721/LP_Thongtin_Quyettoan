import logging
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

from app import logger_setup


class LoggerStartupTests(unittest.TestCase):
    def test_restarting_pm_removes_previous_log_files(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            logs_dir = Path(tmp) / "Logs"
            logs_dir.mkdir()
            for name in ("app_20261019.log", "pad_launcher.log", "rpa_excel_helper.log"):
                (logs_dir / name).write_text("old session\n", encoding="utf-8")
            (logs_dir / "keep.txt").write_text("keep", encoding="utf-8")

            with patch.object(logger_setup, "LOGGER_NAME", "QuyetToan.startup_test"):
                logger = logging.getLogger(logger_setup.LOGGER_NAME)
                try:
                    logger_setup.setup_logger(str(logs_dir))
                    self.assertEqual(sorted(path.name for path in logs_dir.glob("*.log")),
                                     [f"app_{logger_setup.datetime.now():%Y%m%d}.log"])
                    logger.info("first session")

                    for handler in logger.handlers[:]:
                        logger.removeHandler(handler)
                        handler.close()

                    logger_setup.setup_logger(str(logs_dir))
                    logger.info("second session")
                    current_log = next(logs_dir.glob("app_*.log"))
                    content = current_log.read_text(encoding="utf-8")
                    self.assertIn("second session", content)
                    self.assertNotIn("first session", content)
                    self.assertEqual((logs_dir / "keep.txt").read_text(encoding="utf-8"), "keep")
                finally:
                    for handler in logger.handlers[:]:
                        logger.removeHandler(handler)
                        handler.close()


if __name__ == "__main__":
    unittest.main()
