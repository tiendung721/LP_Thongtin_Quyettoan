import json
import os
import tempfile
import unittest
from pathlib import Path

from app.config import AppConfig, DAILY_TRACKING_FILENAME, rebase_moved_install


def _same_path(left: str, right: str) -> bool:
    return os.path.normcase(os.path.abspath(left)) == os.path.normcase(os.path.abspath(right))


class AppConfigSingleOutputTests(unittest.TestCase):
    def test_rebase_moved_install_uses_customer_local_database_root(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            actual = Path(tmp) / "customer_app"
            config_path = actual / "Config" / "settings.json"
            old = Path(tmp) / "old_machine_app"
            data = {
                "app_root": str(old),
                "output_folder": str(old / "output"),
                "pad_input_expense_bat_path": str(old / "Launcher" / "run_input_expense.bat"),
                "daily_tracking_file": str(Path(tmp) / "old_user" / "daily.xlsx"),
            }
            result = rebase_moved_install(data, str(config_path))
            self.assertTrue(_same_path(result["app_root"], str(actual)))
            self.assertTrue(_same_path(result["output_folder"], str(actual / "output")))
            self.assertTrue(_same_path(
                result["pad_input_expense_bat_path"],
                str(actual / "Launcher" / "run_input_expense.bat"),
            ))
            self.assertTrue(_same_path(
                result["daily_tracking_file"],
                str(actual / "output" / DAILY_TRACKING_FILENAME),
            ))

    def test_load_migrates_legacy_split_folders_to_single_output(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            config_path = root / "Config" / "settings.json"
            config_path.parent.mkdir()
            config_path.write_text(
                json.dumps(
                    {
                        "app_root": str(root),
                        "download_folder": str(root / "Downloads"),
                        "output_folder": str(root / "Outputs"),
                        "daily_tracking_file": str(root / "Daily" / DAILY_TRACKING_FILENAME),
                    },
                    ensure_ascii=False,
                ),
                encoding="utf-8",
            )

            config = AppConfig.load(str(config_path))

            expected_output = str(root / "output")
            self.assertTrue(_same_path(config.output_folder, expected_output))
            self.assertTrue(
                _same_path(config.daily_tracking_file, str(root / "output" / DAILY_TRACKING_FILENAME))
            )

            saved = json.loads(config_path.read_text(encoding="utf-8"))
            self.assertNotIn("download_folder", saved)
            self.assertTrue(_same_path(saved["output_folder"], expected_output))

    def test_ensure_folders_creates_only_single_output_data_folder(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            config = AppConfig(
                {
                    "app_root": str(root),
                    "download_folder": str(root / "Downloads"),
                    "output_folder": str(root / "Outputs"),
                    "daily_tracking_file": str(root / "Daily" / DAILY_TRACKING_FILENAME),
                },
                str(root / "Config" / "settings.json"),
            )

            config.ensure_folders()

            self.assertTrue((root / "output").is_dir())
            self.assertFalse((root / "Downloads").exists())
            self.assertFalse((root / "Outputs").exists())
            self.assertFalse((root / "Daily").exists())

    def test_set_output_folder_keeps_user_selected_daily_filename(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            config = AppConfig(
                {
                    "app_root": str(root),
                    "output_folder": str(root / "output"),
                    "daily_tracking_file": str(root / "output" / DAILY_TRACKING_FILENAME),
                },
                str(root / "Config" / "settings.json"),
            )

            selected_daily = root / "output" / "Hàng ngày 2026.xlsx"
            config.set_output_folder(
                str(root / "output"),
                daily_tracking_file=str(selected_daily),
            )

            self.assertTrue(
                _same_path(config.daily_tracking_file, str(selected_daily))
            )

    def test_set_output_folder_preserves_user_selected_daily_path_outside_output(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp) / "app"
            external_daily = Path(tmp) / "Downloads" / "Hàng ngày 2026.xlsx"
            config = AppConfig(
                {
                    "app_root": str(root),
                    "output_folder": str(root / "output"),
                    "daily_tracking_file": str(root / "output" / DAILY_TRACKING_FILENAME),
                },
                str(root / "Config" / "settings.json"),
            )

            config.set_output_folder(
                str(root / "output"),
                daily_tracking_file=str(external_daily),
            )

            self.assertTrue(
                _same_path(config.daily_tracking_file, str(external_daily))
            )

    def test_load_preserves_daily_path_outside_app_root(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp) / "app"
            external_daily = Path(tmp) / "Downloads" / "Hàng ngày 2026.xlsx"
            config_path = root / "Config" / "settings.json"
            config_path.parent.mkdir(parents=True)
            config_path.write_text(
                json.dumps(
                    {
                        "app_root": str(root),
                        "output_folder": str(root / "output"),
                        "daily_tracking_file": str(external_daily),
                    },
                    ensure_ascii=False,
                ),
                encoding="utf-8",
            )

            config = AppConfig.load(str(config_path))

            self.assertTrue(
                _same_path(config.daily_tracking_file, str(external_daily))
            )


if __name__ == "__main__":
    unittest.main()
