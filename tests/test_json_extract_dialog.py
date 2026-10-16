from __future__ import annotations

import json
import os
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

from PySide6.QtWidgets import QApplication

from app.daily_import_ui import JsonExtractDialog


class JsonExtractDialogTests(unittest.TestCase):
    def setUp(self) -> None:
        self.app = QApplication.instance() or QApplication([])
        self.temp = tempfile.TemporaryDirectory()
        self.path = Path(self.temp.name) / "extract.json"
        self.path.write_text(json.dumps({
            "metadata": {},
            "du_lieu_boc_tach": [
                {"stt_hien_thi": 3, "loai_chung_tu": "Khoản chi", "file_nguon": "third.jpg", "so_container": "CCC", "ngay_dong": "03/05/2026", "tong_tien": 300},
                {"stt_hien_thi": 1, "loai_chung_tu": "Khoản chi", "file_nguon": "first.jpg", "so_container": "AAA", "ngay_dong": "01/05/2026", "tong_tien": 100, "canh_bao": ["Cần xem lại"]},
                {"stt_hien_thi": 2, "loai_chung_tu": "Khoản chi", "file_nguon": "second.jpg", "so_container": "BBB", "ngay_dong": "02/05/2026", "tong_tien": 200},
            ],
        }, ensure_ascii=False), encoding="utf-8")

    def tearDown(self) -> None:
        self.temp.cleanup()

    def test_sort_and_filter_keep_actions_on_original_data_row(self) -> None:
        dialog = JsonExtractDialog(str(self.path))
        self.assertEqual(dialog.type_filter.itemText(1), "Khoản chi")
        dialog.type_filter.setCurrentText("Bill")
        self.assertEqual(dialog._visible_indices, [])
        dialog.type_filter.setCurrentText("Mọi chứng từ")
        dialog._sort_by(0)
        self.assertEqual(dialog._visible_indices, [0, 2, 1])
        dialog.status_filter.setCurrentText("Có cảnh báo")
        self.assertEqual(dialog._visible_indices, [1])
        self.assertEqual(dialog._selected_index(), 1)
        with patch("app.daily_import_ui.notify.confirm", return_value=True):
            dialog.delete_row()
        self.assertEqual([row["file_nguon"] for row in dialog.rows], ["third.jpg", "second.jpg"])


if __name__ == "__main__":
    unittest.main()
