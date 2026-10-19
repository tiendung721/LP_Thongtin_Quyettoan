from __future__ import annotations

import json
import os
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

from PySide6.QtWidgets import QApplication

from app.daily_import import DOC_EXPENSE
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
        self.assertEqual(dialog.summary_values["files"].text(), "3")
        self.assertEqual(dialog.summary_values["total"].text(), "3")
        self.assertEqual(dialog.summary_values["warnings"].text(), "1")
        self.assertEqual(dialog.summary_metrics["warnings"].property("tone"), "warning")
        self.assertEqual(dialog.type_values[DOC_EXPENSE].text(), "3")
        self.assertEqual(dialog.type_filter.itemText(1), "Khoản chi")
        dialog.type_filter.setCurrentText("Bill")
        self.assertEqual(dialog._visible_indices, [])
        dialog.type_filter.setCurrentText("Mọi chứng từ")
        dialog._sort_by(0)
        self.assertEqual(dialog._visible_indices, [0, 2, 1])
        self.assertNotIn("Cảnh báo", dialog.COLUMNS)
        self.assertEqual(dialog.status_filter.findText("Có cảnh báo"), -1)
        dialog.status_filter.setCurrentText("Cần kiểm tra")
        self.assertEqual(dialog._visible_indices, [1])
        self.assertEqual(dialog._selected_index(), 1)
        with patch("app.daily_import_ui.notify.confirm", return_value=True):
            dialog.delete_row()
        self.assertEqual([row["file_nguon"] for row in dialog.rows], ["third.jpg", "second.jpg"])
        self.assertEqual(dialog.summary_values["total"].text(), "2")
        self.assertEqual(dialog.summary_values["warnings"].text(), "0")
        self.assertEqual(dialog.summary_metrics["warnings"].property("tone"), "quiet")
        self.assertEqual(dialog.type_values[DOC_EXPENSE].text(), "2")

    def test_review_table_shows_close_date_cargo_and_readable_warning(self) -> None:
        payload = json.loads(self.path.read_text(encoding="utf-8"))
        row = payload["du_lieu_boc_tach"][0]
        row.update({
            "ngay_chay": "08/05/2026", "loai_hang": "Vôi rời",
            "so_hd": "123", "don_gia": 90000,
            "canh_bao": [{"noi_dung": "Cần đối chiếu hóa đơn"}],
        })
        self.path.write_text(json.dumps(payload, ensure_ascii=False), encoding="utf-8")
        dialog = JsonExtractDialog(str(self.path))
        self.assertIn("Loại hàng", dialog.COLUMNS)
        self.assertIn("Thông tin chung", dialog.COLUMNS)
        self.assertIn("Ngày đóng", dialog.COLUMNS)
        self.assertEqual(dialog._row_cells(0, row, DOC_EXPENSE)[4], "03/05/2026")
        self.assertEqual(dialog._row_cells(0, row, DOC_EXPENSE)[5], "Vôi rời")
        self.assertIn("HĐ 123", dialog._row_cells(0, row, DOC_EXPENSE)[6])
        self.assertIn("Cần đối chiếu hóa đơn", dialog._row_tooltip(row, ""))
        self.assertNotIn("{", dialog._row_tooltip(row, ""))
        self.assertEqual(dialog.table.columnWidth(2), 160)
        self.assertEqual(dialog.summary_values["ok"].text(), "1")

    def test_low_confidence_is_visible_as_warning(self) -> None:
        payload = json.loads(self.path.read_text(encoding="utf-8"))
        payload["du_lieu_boc_tach"][0]["do_tin_cay"] = "Thấp"
        self.path.write_text(json.dumps(payload, ensure_ascii=False), encoding="utf-8")
        dialog = JsonExtractDialog(str(self.path))
        self.assertEqual(dialog.summary_values["warnings"].text(), "2")
        self.assertEqual(dialog.summary_values["need_check"].text(), "2")
        self.assertEqual(dialog._row_cells(0, dialog.rows[0], DOC_EXPENSE)[9], "Cần kiểm tra")
        dialog.status_filter.setCurrentText("Cần kiểm tra")
        self.assertEqual(set(dialog._visible_indices), {0, 1})


if __name__ == "__main__":
    unittest.main()
