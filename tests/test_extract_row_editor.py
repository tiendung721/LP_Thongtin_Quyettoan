from __future__ import annotations

import os
import unittest

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

from PySide6.QtWidgets import QApplication, QCheckBox

from app.daily_import import parse_money
from app.extract_row_editor import ExtractRowEditor, warning_text


class ExtractRowEditorTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        cls.app = QApplication.instance() or QApplication([])

    def test_form_changes_fields_with_document_type(self) -> None:
        row = {
            "loai_chung_tu": "Phiếu cân", "so_container": "ABCD1234567",
            "ngay_dong": "26/06/2026", "so_tan": 27.18, "ma_md5_file": "secret",
        }
        editor = ExtractRowEditor(row)
        self.assertIn("so_tan", editor.editors)
        self.assertNotIn("so_hd", editor.editors)
        self.assertNotIn("ma_md5_file", editor.editors)
        self.assertNotIn("truong_khac", editor.editors)
        editor.type_combo.setCurrentText("Khoản chi")
        self.assertIn("so_hd", editor.editors)
        self.assertNotIn("bien_so_xe", editor.editors)
        self.assertEqual(editor.updated_row()["ma_md5_file"], "secret")

    def test_warning_is_informational_and_clears_after_relevant_valid_edit(self) -> None:
        warning = {"noi_dung": "Kiểm tra VAT", "ma": "VAT_1"}
        row = {
            "loai_chung_tu": "Khoản chi", "so_container": "ABCD1234567",
            "ngay_dong": "26/06/2026", "ngay_chay": "28/06/2026",
            "thanh_tien": 100, "vat": 10, "tong_tien": 110,
            "trang_thai": "Cần kiểm tra", "canh_bao": [warning],
            "do_tin_cay": "thấp", "ma_md5_file": "keep-md5",
            "truong_khac": {"keep": True},
        }
        editor = ExtractRowEditor(row)
        self.assertEqual(warning_text(warning), "Kiểm tra VAT")
        self.assertEqual(editor.findChildren(QCheckBox), [])
        self.assertEqual(editor.updated_row()["canh_bao"], [warning])
        editor.editors["vat"].setText("12")
        self.assertEqual(editor.updated_row()["canh_bao"], [warning])
        editor.editors["tong_tien"].setText("112")
        updated = editor.updated_row()
        self.assertEqual(updated["canh_bao"], [])
        self.assertEqual(updated["trang_thai"], "OK")
        self.assertEqual(updated["do_tin_cay"], "Đã kiểm tra")
        self.assertEqual(updated["ma_md5_file"], "keep-md5")
        self.assertEqual(updated["truong_khac"], {"keep": True})

    def test_inconsistent_total_keeps_warning_and_review_status(self) -> None:
        row = {
            "loai_chung_tu": "Khoản chi", "so_container": "ABCD1234567",
            "ngay_chay": "28/06/2026", "thanh_tien": 100, "vat": 10,
            "tong_tien": 999, "trang_thai": "OK", "canh_bao": ["Kiểm tra tiền"],
        }
        editor = ExtractRowEditor(row)
        editor.editors["tong_tien"].setText("998")
        updated = editor.updated_row()
        self.assertEqual(updated["canh_bao"], ["Kiểm tra tiền"])
        self.assertEqual(updated["trang_thai"], "Cần kiểm tra")

    def test_money_warning_is_not_cleared_by_unrelated_edit(self) -> None:
        row = {
            "loai_chung_tu": "Khoản chi", "so_container": "ABCD1234567",
            "ngay_chay": "28/06/2026", "thanh_tien": 100, "vat": 10,
            "tong_tien": 110, "trang_thai": "Cần kiểm tra",
            "canh_bao": ["Kiểm tra tiền"],
        }
        editor = ExtractRowEditor(row)
        editor.editors["so_container"].setText("EFGH1234567")
        updated = editor.updated_row()
        self.assertEqual(updated["canh_bao"], ["Kiểm tra tiền"])
        self.assertEqual(updated["trang_thai"], "Cần kiểm tra")

    def test_container_and_seal_warning_clears_only_after_relevant_edit(self) -> None:
        warning = "Chuỗi chứa container và số chì không có nhãn/phân tách rõ; cần đối chiếu ảnh gốc."
        row = {
            "loai_chung_tu": "Phiếu cân", "so_container": "ABCD1234567",
            "so_chi_seal": "SEAL01", "ngay_dong": "26/06/2026",
            "loai_hang": "Vôi rời", "canh_bao": [warning], "trang_thai": "Cần kiểm tra",
        }
        editor = ExtractRowEditor(row)
        editor.editors["loai_hang"].setText("Vôi rời L1")
        self.assertEqual(editor.updated_row()["canh_bao"], [warning])
        editor.editors["so_chi_seal"].setText("SEAL02")
        updated = editor.updated_row()
        self.assertEqual(updated["canh_bao"], [])
        self.assertEqual(updated["trang_thai"], "OK")

    def test_unresolved_second_warning_prevents_ok(self) -> None:
        row = {
            "loai_chung_tu": "Phiếu cân", "so_container": "ABCD1234567",
            "so_chi_seal": "SEAL01", "ngay_dong": "26/06/2026",
            "bien_so_xe": "15C-12345", "trang_thai": "Cần kiểm tra",
            "canh_bao": ["Kiểm tra số chì", "Kiểm tra biển số xe"],
        }
        editor = ExtractRowEditor(row)
        editor.editors["so_chi_seal"].setText("SEAL02")
        updated = editor.updated_row()
        self.assertEqual(updated["canh_bao"], ["Kiểm tra biển số xe"])
        self.assertEqual(updated["trang_thai"], "Cần kiểm tra")

    def test_vietnamese_money_input_is_not_read_as_decimal(self) -> None:
        self.assertEqual(parse_money("90.000"), 90000)
        row = {"loai_chung_tu": "Khoản chi", "so_container": "ABCD1234567", "ngay_chay": "28/06/2026"}
        editor = ExtractRowEditor(row)
        editor.editors["gia_vat_lieu"].setText("90.000")
        editor.editors["tong_tien"].setText("3.587.760")
        updated = editor.updated_row()
        self.assertEqual(updated["gia_vat_lieu"], 90000)
        self.assertEqual(updated["tong_tien"], 3587760)


if __name__ == "__main__":
    unittest.main()
