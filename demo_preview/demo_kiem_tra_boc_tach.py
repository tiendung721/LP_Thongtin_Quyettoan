"""Demo giao diện kiểm tra và sửa dòng bóc tách; chỉ dùng dữ liệu mẫu trong bộ nhớ.

Chạy: python demo_kiem_tra_boc_tach.py
Tạo ảnh xem trước: python demo_kiem_tra_boc_tach.py --preview
"""

from __future__ import annotations

import copy
import sys
from pathlib import Path

from PySide6.QtCore import Qt
from PySide6.QtGui import QColor, QFontDatabase, QFontMetrics, QIcon
from PySide6.QtWidgets import (
    QApplication, QDialog, QFrame, QGridLayout, QHBoxLayout, QHeaderView,
    QLabel, QLineEdit, QMessageBox, QPushButton, QScrollArea, QTableWidget,
    QTableWidgetItem, QVBoxLayout, QWidget,
)

from app.theme import apply_theme


ROOT = Path(__file__).resolve().parent
PREVIEW_DIR = ROOT / "demo_preview"

SAMPLES = [
    {
        "stt": "1", "loai": "Phiếu cân",
        "file": "J36_2804058796692620036_97c17269314e282c8b3d819101fead16(9).jpg",
        "ngay_dong": "26/06/2026", "container": "DRYU3085088",
        "bien_so_xe": "15C-08722", "so_chi_seal": "KIKAI846210",
        "so_tan": "27,18", "loai_hang": "VÔI RỜI L1", "noi_dong": "NAM ĐÔ",
        "nguoi_nhan": "Công ty Vật liệu A", "ten_tau": "", "ngay_chay": "",
        "vt_bien": "", "trang_thai": "Cần kiểm tra",
        "canh_bao": ["Số chì/Seal chưa rõ trên ảnh chứng từ.", "Cần xác nhận lại số tấn trước khi nhập."],
        "md5": "02d8925977b031d80648505bbfbd231e", "truong_khac": {"nguon": "demo"},
    },
    {
        "stt": "2", "loai": "Bill",
        "file": "BILL_VIMU2250494_Booking_20260627_scan_final.pdf",
        "ngay_dong": "26/06/2026", "container": "VIMU2250494",
        "so_bill": "BL260627014", "ten_tau": "M/V ORIENT STAR",
        "ngay_chay": "28/06/2026", "so_chi_seal": "SEAL582031",
        "so_tan": "27,18", "loai_hang": "VÔI RỜI L1", "noi_dong": "NAM ĐÔ",
        "nguoi_nhan": "Công ty Vật liệu A", "vt_bien": "Hải Phòng",
        "trang_thai": "OK", "canh_bao": [],
        "md5": "92a47dbfa55a021e", "truong_khac": {},
    },
    {
        "stt": "3", "loai": "Khoản chi",
        "file": "CHI_PHI_GO_PLET_THANG_06_2026_BAN_CHOT_CUOI_CUNG.xlsx",
        "ngay_dong": "26/06/2026", "container": "DRYU3085088",
        "ngay_chay": "28/06/2026", "loai_hang": "VÔI RỜI L1",
        "so_hd": "0002481", "gia_vat_lieu": "90.000", "don_gia": "120.000",
        "thanh_tien": "3.261.600", "vat": "326.160", "tong_tien": "3.587.760",
        "trang_thai": "OK", "canh_bao": [],
        "md5": "833a02e4f1c2", "truong_khac": {},
    },
]

COMMON = [
    ("ngay_dong", "Ngày đóng"), ("container", "Số Container"),
]
FIELDS = {
    "Phiếu cân": [
        ("bien_so_xe", "Biển số xe"), ("so_chi_seal", "Số chì/Seal"),
        ("so_tan", "Số tấn"), ("loai_hang", "Loại hàng"),
        ("noi_dong", "Nơi đóng"), ("nguoi_nhan", "Người nhận"),
        ("ten_tau", "Tên tàu"), ("ngay_chay", "Ngày chạy"),
        ("vt_bien", "VT biển"),
    ],
    "Bill": [
        ("so_bill", "Số B/L"), ("ten_tau", "Tên tàu"),
        ("ngay_chay", "Ngày chạy"), ("so_chi_seal", "Số chì/Seal"),
        ("so_tan", "Số tấn"), ("loai_hang", "Loại hàng"),
        ("noi_dong", "Nơi đóng"), ("nguoi_nhan", "Người nhận"),
        ("vt_bien", "VT biển"),
    ],
    "Khoản chi": [
        ("ngay_chay", "Ngày chạy"), ("loai_hang", "Loại hàng"),
        ("so_hd", "Số HĐ"), ("gia_vat_lieu", "Giá vật liệu"),
        ("don_gia", "Đơn giá"), ("thanh_tien", "Thành tiền"),
        ("vat", "VAT"), ("tong_tien", "Tổng tiền"),
    ],
}

DEMO_STYLE = """
QDialog#demoEdit, QDialog#demoTable { background: #F4F6F9; }
QFrame#demoCard { background: #FFFFFF; border: 1px solid #DFE6F0; border-radius: 10px; }
QLabel#demoTitle { color: #173B65; font-size: 16pt; font-weight: 700; }
QLabel#demoSection { color: #173B65; font-size: 11pt; font-weight: 700; }
QLabel#demoMuted { color: #64748B; font-size: 9pt; }
QLabel#demoBadge { border-radius: 10px; padding: 5px 10px; font-weight: 700; }
QLabel#demoBadge[tone="ok"] { color: #15803D; background: #DCFCE7; }
QLabel#demoBadge[tone="warning"] { color: #92400E; background: #FEF3C7; }
QLabel#demoBadge[tone="error"] { color: #B91C1C; background: #FEE2E2; }
QLabel#demoWarning { padding: 10px; border-radius: 8px; color: #92400E; background: #FFFBEB; border: 1px solid #FDE68A; }
QLabel#demoWarning[tone="ok"] { color: #15803D; background: #F0FDF4; border-color: #BBF7D0; }
QLabel#demoFieldLabel { color: #475569; font-weight: 600; }
QPushButton#demoFile { text-align: left; color: #1D4ED8; background: #F8FAFC; }
"""


def info_for(row: dict) -> str:
    if row["loai"] == "Phiếu cân":
        return f"Xe {row.get('bien_so_xe', '—')}  •  Seal {row.get('so_chi_seal', '—')}"
    if row["loai"] == "Bill":
        return f"B/L {row.get('so_bill', '—')}  •  {row.get('ten_tau', '—')}"
    return f"HĐ {row.get('so_hd', '—')}  •  Đơn giá {row.get('don_gia', '—')}"


def label(text: str, name: str = "") -> QLabel:
    result = QLabel(text)
    if name:
        result.setObjectName(name)
    return result


class EditDemo(QDialog):
    def __init__(self, row: dict, parent=None):
        super().__init__(parent)
        self.row = copy.deepcopy(row)
        self.editors: dict[str, QLineEdit] = {}
        self.setObjectName("demoEdit")
        self.setWindowTitle(f"Sửa dòng bóc tách · {row['loai']} · Demo")
        self.resize(760, 820)
        self.setMinimumSize(620, 620)

        outer = QVBoxLayout(self)
        outer.setContentsMargins(18, 15, 18, 14)
        outer.setSpacing(11)
        heading = QHBoxLayout()
        heading.addWidget(label("Sửa dòng bóc tách", "demoTitle"))
        heading.addStretch()
        heading.addWidget(label(row["loai"], "demoBadge"))
        outer.addLayout(heading)
        outer.addWidget(label("Dữ liệu mẫu · thay đổi chỉ tồn tại trong bản demo", "demoMuted"))

        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.NoFrame)
        body = QWidget()
        body_layout = QVBoxLayout(body)
        body_layout.setContentsMargins(0, 0, 5, 4)
        body_layout.setSpacing(11)

        overview = QFrame()
        overview.setObjectName("demoCard")
        ov = QVBoxLayout(overview)
        ov.setContentsMargins(14, 12, 14, 12)
        ov.addWidget(label("CHỨNG TỪ", "demoSection"))
        ov.addWidget(label(f"Dòng {row['stt']}  •  {row['loai']}", "demoMuted"))
        file_button = QPushButton()
        file_button.setObjectName("demoFile")
        file_button.setText(QFontMetrics(file_button.font()).elidedText(row["file"], Qt.ElideMiddle, 580))
        file_button.setToolTip(row["file"])
        file_button.clicked.connect(lambda: QMessageBox.information(self, "Tên file nguồn đầy đủ", row["file"]))
        ov.addWidget(file_button)
        body_layout.addWidget(overview)

        status = QFrame()
        status.setObjectName("demoCard")
        st = QVBoxLayout(status)
        st.setContentsMargins(14, 12, 14, 12)
        top = QHBoxLayout()
        top.addWidget(label("TRẠNG THÁI VÀ CẢNH BÁO", "demoSection"))
        top.addStretch()
        warnings = row.get("canh_bao") or []
        tone = "warning" if warnings else "ok"
        badge = label("CẦN KIỂM TRA" if warnings else "OK", "demoBadge")
        badge.setProperty("tone", tone)
        top.addWidget(badge)
        st.addLayout(top)
        warning = label("\n".join(f"• {item}" for item in warnings) if warnings else "✓ Không có cảnh báo", "demoWarning")
        warning.setProperty("tone", tone)
        warning.setWordWrap(True)
        st.addWidget(warning)
        body_layout.addWidget(status)

        self._fields_card(body_layout, "THÔNG TIN CHUNG", COMMON)
        self._fields_card(body_layout, f"THÔNG TIN {row['loai'].upper()}", FIELDS[row["loai"]])
        body_layout.addStretch()
        scroll.setWidget(body)
        outer.addWidget(scroll, 1)

        footer = QHBoxLayout()
        footer.addWidget(label("Các trường kỹ thuật được giữ trong dữ liệu nhưng không hiển thị.", "demoMuted"))
        footer.addStretch()
        done = QPushButton("Xong")
        done.setProperty("variant", "primary")
        done.clicked.connect(self._done)
        cancel = QPushButton("Hủy")
        cancel.clicked.connect(self.reject)
        footer.addWidget(done)
        footer.addWidget(cancel)
        outer.addLayout(footer)

    def _fields_card(self, parent: QVBoxLayout, title: str, fields: list[tuple[str, str]]) -> None:
        card = QFrame()
        card.setObjectName("demoCard")
        vertical = QVBoxLayout(card)
        vertical.setContentsMargins(14, 12, 14, 14)
        vertical.setSpacing(9)
        vertical.addWidget(label(title, "demoSection"))
        grid = QGridLayout()
        grid.setHorizontalSpacing(13)
        grid.setVerticalSpacing(10)
        for index, (key, title) in enumerate(fields):
            column = index % 2
            row = index // 2
            field = QVBoxLayout()
            field.setSpacing(3)
            field.addWidget(label(title, "demoFieldLabel"))
            editor = QLineEdit(str(self.row.get(key) or ""))
            editor.setPlaceholderText("Chưa có dữ liệu")
            self.editors[key] = editor
            field.addWidget(editor)
            grid.addLayout(field, row, column)
        vertical.addLayout(grid)
        parent.addWidget(card)

    def _done(self) -> None:
        for key, editor in self.editors.items():
            self.row[key] = editor.text().strip()
        self.accept()


class TableDemo(QDialog):
    HEADERS = [
        "STT", "Loại chứng từ", "File nguồn", "Container", "Ngày đóng",
        "Loại hàng", "Thông tin chung", "Số tấn / Tổng tiền", "Giá VL",
        "Trạng thái", "Cảnh báo",
    ]

    def __init__(self):
        super().__init__()
        self.rows = copy.deepcopy(SAMPLES)
        self.setObjectName("demoTable")
        self.setWindowTitle("Kiểm tra dữ liệu bóc tách · Demo đề xuất")
        self.resize(1530, 670)
        self.setMinimumSize(1120, 500)
        outer = QVBoxLayout(self)
        outer.setContentsMargins(14, 13, 14, 13)
        outer.setSpacing(10)
        outer.addWidget(label("Kiểm tra dữ liệu bóc tách", "demoTitle"))
        outer.addWidget(label("Dữ liệu mẫu · bấm vào tên file để xem đầy đủ · nhấp đúp một dòng để xem cửa sổ sửa", "demoMuted"))

        self.table = QTableWidget(len(self.rows), len(self.HEADERS))
        self.table.setHorizontalHeaderLabels(self.HEADERS)
        self.table.setAlternatingRowColors(True)
        self.table.setSelectionBehavior(QTableWidget.SelectRows)
        self.table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.table.verticalHeader().hide()
        header = self.table.horizontalHeader()
        header.setSectionResizeMode(QHeaderView.Fixed)
        for column, width in enumerate((52, 120, 165, 132, 112, 122, 255, 145, 92, 120, 110)):
            self.table.setColumnWidth(column, width)
        header.setSectionResizeMode(6, QHeaderView.Stretch)
        self.table.cellClicked.connect(self._clicked)
        self.table.cellDoubleClicked.connect(lambda row, _column: self._open(row))
        outer.addWidget(self.table, 1)
        self._populate()

        footer = QHBoxLayout()
        footer.addWidget(label("3 dòng mẫu: Phiếu cân · Bill · Khoản chi", "demoMuted"))
        footer.addStretch()
        edit = QPushButton("Sửa dòng")
        edit.setProperty("variant", "primary")
        edit.clicked.connect(lambda: self._open(max(0, self.table.currentRow())))
        footer.addWidget(edit)
        outer.addLayout(footer)
        self.table.selectRow(0)

    def _populate(self) -> None:
        for index, row in enumerate(self.rows):
            kind = row["loai"]
            amount = (row.get("tong_tien") or "") if kind == "Khoản chi" else (row.get("so_tan") or "") + " tấn"
            values = [
                row["stt"], kind, row["file"], row["container"], row["ngay_dong"],
                row.get("loai_hang", ""), info_for(row), amount,
                row.get("gia_vat_lieu", ""), row["trang_thai"],
                f"{len(row['canh_bao'])} cảnh báo" if row["canh_bao"] else "OK",
            ]
            for column, value in enumerate(values):
                item = QTableWidgetItem(value)
                if column == 2:
                    item.setToolTip(row["file"] + "\nBấm để xem tên đầy đủ")
                if column == 10:
                    item.setForeground(QColor("#B45309" if row["canh_bao"] else "#15803D"))
                if column == 9:
                    item.setForeground(QColor("#B45309" if row["trang_thai"] != "OK" else "#15803D"))
                self.table.setItem(index, column, item)
            self.table.setRowHeight(index, 42)

    def _clicked(self, row: int, column: int) -> None:
        if column == 2:
            QMessageBox.information(self, "Tên file nguồn đầy đủ", self.rows[row]["file"])

    def _open(self, index: int) -> None:
        dialog = EditDemo(self.rows[index], self)
        if dialog.exec() == QDialog.Accepted:
            self.rows[index] = dialog.row
            self._populate()


def main() -> int:
    app = QApplication(sys.argv)
    # Qt offscreen trên Windows không tự liệt kê system fonts; nạp font để ảnh PNG có dấu.
    for font_name in ("segoeui.ttf", "segoeuib.ttf"):
        font_path = Path("C:/Windows/Fonts") / font_name
        if font_path.is_file():
            QFontDatabase.addApplicationFont(str(font_path))
    apply_theme(app)
    app.setStyleSheet(app.styleSheet() + DEMO_STYLE)
    icon = QIcon(str(ROOT / "Assets" / "KIKAI_PM.ico"))
    if not icon.isNull():
        app.setWindowIcon(icon)
    window = TableDemo()
    window.show()
    if "--preview" in sys.argv:
        PREVIEW_DIR.mkdir(exist_ok=True)
        app.processEvents()
        window.grab().save(str(PREVIEW_DIR / "demo_boc_tach_bang.png"))
        for row, slug in zip(window.rows, ("phieu_can", "bill", "khoan_chi")):
            dialog = EditDemo(row)
            dialog.show()
            app.processEvents()
            dialog.grab().save(str(PREVIEW_DIR / f"demo_boc_tach_{slug}.png"))
            dialog.close()
        return 0
    return app.exec()


if __name__ == "__main__":
    sys.exit(main())
