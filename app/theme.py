"""Bộ giao diện (QSS) tông sáng, hiện đại cho Trợ Lý Quyết Toán RPA.

Tách phần trang trí (màu sắc, bo góc, khoảng cách) khỏi logic giao diện để dễ
chỉnh sửa. Chỉ cần sửa ``STYLESHEET`` bên dưới là đổi được diện mạo toàn app.
"""

from __future__ import annotations

from PySide6.QtWidgets import QApplication, QWidget

# Bảng màu chủ đạo (dùng lại ở nơi cần đặt màu bằng code).
COLORS = {
    "bg": "#F4F6F9",
    "card": "#FFFFFF",
    "border": "#E5E7EB",
    "text": "#1F2937",
    "muted": "#6B7280",
    "primary": "#2563EB",
    "success": "#16A34A",
    "danger": "#DC2626",
}

STYLESHEET = """
* {
    font-family: "Segoe UI", "Segoe UI Variable", Arial, sans-serif;
    font-size: 10.5pt;
    color: #1F2937;
}

QWidget#appRoot { background: #F4F6F9; }
QWidget#page { background: #F4F6F9; }
QStackedWidget#content { background: #F4F6F9; }

/* ---------- Thanh điều hướng bên trái ---------- */
QWidget#sidebar {
    background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
        stop:0 #152D4D, stop:1 #0C1A31);
    border-right: 1px solid #172F50;
}
QWidget#brandBanner { background: transparent; }
QLabel#appTitle { font-size: 15pt; font-weight: 700; color: #FFFFFF; padding: 20px 16px 2px 16px; }
QLabel#appSubtitle { font-size: 9pt; color: #B8C9E2; padding: 0 16px 14px 16px; }
QListWidget#navList { background: transparent; border: none; outline: 0; }
QListWidget#navList::item { padding: 11px 14px; margin: 3px 10px; border-radius: 8px; color: #FFFFFF; }
QListWidget#navList::item:hover { background: #25436B; }
QListWidget#navList::item:selected { background: #2563EB; color: #FFFFFF; font-weight: 600; }
QLabel#navFooter { color: #9EB2D0; font-size: 8.5pt; padding: 10px 16px 14px 16px; }

/* ---------- Thẻ (card) ---------- */
QFrame#card { background: #FFFFFF; border: 1px solid #E5E7EB; border-radius: 12px; }
QLabel#cardTitle { font-size: 11.5pt; font-weight: 700; color: #111827; }
QLabel#cardHint { color: #6B7280; }

/* ---------- Nhãn trợ giúp ---------- */
QLabel#statusText { color: #1D4ED8; font-weight: 600; }
QLabel#fileName { font-size: 11pt; font-weight: 600; color: #111827; }
QLabel#metaText { color: #6B7280; font-size: 9.5pt; }
QLabel#noteText { color: #B91C1C; }
QLabel#hintText { color: #475569; }
QLabel#formLabel { color: #374151; }

/* ---------- Thẻ theo bước: viền nổi khi đang làm ---------- */
QFrame#card[active="true"] { border: 1px solid #93C5FD; background: #FCFDFF; }

/* ---------- Huy hiệu số bước ---------- */
QLabel#stepBadge {
    border-radius: 13px; font-size: 10pt; font-weight: 700;
    background: #F1F5F9; color: #64748B;
}
QLabel#stepBadge[tone="done"]   { background: #16A34A; color: #FFFFFF; }
QLabel#stepBadge[tone="active"] { background: #2563EB; color: #FFFFFF; }
QLabel#stepBadge[tone="wait"]   { background: #F1F5F9; color: #94A3B8; }
QLabel#stepBadge[tone="warn"]   { background: #F59E0B; color: #FFFFFF; }
QLabel#stepBadge[tone="error"]  { background: #DC2626; color: #FFFFFF; }

/* ---------- Chip trạng thái ---------- */
QLabel#statusChip { padding: 3px 10px; border-radius: 10px; font-size: 9pt; font-weight: 600; }
QLabel#statusChip[tone="done"]   { background: #DCFCE7; color: #15803D; }
QLabel#statusChip[tone="active"] { background: #DBEAFE; color: #1D4ED8; }
QLabel#statusChip[tone="wait"]   { background: #F1F5F9; color: #64748B; }
QLabel#statusChip[tone="warn"]   { background: #FEF3C7; color: #B45309; }
QLabel#statusChip[tone="error"]  { background: #FEE2E2; color: #B91C1C; }
QLabel#statusChip[tone="locked"] { background: #F1F5F9; color: #94A3B8; }

/* ---------- Nút bật/tắt hướng dẫn (dạng chữ) ---------- */
QPushButton#guideToggle {
    background: transparent; border: none; color: #6B7280;
    padding: 2px 6px; font-size: 9pt; text-align: right;
}
QPushButton#guideToggle:hover { color: #2563EB; background: transparent; }

/* ---------- Hộp hướng dẫn (bung khi cần) ---------- */
QFrame#guideBox { background: #F8FAFC; border: 1px solid #E5E7EB; border-radius: 8px; }
QLabel#guideText { color: #475569; }

/* ---------- Tổng hợp ngắn ở màn hình kiểm tra chứng từ ---------- */
QFrame#extractSummary { background: #F8FAFC; border: 1px solid #E5E7EB; border-radius: 8px; }
QFrame#extractMetric { background: #FFFFFF; border: 1px solid #E2E8F0; border-radius: 6px; }
QFrame#extractMetric[tone="success"] { background: #F0FDF4; border-color: #BBF7D0; }
QFrame#extractMetric[tone="warning"] { background: #FFFBEB; border-color: #FDE68A; }
QLabel#extractMetricValue { color: #1D4ED8; font-size: 15pt; font-weight: 700; }
QLabel#extractMetricValue[tone="success"] { color: #15803D; }
QLabel#extractMetricValue[tone="warning"] { color: #B45309; }
QLabel#extractMetricValue[tone="quiet"] { color: #64748B; }
QLabel#extractMetricLabel { color: #64748B; font-size: 8.5pt; }
QFrame#extractTypeChip { background: #EEF2F7; border: 1px solid #E2E8F0; border-radius: 6px; }
QLabel#extractTypeLabel { color: #475569; font-size: 8.5pt; }
QLabel#extractTypeValue { color: #1E3A5F; font-size: 9pt; font-weight: 700; }

/* ---------- Nút bấm ---------- */
QPushButton {
    background: #FFFFFF; color: #374151; border: 1px solid #D1D5DB;
    border-radius: 8px; padding: 7px 14px; font-size: 10pt;
}
QPushButton:hover { background: #F3F4F6; }
QPushButton:pressed { background: #E5E7EB; }
QPushButton:disabled { color: #9CA3AF; background: #F3F4F6; border-color: #E5E7EB; }

QPushButton[variant="primary"] { background: #2563EB; color: #FFFFFF; border: none; }
QPushButton[variant="primary"]:hover { background: #1D4ED8; }
QPushButton[variant="primary"]:pressed { background: #1E40AF; }
QPushButton[variant="primary"]:disabled { background: #BFDBFE; color: #F8FAFC; }

QPushButton[variant="success"] { background: #16A34A; color: #FFFFFF; border: none; }
QPushButton[variant="success"]:hover { background: #15803D; }
QPushButton[variant="success"]:pressed { background: #166534; }
QPushButton[variant="success"]:disabled { background: #BBF7D0; color: #F8FAFC; }

QPushButton[variant="danger"] { background: #DC2626; color: #FFFFFF; border: none; }
QPushButton[variant="danger"]:hover { background: #B91C1C; }
QPushButton[variant="danger"]:pressed { background: #991B1B; }

/* Nút hành động chính của mỗi bước: cỡ vừa, không kéo giãn hết thẻ. */
QPushButton#actionButton {
    min-height: 20px; font-size: 10.5pt; font-weight: 600; padding: 9px 20px;
}

/* ---------- Ô nhập ---------- */
QLineEdit { background: #FFFFFF; border: 1px solid #D1D5DB; border-radius: 8px; padding: 7px 10px; }
QLineEdit:focus { border: 1px solid #2563EB; }
/* Ô sửa mở ngay trong bảng/danh sách (inline editor): bỏ padding & bo góc lớn
   để chữ không bị cắt/che trong dòng thấp. */
QAbstractItemView QLineEdit {
    padding: 0px 3px;
    border-radius: 0px;
    min-height: 0px;
    background: #FFFFFF;
}
QCheckBox { spacing: 8px; color: #374151; }

/* ---------- Bảng chọn SQT và khung chi tiết ---------- */
QLabel#sqtDialogTitle { font-size: 16pt; font-weight: 700; color: #102F54; }
QLabel#sqtDialogKind, QLabel#sqtFilterLabel { color: #476889; font-size: 9pt; }
QLabel#sqtCountBadge { color: #1D4ED8; background: #E6F0FF; padding: 7px 11px; border-radius: 9px; font-weight: 600; }
QLabel#sqtSourceLabel { background: #FFFFFF; border: 1px solid #D1D5DB; border-radius: 8px; padding: 9px 10px; }
QLabel#sqtTopSummary { color: #173B65; background: #EFF6FF; border: 1px solid #D3E4FA; border-radius: 8px; padding: 8px 12px; font-weight: 600; }
QLabel#sqtSelectionSummary, QLabel#sqtDetailMeta { color: #55708D; font-size: 9.5pt; }
QLabel#sqtDetailTitle { color: #12365F; font-size: 12pt; font-weight: 700; }
QLabel#sqtDetailStatus { border-radius: 8px; padding: 4px 9px; font-weight: 600; }
QLabel#sqtDetailStatus[tone="pending"], QLabel#sqtDetailStatus[tone="partial"] { color: #9A5B0A; background: #FFF0D8; }
QLabel#sqtDetailStatus[tone="done"] { color: #157347; background: #DCF5E7; }
QLabel#sqtDetailRowTitle, QLabel#sqtDetailAmount { color: #153B68; font-weight: 700; }
QFrame#sqtDetailPanel, QFrame#sqtDetailCard { background: #FFFFFF; border: 1px solid #D8E3F0; border-radius: 9px; }
QFrame#sqtDivider { background: #DBE5F0; min-height: 1px; max-height: 1px; }
QWidget#sqtDetailHost { background: #F6F8FB; }
QTableWidget#sqtSelectionTable {
    background: #FFFFFF; alternate-background-color: #F8FAFD;
    border: 1px solid #D8E3F0; border-radius: 8px;
    gridline-color: #E0E8F2; selection-background-color: #DCEAFF;
    selection-color: #12365F; outline: 0;
}
QTableWidget#sqtSelectionTable::item { padding: 6px 8px; }
QTableWidget#sqtSelectionTable::item:selected { background: #DCEAFF; color: #12365F; }
QTableWidget#sqtSelectionTable QHeaderView::section {
    background: #EAF0F7; color: #173D68; border: 0;
    border-right: 1px solid #D6E1EE; border-bottom: 1px solid #D6E1EE;
    padding: 8px 5px; font-weight: 600;
}

/* ---------- Thanh tiến trình ---------- */
QProgressBar {
    background: #EEF2F7; border: none; border-radius: 3px;
    max-height: 6px; min-height: 6px;
}
QProgressBar::chunk { background: #2563EB; border-radius: 3px; }

/* ---------- Thanh tiến trình các bước ---------- */
QLabel#stepCircle { border-radius: 16px; font-weight: 700; font-size: 11pt; }
QLabel#stepCircle[state="pending"] { background: #E5E7EB; color: #9CA3AF; }
QLabel#stepCircle[state="active"] { background: #2563EB; color: #FFFFFF; }
QLabel#stepCircle[state="done"] { background: #16A34A; color: #FFFFFF; }
QLabel#stepLabel { color: #9CA3AF; font-size: 9.5pt; }
QLabel#stepLabel[state="active"] { color: #1D4ED8; font-weight: 600; }
QLabel#stepLabel[state="done"] { color: #16A34A; }
QFrame#stepLine { background: #E5E7EB; border: none; }

/* ---------- Bảng ---------- */
QTableWidget {
    background: #FFFFFF; border: 1px solid #E5E7EB; border-radius: 8px;
    gridline-color: #F1F3F5; alternate-background-color: #FAFBFC;
}
QTableWidget::item { padding: 4px 6px; }
QTableWidget::item:selected { background: #EFF6FF; color: #1F2937; }
QHeaderView::section {
    background: #F9FAFB; color: #374151; padding: 8px; border: none;
    border-bottom: 1px solid #E5E7EB; font-weight: 600;
}
QTableCornerButton::section { background: #F9FAFB; border: none; }

/* ---------- Nhật ký ---------- */
QPlainTextEdit {
    background: #F9FAFB; color: #374151; border: 1px solid #E5E7EB;
    border-radius: 8px; font-family: "Consolas", "Courier New", monospace; font-size: 9.5pt;
}

/* ---------- Vùng cuộn ---------- */
QScrollArea { border: none; background: transparent; }
QScrollBar:vertical { background: transparent; width: 12px; margin: 2px; }
QScrollBar::handle:vertical { background: #CBD1DA; border-radius: 5px; min-height: 30px; }
QScrollBar::handle:vertical:hover { background: #AEB6C2; }
QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical { height: 0; }
QScrollBar::add-page:vertical, QScrollBar::sub-page:vertical { background: transparent; }

/* ---------- Hộp thông báo của ứng dụng ---------- */
QDialog#notificationDialog { background: #FFFFFF; }
QFrame#notificationAccent { border: none; background: #2563EB; }
QFrame#notificationAccent[tone="success"] { background: #16A34A; }
QFrame#notificationAccent[tone="warning"] { background: #F59E0B; }
QFrame#notificationAccent[tone="error"] { background: #DC2626; }
QFrame#notificationAccent[tone="question"] { background: #2563EB; }
QLabel#notificationTitle { color: #111827; font-size: 14pt; font-weight: 700; }
QLabel#notificationMessage { color: #374151; font-size: 10.5pt; line-height: 1.4; }
QLabel#notificationContext { color: #6B7280; font-size: 9pt; }
QFrame#notificationNotice {
    background: #FFFBEB; border: 1px solid #FDE68A; border-radius: 8px;
}
QLabel#notificationNoticeText { color: #92400E; }
QFrame#notificationMetric {
    background: #F8FAFC; border: 1px solid #E5E7EB; border-radius: 8px;
}
QFrame#notificationMetric[tone="success"] { background: #F0FDF4; border-color: #BBF7D0; }
QFrame#notificationMetric[tone="warning"] { background: #FFFBEB; border-color: #FDE68A; }
QFrame#notificationMetric[tone="error"] { background: #FEF2F2; border-color: #FECACA; }
QLabel#notificationMetricValue { color: #1D4ED8; font-size: 15pt; font-weight: 700; }
QLabel#notificationMetricValue[tone="success"] { color: #15803D; }
QLabel#notificationMetricValue[tone="warning"] { color: #B45309; }
QLabel#notificationMetricValue[tone="error"] { color: #B91C1C; }
QLabel#notificationMetricLabel { color: #64748B; font-size: 8.5pt; }
QToolButton#notificationDetailsToggle {
    background: transparent; border: none; color: #475569; padding: 3px 0;
    font-weight: 600;
}
QToolButton#notificationDetailsToggle:hover { color: #2563EB; }
QFrame#notificationDetails {
    background: #F8FAFC; border: 1px solid #E5E7EB; border-radius: 8px;
}
"""


def apply_theme(app: QApplication) -> None:
    """Áp bộ style tông sáng cho toàn ứng dụng."""
    app.setStyleSheet(STYLESHEET)


def repolish(widget: QWidget) -> None:
    """Ép Qt vẽ lại widget sau khi đổi thuộc tính động (dynamic property)."""
    style = widget.style()
    style.unpolish(widget)
    style.polish(widget)
    widget.update()
