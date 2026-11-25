"""界面统一视觉：配色、字体与 Qt 样式表（沿用原 tkinter 版配色）。"""

from __future__ import annotations

FONT_CANDIDATES = ("Microsoft YaHei UI", "Microsoft YaHei", "微软雅黑", "SimHei", "SimSun")

COLORS = {
    "primary": "#1F4E79",
    "primary_hover": "#2A6399",
    "primary_soft": "#EAF1F8",
    "bg": "#F7F8FA",
    "card": "#FFFFFF",
    "border": "#E3E6EA",
    "text": "#1F2937",
    "muted": "#6B7280",
    "success": "#15803D",
    "warning": "#B45309",
    "error": "#B91C1C",
    "disabled": "#9AA3AF",
    "nav_hover": "#F2F6FA",
}

_family = FONT_CANDIDATES[0]


def font_family() -> str:
    return _family


def _qss() -> str:
    c = COLORS
    return f"""
QWidget {{ background: {c['bg']}; color: {c['text']}; }}
QMainWindow, QStackedWidget, QScrollArea {{ background: {c['bg']}; }}
QScrollArea {{ border: none; }}
QLabel {{ background: transparent; }}
QFrame#Card {{ background: {c['card']}; border: 1px solid {c['border']}; border-radius: 6px; }}
QFrame#Card QLabel, QFrame#Card QCheckBox {{ background: transparent; }}
QLabel#PageTitle {{ font-size: 18px; font-weight: 600; }}
QLabel#Desc, QLabel#CardDesc {{ color: {c['muted']}; }}
QLabel#SectionTitle {{ font-size: 15px; font-weight: 600; }}
QLabel#Status {{ color: {c['muted']}; font-size: 12px; }}
QLabel#NavTitle {{ color: {c['muted']}; font-size: 12px; font-weight: 600; }}
QLabel#PathBox {{ background: #FFFFFF; border: 1px solid {c['border']}; border-radius: 4px; padding: 4px 6px; }}
QPushButton {{ background: {c['card']}; border: 1px solid {c['border']}; border-radius: 6px; padding: 8px 18px; font-size: 14px; }}
QPushButton:hover {{ background: {c['nav_hover']}; }}
QPushButton:disabled {{ background: {c['bg']}; color: {c['disabled']}; }}
QPushButton#Primary {{ background: {c['primary']}; color: #FFFFFF; border: none; font-size: 16px; font-weight: 600; padding: 10px 28px; }}
QPushButton#Primary:hover {{ background: {c['primary_hover']}; }}
QPushButton#Primary:disabled {{ background: #B7C4D1; color: #F2F5F8; }}
QPushButton#Nav {{ text-align: left; border: none; background: {c['card']}; padding: 10px 14px; font-size: 15px; border-radius: 6px; }}
QPushButton#Nav:hover {{ background: {c['nav_hover']}; }}
QPushButton#Nav:checked {{ background: {c['primary_soft']}; color: {c['primary']}; font-weight: 600; }}
QPushButton#LogTool {{ padding: 4px 10px; font-size: 12px; }}
QCheckBox {{ background: transparent; spacing: 6px; }}
QRadioButton {{ background: transparent; spacing: 8px; }}
QRadioButton::indicator {{ width: 16px; height: 16px; border: 2px solid {c['disabled']}; border-radius: 9px; background: {c['card']}; }}
QRadioButton::indicator:hover {{ border-color: {c['primary_hover']}; }}
QRadioButton::indicator:checked {{ border: 5px solid {c['primary']}; border-radius: 13px; background: {c['card']}; }}
QRadioButton:checked {{ color: {c['primary']}; font-weight: 600; }}
QPlainTextEdit#LogText {{ background: #FBFCFD; border: 1px solid {c['border']}; border-radius: 4px; font-size: 13px; }}
QLineEdit, QSpinBox, QComboBox {{ background: #FFFFFF; border: 1px solid {c['border']}; border-radius: 4px; padding: 4px 6px; }}
QListWidget#ManualChapters {{ background: #FFFFFF; border: 1px solid {c['border']}; border-radius: 6px; padding: 4px; outline: 0; }}
QListWidget#ManualChapters::item {{ padding: 8px 10px; border-radius: 4px; }}
QListWidget#ManualChapters::item:selected {{ background: {c['primary_soft']}; color: {c['primary']}; font-weight: 600; }}
QTextBrowser#ManualView {{ background: #FFFFFF; border: 1px solid {c['border']}; border-radius: 6px; }}
QTabWidget::pane {{ border: 1px solid {c['border']}; border-radius: 6px; background: {c['card']}; top: -1px; }}
QTabBar::tab {{ background: {c['bg']}; border: 1px solid {c['border']}; border-bottom: none; padding: 8px 20px; border-top-left-radius: 6px; border-top-right-radius: 6px; }}
QTabBar::tab:selected {{ background: {c['card']}; color: {c['primary']}; font-weight: 600; }}
QStatusBar {{ background: {c['bg']}; }}
QStatusBar QLabel {{ background: transparent; }}
QProgressBar {{ border: 1px solid {c['border']}; border-radius: 4px; background: {c['card']}; }}
QProgressBar::chunk {{ background: {c['primary']}; border-radius: 3px; }}
QSplitter::handle {{ background: {c['border']}; }}
QScrollBar:vertical {{ background: {c['bg']}; width: 10px; }}
QScrollBar::handle:vertical {{ background: {c['border']}; border-radius: 5px; min-height: 30px; }}
QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {{ height: 0; }}
QScrollBar:horizontal {{ background: {c['bg']}; height: 10px; }}
QScrollBar::handle:horizontal {{ background: {c['border']}; border-radius: 5px; min-width: 30px; }}
QScrollBar::add-line:horizontal, QScrollBar::sub-line:horizontal {{ width: 0; }}
"""


def apply(app) -> str:
    """应用字体与样式表，返回实际使用的字体名。"""
    global _family
    from PySide6.QtGui import QFont, QFontDatabase

    families = set(QFontDatabase.families())
    for name in FONT_CANDIDATES:
        if name in families:
            _family = name
            break
    # 整数像素字号 + 完整 hinting：高分屏下避免 pt 换算出小数像素导致笔画发虚
    font = QFont(_family)
    font.setPixelSize(14)
    font.setHintingPreference(QFont.PreferFullHinting)
    app.setFont(font)
    app.setStyleSheet(_qss())
    return _family