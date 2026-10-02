from __future__ import annotations

from dataclasses import dataclass

from PySide6 import QtCore, QtGui, QtWidgets


@dataclass(frozen=True)
class GlassTokens:
    """Single source of truth for the rapid_main visual language."""

    maroon: str = "#7A0219"
    maroon_dark: str = "#4C0010"
    gold: str = "#FDB515"
    ink: str = "#261E21"
    muted: str = "#6F6265"
    surface: str = "rgba(255, 255, 255, 214)"
    surface_strong: str = "rgba(255, 255, 255, 235)"
    border: str = "rgba(255, 255, 255, 178)"
    outline: str = "rgba(122, 2, 25, 56)"


TOKENS = GlassTokens()


class GlassBackdrop(QtWidgets.QWidget):
    """Paint the non-animated layered background behind the main workspace."""

    def __init__(self, parent: QtWidgets.QWidget | None = None) -> None:
        super().__init__(parent)
        self.setObjectName("appBackdrop")
        self.setAttribute(QtCore.Qt.WidgetAttribute.WA_StyledBackground, True)

    def paintEvent(self, event: QtGui.QPaintEvent) -> None:  # type: ignore[override]
        painter = QtGui.QPainter(self)
        painter.setRenderHint(QtGui.QPainter.RenderHint.Antialiasing, True)

        base = QtGui.QLinearGradient(0, 0, self.width(), self.height())
        base.setColorAt(0.0, QtGui.QColor("#F8F4EC"))
        base.setColorAt(0.46, QtGui.QColor("#EFE7DF"))
        base.setColorAt(1.0, QtGui.QColor("#E7DDD8"))
        painter.fillRect(self.rect(), base)

        painter.setPen(QtCore.Qt.PenStyle.NoPen)
        painter.setBrush(QtGui.QColor(122, 2, 25, 25))
        painter.drawEllipse(
            QtCore.QRectF(
                -self.width() * 0.12,
                -self.height() * 0.22,
                self.width() * 0.72,
                self.height() * 0.78,
            )
        )
        painter.setBrush(QtGui.QColor(253, 181, 21, 30))
        painter.drawEllipse(
            QtCore.QRectF(
                self.width() * 0.57,
                -self.height() * 0.18,
                self.width() * 0.58,
                self.height() * 0.64,
            )
        )
        painter.setBrush(QtGui.QColor(70, 100, 112, 18))
        painter.drawEllipse(
            QtCore.QRectF(
                self.width() * 0.48,
                self.height() * 0.56,
                self.width() * 0.64,
                self.height() * 0.58,
            )
        )
        painter.end()
        super().paintEvent(event)


MAIN_GLASS_QSS = f"""
QMainWindow#rapidMainWindow {{
    background: #eee6df;
}}
QWidget#appBackdrop {{
    background: transparent;
}}
QWidget#appBackdrop QStackedWidget,
QWidget#appBackdrop QStackedWidget > QWidget {{
    background: transparent;
}}
QFrame#sidebar {{
    background: rgba(255, 255, 255, 168);
    border: 1px solid rgba(255, 255, 255, 196);
    border-right-color: rgba(122, 2, 25, 46);
    border-radius: 0px 22px 22px 0px;
}}
QFrame#header {{
    background: rgba(255, 255, 255, 205);
    border: 1px solid rgba(255, 255, 255, 220);
    border-top: 3px solid {TOKENS.maroon};
    border-bottom-color: rgba(122, 2, 25, 46);
    border-radius: 0px;
}}
QFrame#card, QFrame#livePanel {{
    background: {TOKENS.surface};
    border: 1px solid {TOKENS.border};
    border-bottom-color: rgba(122, 2, 25, 42);
    border-radius: 20px;
}}
QFrame#card QWidget, QFrame#livePanel QWidget {{
    background: transparent;
}}
QFrame#card QTableView, QFrame#card QTableWidget,
QFrame#card QListView, QFrame#card QTreeView,
QFrame#card QPlainTextEdit, QFrame#card QTextEdit {{
    background: rgba(255, 255, 255, 178);
    border: 1px solid rgba(122, 2, 25, 45);
    border-radius: 12px;
}}
QPushButton {{
    min-height: 22px;
    background: rgba(255, 255, 255, 178);
    border: 1px solid rgba(122, 2, 25, 62);
    border-radius: 11px;
    padding: 8px 13px;
    color: {TOKENS.ink};
}}
QPushButton:hover {{
    background: rgba(255, 255, 255, 226);
    border-color: rgba(122, 2, 25, 105);
}}
QPushButton:pressed {{
    background: rgba(242, 229, 225, 235);
}}
QPushButton:focus, QLineEdit:focus, QComboBox:focus,
QSpinBox:focus, QDoubleSpinBox:focus, QTableView:focus,
QPlainTextEdit:focus, QTextEdit:focus {{
    border: 2px solid {TOKENS.maroon};
}}
QPushButton:disabled {{
    color: rgba(56, 46, 48, 105);
    background: rgba(255, 255, 255, 92);
    border-color: rgba(80, 60, 65, 35);
}}
QPushButton#accent, QPushButton#headerBtnExit {{
    color: white;
    font-weight: 700;
    background: qlineargradient(
        x1:0, y1:0, x2:1, y2:1,
        stop:0 {TOKENS.maroon}, stop:1 {TOKENS.maroon_dark}
    );
    border: 1px solid rgba(255, 255, 255, 90);
}}
QPushButton#accent:hover, QPushButton#headerBtnExit:hover {{
    background: qlineargradient(
        x1:0, y1:0, x2:1, y2:1,
        stop:0 #930425, stop:1 #600014
    );
}}
QPushButton#navBtn {{
    min-height: 30px;
    background: transparent;
    border: 1px solid transparent;
    border-radius: 12px;
    padding: 7px 12px;
    text-align: left;
    color: #493B3E;
    font-size: 12px;
    font-weight: 560;
}}
QPushButton#navBtn:hover {{
    background: rgba(255, 255, 255, 145);
    border-color: rgba(122, 2, 25, 38);
}}
QPushButton#navBtn:checked {{
    background: rgba(122, 2, 25, 27);
    border-color: rgba(122, 2, 25, 70);
    color: {TOKENS.maroon};
    font-weight: 700;
}}
QPushButton#navBtn:focus {{
    border: 2px solid {TOKENS.maroon};
}}
QLineEdit, QComboBox, QDoubleSpinBox, QSpinBox {{
    min-height: 24px;
    background: rgba(255, 255, 255, 205);
    border: 1px solid rgba(122, 2, 25, 55);
    border-radius: 10px;
    padding: 6px 8px;
    selection-background-color: {TOKENS.maroon};
    selection-color: white;
}}
QLabel#title {{
    color: {TOKENS.maroon};
    font-size: 25px;
    font-weight: 750;
}}
QLabel#subtitle, QLabel#readLbl {{
    color: {TOKENS.muted};
}}
QLabel#sectionHdr {{
    color: #76666A;
    font-size: 10px;
    font-weight: 750;
    letter-spacing: 1.3px;
}}
QLabel#valuePill, QLabel#valueMonospace, QLabel#readingDisplay {{
    background: rgba(255, 255, 255, 178);
    border: 1px solid rgba(122, 2, 25, 45);
    border-radius: 10px;
}}
QHeaderView::section {{
    background: rgba(255, 255, 255, 205);
    border: none;
    border-right: 1px solid rgba(122, 2, 25, 30);
    border-bottom: 1px solid rgba(122, 2, 25, 48);
    padding: 7px;
    color: #493B3E;
    font-weight: 650;
}}
QTableWidget, QTableView {{
    alternate-background-color: rgba(248, 242, 236, 178);
    gridline-color: rgba(122, 2, 25, 26);
    selection-background-color: rgba(253, 181, 21, 100);
    selection-color: {TOKENS.ink};
}}
QToolTip {{
    color: {TOKENS.ink};
    background: rgba(255, 255, 255, 245);
    border: 1px solid rgba(122, 2, 25, 72);
    border-radius: 7px;
    padding: 5px 7px;
}}
"""


def apply_main_glass_theme(app: QtWidgets.QApplication) -> None:
    """Append rapid_main-specific glass styling after the shared application theme."""

    if MAIN_GLASS_QSS not in app.styleSheet():
        app.setStyleSheet(app.styleSheet() + MAIN_GLASS_QSS)


def install_glass_elevation(root: QtWidgets.QWidget) -> int:
    """Apply restrained elevation to main content cards, once per card."""

    count = 0
    for card in root.findChildren(QtWidgets.QFrame):
        if card.objectName() not in {"card", "livePanel"}:
            continue
        if card.graphicsEffect() is not None:
            continue
        shadow = QtWidgets.QGraphicsDropShadowEffect(card)
        shadow.setBlurRadius(26)
        shadow.setOffset(0, 7)
        shadow.setColor(QtGui.QColor(48, 30, 35, 42))
        card.setGraphicsEffect(shadow)
        count += 1
    return count
