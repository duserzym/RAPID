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


_STATUS_PREFIXES = {
    "neutral": "STATUS",
    "ready": "READY",
    "active": "ACTIVE",
    "warning": "WARNING",
    "error": "ERROR",
    "unavailable": "UNAVAILABLE",
    "simulated": "SIMULATED",
}


def set_semantic_status(
    widget: QtWidgets.QLabel,
    text: str,
    level: str = "neutral",
    *,
    accessible_name: str = "Status",
    show_prefix: bool = True,
) -> None:
    """Set a visible, accessible, stylesheet-driven semantic state.

    The textual prefix keeps state understandable without color.  Callers may
    suppress it for a numeric readout only when a nearby status label carries
    the same state in words.
    """

    normalized = str(level).strip().lower()
    if normalized not in _STATUS_PREFIXES:
        raise ValueError(f"Unknown semantic status level: {level!r}")

    message = str(text).strip() or "No status available"
    prefix = _STATUS_PREFIXES[normalized]
    visible = f"{prefix} — {message}" if show_prefix else message
    widget.setObjectName("statusText" if show_prefix else widget.objectName())
    widget.setProperty("status", normalized)
    widget.setText(visible)
    widget.setAccessibleName(accessible_name)
    widget.setAccessibleDescription(f"{prefix.title()} state. {message}")

    style = widget.style()
    style.unpolish(widget)
    style.polish(widget)
    widget.update()


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
QDialog#glassDialog {{
    color: {TOKENS.ink};
    background: qlineargradient(
        x1:0, y1:0, x2:1, y2:1,
        stop:0 #f8f4ec, stop:0.55 #eee6df, stop:1 #e8dcd8
    );
}}
QDialog#glassDialog QFrame#dialogCard {{
    background: {TOKENS.surface};
    border: 1px solid {TOKENS.border};
    border-bottom-color: rgba(122, 2, 25, 42);
    border-radius: 16px;
}}
QDialog#glassDialog QFrame#dialogHero {{
    background: qlineargradient(
        x1:0, y1:0, x2:1, y2:1,
        stop:0 {TOKENS.maroon_dark}, stop:1 {TOKENS.maroon}
    );
    border: none;
    border-bottom: 3px solid {TOKENS.gold};
}}
QDialog#glassDialog QWidget#dialogBody {{
    background: {TOKENS.surface};
    border: none;
}}
QDialog#glassDialog QScrollArea,
QDialog#glassDialog QScrollArea > QWidget > QWidget {{
    background: transparent;
    border: none;
}}
QDialog#glassDialog QLabel#dialogTitle {{
    color: {TOKENS.maroon};
    font-size: 16px;
    font-weight: 750;
}}
QDialog#glassDialog QLabel#dialogHeroTitle {{
    color: white;
    background: transparent;
    font-size: 28px;
    font-weight: 800;
}}
QDialog#glassDialog QLabel#dialogHeroSubtitle {{
    color: rgba(255, 255, 255, 190);
    background: transparent;
    font-size: 12px;
}}
QDialog#glassDialog QLabel#dialogSubtitle,
QDialog#glassDialog QLabel#unitLabel {{
    color: {TOKENS.muted};
}}
QDialog#glassDialog QLabel#guidanceText {{
    color: #4d3a39;
    background: rgba(255, 255, 255, 145);
    border: 1px solid rgba(122, 2, 25, 38);
    border-radius: 12px;
    padding: 10px 12px;
}}
QDialog#glassDialog QLabel#metaKey {{
    color: {TOKENS.muted};
    font-weight: 650;
}}
QDialog#glassDialog QLabel#metaValue {{
    color: {TOKENS.ink};
}}
QDialog#glassDialog QLabel#readingDisplay {{
    color: {TOKENS.ink};
    background: rgba(255, 255, 255, 155);
    border: 1px solid rgba(122, 2, 25, 35);
    border-radius: 12px;
    padding: 8px 12px;
    font-size: 32px;
    font-weight: 800;
}}
QDialog#glassDialog QLabel#readingDisplay[status="ready"] {{
    color: #17653a;
    border-color: rgba(23, 101, 58, 90);
}}
QDialog#glassDialog QLabel#readingDisplay[status="warning"],
QDialog#glassDialog QLabel#readingDisplay[status="error"] {{
    color: #9c241f;
    border-color: rgba(156, 36, 31, 105);
}}
QDialog#glassDialog QLabel#statusText {{
    color: #493b3e;
    background: rgba(255, 255, 255, 145);
    border: 1px solid rgba(73, 59, 62, 48);
    border-radius: 9px;
    padding: 6px 9px;
    font-size: 11px;
    font-weight: 650;
}}
QDialog#glassDialog QLabel#statusText[status="ready"] {{
    color: #14532d;
    background: rgba(220, 252, 231, 185);
    border-color: rgba(21, 128, 61, 80);
}}
QDialog#glassDialog QLabel#statusText[status="active"] {{
    color: #713f12;
    background: rgba(254, 249, 195, 190);
    border-color: rgba(202, 138, 4, 85);
}}
QDialog#glassDialog QLabel#statusText[status="warning"] {{
    color: #7c2d12;
    background: rgba(255, 237, 213, 190);
    border-color: rgba(180, 83, 9, 90);
}}
QDialog#glassDialog QLabel#statusText[status="error"],
QDialog#glassDialog QLabel#statusText[status="unavailable"] {{
    color: #7f1d1d;
    background: rgba(254, 226, 226, 195);
    border-color: rgba(185, 28, 28, 95);
}}
QDialog#glassDialog QLabel#statusText[status="simulated"] {{
    color: #4c1d95;
    background: rgba(237, 233, 254, 195);
    border-color: rgba(109, 40, 217, 85);
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
