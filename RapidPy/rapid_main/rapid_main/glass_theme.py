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

    @staticmethod
    def _blob(painter: QtGui.QPainter, center: QtCore.QPointF, radius: float, rgba: tuple[int, int, int, int]) -> None:
        gradient = QtGui.QRadialGradient(center, radius)
        color = QtGui.QColor(*rgba)
        gradient.setColorAt(0.0, color)
        middle = QtGui.QColor(color)
        middle.setAlpha(int(rgba[3] * 0.5))
        gradient.setColorAt(0.5, middle)
        edge = QtGui.QColor(color)
        edge.setAlpha(0)
        gradient.setColorAt(1.0, edge)
        painter.setBrush(QtGui.QBrush(gradient))
        painter.drawEllipse(center, radius, radius)

    def paintEvent(self, event: QtGui.QPaintEvent) -> None:  # type: ignore[override]
        """A soft, macOS-wallpaper-like colour field for the frosted panels.

        Qt cannot blur what is behind a widget, so the "vibrancy" comes from
        large, low-contrast colour fields that read through translucent
        sidebar, toolbar and tiles.
        """
        painter = QtGui.QPainter(self)
        painter.setRenderHint(QtGui.QPainter.RenderHint.Antialiasing, True)
        w, h = float(self.width()), float(self.height())
        base = QtGui.QLinearGradient(0, 0, w, h)
        base.setColorAt(0.0, QtGui.QColor("#F7ECEA"))
        base.setColorAt(0.5, QtGui.QColor("#EFE6EE"))
        base.setColorAt(1.0, QtGui.QColor("#F4EADB"))
        painter.fillRect(self.rect(), base)
        painter.setPen(QtCore.Qt.PenStyle.NoPen)
        span = max(w, h)
        self._blob(painter, QtCore.QPointF(w * 0.12, h * 0.10), span * 0.55, (122, 2, 25, 52))
        self._blob(painter, QtCore.QPointF(w * 0.88, h * 0.08), span * 0.45, (253, 181, 21, 70))
        self._blob(painter, QtCore.QPointF(w * 0.72, h * 0.95), span * 0.50, (124, 108, 196, 42))
        self._blob(painter, QtCore.QPointF(w * 0.05, h * 0.95), span * 0.35, (15, 118, 110, 30))
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


#: macOS-style glass layer for rapid_main, applied after every other theme.
#: Translucent "vibrancy" sidebar and unified toolbar, window-like tiles with
#: traffic lights, capsule status pills, hairline separators, 7 px controls
#: with an accent focus ring, overlay scrollbars and rounded menus.
MACOS_GLASS_QSS = f"""
/* Base reset: the shared theme paints every widget opaque beige, which hides
   the backdrop.  Containers and labels become transparent so the colour field
   reads through; every surface that must stay opaque is re-declared below. */
QWidget {{ background: transparent; color: #2b2326; }}
QMainWindow, QDialog, QMessageBox, QInputDialog, QProgressDialog, QFileDialog {{ background: #f6f1f0; }}
QAbstractItemView {{ background: rgba(255, 255, 255, 240); }}
QPlainTextEdit, QTextEdit, QTextBrowser {{
    background: rgba(255, 255, 255, 205);
    border: 1px solid rgba(0, 0, 0, 26);
    border-radius: 9px;
    selection-background-color: {TOKENS.maroon};
    selection-color: white;
}}
QGroupBox {{
    background: rgba(255, 255, 255, 95);
    border: 1px solid rgba(0, 0, 0, 22);
    border-radius: 10px;
    margin-top: 14px;
    padding-top: 6px;
}}
QGroupBox::title {{ subcontrol-origin: margin; left: 10px; padding: 0 4px; color: #4a3f42; font-weight: 600; }}
QTabWidget::pane {{ background: rgba(255, 255, 255, 120); border: 1px solid rgba(0, 0, 0, 22); border-radius: 10px; }}
QTabBar::tab {{
    background: rgba(255, 255, 255, 120);
    border: 1px solid rgba(0, 0, 0, 22);
    border-radius: 7px;
    padding: 4px 14px;
    margin: 2px 3px 4px 0px;
    color: #4a3f42;
}}
QTabBar::tab:selected {{ background: rgba(255, 255, 255, 240); color: {TOKENS.maroon}; font-weight: 650; }}
QProgressBar {{
    background: rgba(255, 255, 255, 170);
    border: 1px solid rgba(0, 0, 0, 24);
    border-radius: 6px;
    text-align: center;
}}
QProgressBar::chunk {{ background: {TOKENS.maroon}; border-radius: 5px; }}
QMainWindow#rapidMainWindow {{ background: #efe6e4; }}
QMenuBar {{
    background: rgba(255, 255, 255, 120);
    border-bottom: 1px solid rgba(0, 0, 0, 18);
    padding: 1px 6px;
}}
QMenuBar::item {{ background: transparent; padding: 3px 9px; border-radius: 5px; color: #2b2326; }}
QMenuBar::item:selected {{ background: rgba(0, 0, 0, 22); }}
QMenu {{
    background: rgba(250, 248, 248, 246);
    border: 1px solid rgba(0, 0, 0, 30);
    border-radius: 10px;
    padding: 5px;
}}
QMenu::item {{ padding: 5px 22px 5px 14px; border-radius: 6px; color: #2b2326; }}
QMenu::item:selected {{ background: {TOKENS.maroon}; color: white; }}
QMenu::separator {{ height: 1px; background: rgba(0, 0, 0, 22); margin: 4px 8px; }}

QFrame#header {{
    background: rgba(255, 255, 255, 150);
    border: none;
    border-bottom: 1px solid rgba(0, 0, 0, 22);
    border-radius: 0px;
}}
QLabel#headerTitle {{ color: #2b2326; font-size: 14px; font-weight: 700; letter-spacing: 0.2px; }}
QPushButton#headerBtn, QPushButton#headerBtnHalt {{
    min-height: 24px;
    padding: 3px 12px;
    border-radius: 7px;
    background: rgba(255, 255, 255, 170);
    border: 1px solid rgba(0, 0, 0, 30);
    border-bottom-color: rgba(0, 0, 0, 46);
    color: #2b2326;
    font-size: 12px;
    font-weight: 560;
}}
QPushButton#headerBtn:hover, QPushButton#headerBtnHalt:hover {{ background: rgba(255, 255, 255, 235); }}
QPushButton#headerBtn:pressed, QPushButton#headerBtnHalt:pressed {{ background: rgba(0, 0, 0, 22); }}
QPushButton#headerBtn:checked {{ background: rgba(0, 0, 0, 30); color: #1f1a1c; }}
QPushButton#headerBtnHalt {{ color: #c0262d; font-weight: 650; }}
QPushButton#headerBtnExit {{
    min-height: 24px;
    padding: 3px 13px;
    border-radius: 7px;
    color: white;
    font-weight: 650;
    border: 1px solid rgba(76, 0, 16, 160);
    background: qlineargradient(x1:0, y1:0, x2:0, y2:1, stop:0 #9a1030, stop:1 {TOKENS.maroon});
}}
QPushButton#headerBtnExit:hover {{ background: qlineargradient(x1:0, y1:0, x2:0, y2:1, stop:0 #ad1838, stop:1 #880422); }}

QFrame#sidebar {{
    background: rgba(246, 240, 241, 150);
    border: none;
    border-right: 1px solid rgba(0, 0, 0, 20);
    border-radius: 0px;
}}
QLabel#sectionHdr {{
    color: rgba(60, 50, 54, 130);
    background: transparent;
    font-size: 10px;
    font-weight: 700;
    letter-spacing: 0.6px;
}}
QPushButton#navBtn {{
    min-height: 28px;
    background: transparent;
    border: none;
    border-radius: 7px;
    padding: 4px 10px;
    text-align: left;
    color: #2b2326;
    font-size: 12px;
    font-weight: 500;
}}
QPushButton#navBtn:hover {{ background: rgba(0, 0, 0, 16); }}
QPushButton#navBtn:checked {{
    color: white;
    font-weight: 650;
    background: qlineargradient(x1:0, y1:0, x2:0, y2:1, stop:0 #94102e, stop:1 {TOKENS.maroon});
}}
QPushButton#navBtn:focus {{ border: 2px solid rgba(122, 2, 25, 110); }}

QFrame#card, QFrame#livePanel {{
    background: rgba(255, 255, 255, 168);
    border: 1px solid rgba(255, 255, 255, 210);
    border-bottom-color: rgba(0, 0, 0, 22);
    border-radius: 12px;
}}
QPushButton {{
    min-height: 20px;
    padding: 5px 13px;
    border-radius: 7px;
    background: rgba(255, 255, 255, 215);
    border: 1px solid rgba(0, 0, 0, 30);
    border-bottom-color: rgba(0, 0, 0, 48);
    color: #2b2326;
}}
QPushButton:hover {{ background: rgba(255, 255, 255, 245); }}
QPushButton:pressed {{ background: rgba(0, 0, 0, 24); }}
QPushButton:disabled {{
    color: rgba(43, 35, 38, 95);
    background: rgba(255, 255, 255, 110);
    border-color: rgba(0, 0, 0, 18);
}}
QPushButton#accent {{
    color: white;
    font-weight: 650;
    border: 1px solid rgba(76, 0, 16, 150);
    background: qlineargradient(x1:0, y1:0, x2:0, y2:1, stop:0 #9a1030, stop:1 {TOKENS.maroon});
}}
QPushButton#accent:hover {{ background: qlineargradient(x1:0, y1:0, x2:0, y2:1, stop:0 #ad1838, stop:1 #880422); }}
QPushButton:focus, QLineEdit:focus, QComboBox:focus, QSpinBox:focus,
QDoubleSpinBox:focus, QPlainTextEdit:focus, QTextEdit:focus {{
    border: 2px solid rgba(122, 2, 25, 120);
}}
QLineEdit, QComboBox, QSpinBox, QDoubleSpinBox {{
    min-height: 22px;
    padding: 4px 8px;
    border-radius: 7px;
    background: rgba(255, 255, 255, 235);
    border: 1px solid rgba(0, 0, 0, 34);
    selection-background-color: {TOKENS.maroon};
    selection-color: white;
}}
QCheckBox::indicator, QRadioButton::indicator {{
    width: 14px; height: 14px;
    background: rgba(255, 255, 255, 240);
    border: 1px solid rgba(0, 0, 0, 60);
}}
QCheckBox::indicator {{ border-radius: 4px; }}
QRadioButton::indicator {{ border-radius: 7px; }}
QCheckBox::indicator:checked, QRadioButton::indicator:checked {{
    background: {TOKENS.maroon};
    border: 1px solid rgba(76, 0, 16, 200);
}}
QHeaderView::section {{
    background: rgba(255, 255, 255, 170);
    border: none;
    border-right: 1px solid rgba(0, 0, 0, 16);
    border-bottom: 1px solid rgba(0, 0, 0, 26);
    padding: 5px 7px;
    color: #4a3f42;
    font-weight: 600;
}}
QTableWidget, QTableView, QListWidget, QTreeView {{
    background: rgba(255, 255, 255, 190);
    alternate-background-color: rgba(246, 242, 242, 190);
    border: 1px solid rgba(0, 0, 0, 24);
    border-radius: 9px;
    gridline-color: rgba(0, 0, 0, 16);
    selection-background-color: rgba(122, 2, 25, 200);
    selection-color: white;
}}
QScrollBar:vertical {{ background: transparent; width: 10px; margin: 2px; border: none; }}
QScrollBar:horizontal {{ background: transparent; height: 10px; margin: 2px; border: none; }}
QScrollBar::handle:vertical, QScrollBar::handle:horizontal {{
    background: rgba(0, 0, 0, 70);
    border: none;
    border-radius: 3px;
    min-height: 30px;
    min-width: 30px;
    margin: 1px;
}}
QScrollBar::handle:vertical:hover, QScrollBar::handle:horizontal:hover {{ background: rgba(0, 0, 0, 110); }}
QScrollBar::add-line, QScrollBar::sub-line {{ width: 0px; height: 0px; border: none; background: transparent; }}
QScrollBar::add-page, QScrollBar::sub-page {{ background: transparent; }}
QStatusBar {{
    background: rgba(255, 255, 255, 130);
    border-top: 1px solid rgba(0, 0, 0, 20);
    color: #4a3f42;
}}
QToolTip {{
    color: #2b2326;
    background: rgba(252, 250, 250, 250);
    border: 1px solid rgba(0, 0, 0, 40);
    border-radius: 6px;
    padding: 4px 7px;
}}

QFrame#tile[chrome="mac"] {{
    background: rgba(252, 249, 249, 150);
    border: 1px solid rgba(0, 0, 0, 30);
    border-radius: 12px;
}}
QFrame#tile[chrome="mac"][active="true"] {{ background: rgba(253, 251, 251, 190); }}
QFrame#tile[chrome="mac"] QWidget#tileHeader {{
    background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
        stop:0 rgba(255, 255, 255, 200), stop:1 rgba(245, 240, 241, 170));
    border: none;
    border-bottom: 1px solid rgba(0, 0, 0, 20);
    border-top-left-radius: 12px;
    border-top-right-radius: 12px;
}}
QFrame#tile[chrome="mac"] QLabel#tileTitle {{
    color: rgba(43, 35, 38, 120);
    font-size: 12px;
    font-weight: 600;
    background: transparent;
}}
QFrame#tile[chrome="mac"][active="true"] QLabel#tileTitle {{ color: #2b2326; }}
QPushButton#workspaceButton {{
    border-radius: 7px;
    background: rgba(255, 255, 255, 140);
    border: 1px solid rgba(0, 0, 0, 26);
    color: rgba(43, 35, 38, 120);
    font-weight: 600;
}}
QPushButton#workspaceButton[occupied="true"] {{ color: #2b2326; background: rgba(255, 255, 255, 215); }}
QPushButton#workspaceButton:checked {{
    color: white;
    border-color: rgba(76, 0, 16, 160);
    background: qlineargradient(x1:0, y1:0, x2:0, y2:1, stop:0 #9a1030, stop:1 {TOKENS.maroon});
}}
QDialog#tileLauncher {{
    background: rgba(250, 248, 248, 248);
    border: 1px solid rgba(0, 0, 0, 40);
    border-radius: 14px;
}}
QDialog#glassDialog {{
    background: qlineargradient(x1:0, y1:0, x2:1, y2:1, stop:0 #f8f1f0, stop:0.55 #f1eaf0, stop:1 #f5ede0);
}}
QDialog#glassDialog QFrame#dialogCard {{
    background: rgba(255, 255, 255, 190);
    border: 1px solid rgba(255, 255, 255, 220);
    border-bottom-color: rgba(0, 0, 0, 22);
    border-radius: 12px;
}}
"""


def apply_macos_glass_theme(app: QtWidgets.QApplication) -> None:
    """Append the macOS glass layer; call last so it wins over earlier layers."""

    font = QtGui.QFont("SF Pro Text", 10)
    for family in ("SF Pro Text", "Segoe UI Variable Text", "Segoe UI"):
        candidate = QtGui.QFont(family, 10)
        if QtGui.QFontInfo(candidate).family().lower().startswith(family.split()[0].lower()):
            font = candidate
            break
    app.setFont(font)
    if MACOS_GLASS_QSS not in app.styleSheet():
        app.setStyleSheet(app.styleSheet() + MACOS_GLASS_QSS)
