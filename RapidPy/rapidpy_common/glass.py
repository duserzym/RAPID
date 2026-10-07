"""Glassmorphism visual language shared by the standalone RapidPy instrument apps.

Qt Widgets cannot sample and blur what is behind a widget, so the frosted look
is built from three layers that are cheap to paint:

* ``GlassBackdrop`` paints a warm gradient with large, soft radial colour
  fields (maroon, gold, teal).  Everything placed on it is translucent, so the
  fields read through the panels the way a blurred backdrop would.
* Cards are translucent white with a bright hairline border, a faint top sheen
  and a soft drop shadow, which gives depth without a real blur pass.
* Text, inputs and status chips stay high-contrast so readings and safety
  labels remain legible over the coloured backdrop.

``apply_glassmorphism_theme`` layers on top of ``apply_liquid_glass_theme`` so
the font selection, combo/spin arrows, scrollbars and window bounds guard are
shared with every other RapidPy app.  Only widgets inside a ``GlassBackdrop``
become transparent; dialogs, menus and popups keep their opaque surfaces.
"""

from __future__ import annotations

from dataclasses import dataclass

from PySide6 import QtCore, QtGui, QtWidgets

from .palette import MAROON
from .ui import apply_liquid_glass_theme


@dataclass(frozen=True)
class GlassPalette:
    maroon: str = MAROON
    maroon_dark: str = "#4C0010"
    gold: str = "#FDB515"
    teal: str = "#0F766E"
    ink: str = "#241C1E"
    muted: str = "#5F5154"
    plot_pens: tuple[str, ...] = ("#7A0219", "#0F766E", "#B7791F", "#3B5BA9", "#6B21A8")


GLASS = GlassPalette()

_STATUS_LEVELS = ("neutral", "ready", "active", "warning", "error")


class GlassBackdrop(QtWidgets.QWidget):
    """Paint the soft colour field that the translucent panels sit on."""

    def __init__(self, parent: QtWidgets.QWidget | None = None) -> None:
        super().__init__(parent)
        self.setObjectName("glassBackdrop")
        self.setAttribute(QtCore.Qt.WidgetAttribute.WA_StyledBackground, True)

    @staticmethod
    def _blob(painter: QtGui.QPainter, center: QtCore.QPointF, radius: float, rgba: tuple[int, int, int, int]) -> None:
        gradient = QtGui.QRadialGradient(center, radius)
        color = QtGui.QColor(*rgba)
        gradient.setColorAt(0.0, color)
        mid = QtGui.QColor(color)
        mid.setAlpha(int(rgba[3] * 0.45))
        gradient.setColorAt(0.55, mid)
        edge = QtGui.QColor(color)
        edge.setAlpha(0)
        gradient.setColorAt(1.0, edge)
        painter.setBrush(QtGui.QBrush(gradient))
        painter.drawEllipse(center, radius, radius)

    def paintEvent(self, event: QtGui.QPaintEvent) -> None:  # type: ignore[override]
        painter = QtGui.QPainter(self)
        painter.setRenderHint(QtGui.QPainter.RenderHint.Antialiasing, True)
        w, h = float(self.width()), float(self.height())

        base = QtGui.QLinearGradient(0, 0, w, h)
        base.setColorAt(0.0, QtGui.QColor("#FBF4EA"))
        base.setColorAt(0.45, QtGui.QColor("#F1E4DC"))
        base.setColorAt(1.0, QtGui.QColor("#E4DCE4"))
        painter.fillRect(self.rect(), base)

        painter.setPen(QtCore.Qt.PenStyle.NoPen)
        span = max(w, h)
        self._blob(painter, QtCore.QPointF(w * 0.10, h * 0.05), span * 0.48, (122, 2, 25, 70))
        self._blob(painter, QtCore.QPointF(w * 0.86, h * 0.12), span * 0.40, (253, 181, 21, 92))
        self._blob(painter, QtCore.QPointF(w * 0.62, h * 0.98), span * 0.46, (15, 118, 110, 52))
        self._blob(painter, QtCore.QPointF(w * 0.02, h * 0.92), span * 0.30, (253, 181, 21, 46))
        painter.end()
        super().paintEvent(event)


GLASSMORPHISM_QSS = f"""
QWidget#glassBackdrop,
QWidget#glassBackdrop QWidget {{
    background: transparent;
    color: {GLASS.ink};
}}
QWidget#glassBackdrop QFrame#glassHeader {{
    background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
        stop:0 rgba(255, 255, 255, 0.72), stop:1 rgba(255, 255, 255, 0.46));
    border: 1px solid rgba(255, 255, 255, 0.85);
    border-bottom: 1px solid rgba(122, 2, 25, 0.16);
    border-radius: 20px;
}}
QWidget#glassBackdrop QLabel#glassHeaderTitle {{
    color: {GLASS.maroon};
    font-size: 21px;
    font-weight: 780;
}}
QWidget#glassBackdrop QLabel#glassHeaderSubtitle {{
    color: {GLASS.muted};
}}
QWidget#glassBackdrop QLabel#glassHeaderBadge {{
    color: white;
    background: qlineargradient(x1:0, y1:0, x2:1, y2:1, stop:0 {GLASS.maroon}, stop:1 {GLASS.maroon_dark});
    border: 1px solid rgba(255, 255, 255, 0.55);
    border-radius: 12px;
    padding: 4px 10px;
    font-size: 11px;
    font-weight: 720;
    letter-spacing: 0.6px;
}}
QWidget#glassBackdrop QFrame#card {{
    background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
        stop:0 rgba(255, 255, 255, 0.70), stop:0.08 rgba(255, 255, 255, 0.58),
        stop:1 rgba(255, 255, 255, 0.44));
    border: 1px solid rgba(255, 255, 255, 0.88);
    border-bottom-color: rgba(122, 2, 25, 0.14);
    border-radius: 22px;
}}
QWidget#glassBackdrop QFrame#livePanel {{
    background: qlineargradient(x1:0, y1:0, x2:1, y2:1,
        stop:0 rgba(255, 255, 255, 0.62), stop:1 rgba(255, 255, 255, 0.34));
    border: 1px solid rgba(255, 255, 255, 0.90);
    border-radius: 20px;
}}
QWidget#glassBackdrop QLabel#title {{
    color: {GLASS.maroon};
    font-size: 21px;
    font-weight: 760;
}}
QWidget#glassBackdrop QLabel#cardTitle {{
    color: {GLASS.maroon};
    font-size: 15px;
    font-weight: 760;
}}
QWidget#glassBackdrop QLabel#subtitle {{
    color: {GLASS.muted};
}}
QWidget#glassBackdrop QLabel#sectionHeader {{
    color: {GLASS.maroon};
    font-size: 12px;
    font-weight: 760;
    letter-spacing: 0.4px;
}}
QWidget#glassBackdrop QLabel#hint {{
    color: {GLASS.muted};
    font-size: 11px;
}}
QWidget#glassBackdrop QLabel#mono {{
    font-family: Consolas, "Cascadia Mono", monospace;
}}
QWidget#glassBackdrop QLabel#valuePill,
QWidget#glassBackdrop QPlainTextEdit#valuePill,
QWidget#glassBackdrop QPlainTextEdit#statusPill {{
    background: rgba(255, 255, 255, 0.58);
    border: 1px solid rgba(255, 255, 255, 0.92);
    border-bottom-color: rgba(122, 2, 25, 0.14);
    border-radius: 14px;
    padding: 7px 10px;
    color: {GLASS.ink};
    font-weight: 620;
}}
QWidget#glassBackdrop QLabel#readingDisplay {{
    background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
        stop:0 rgba(255, 255, 255, 0.86), stop:1 rgba(255, 255, 255, 0.60));
    border: 1px solid rgba(255, 255, 255, 0.95);
    border-bottom-color: rgba(122, 2, 25, 0.20);
    border-radius: 20px;
    padding: 14px 18px;
    color: {GLASS.maroon};
}}
QWidget#glassBackdrop QLabel#statusChip {{
    background: rgba(255, 255, 255, 0.62);
    border: 1px solid rgba(73, 59, 62, 0.20);
    border-radius: 11px;
    padding: 4px 10px;
    color: #493b3e;
    font-weight: 680;
}}
QWidget#glassBackdrop QLabel#statusChip[status="ready"] {{
    color: #14532d; background: rgba(220, 252, 231, 0.80); border-color: rgba(21, 128, 61, 0.40);
}}
QWidget#glassBackdrop QLabel#statusChip[status="active"] {{
    color: #713f12; background: rgba(254, 249, 195, 0.82); border-color: rgba(202, 138, 4, 0.45);
}}
QWidget#glassBackdrop QLabel#statusChip[status="warning"] {{
    color: #7c2d12; background: rgba(255, 237, 213, 0.85); border-color: rgba(180, 83, 9, 0.45);
}}
QWidget#glassBackdrop QLabel#statusChip[status="error"] {{
    color: #7f1d1d; background: rgba(254, 226, 226, 0.88); border-color: rgba(185, 28, 28, 0.50);
}}
QWidget#glassBackdrop QPushButton {{
    background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
        stop:0 rgba(255, 255, 255, 0.82), stop:1 rgba(255, 255, 255, 0.56));
    border: 1px solid rgba(255, 255, 255, 0.95);
    border-bottom-color: rgba(122, 2, 25, 0.28);
    border-radius: 13px;
    padding: 8px 14px;
    color: {GLASS.ink};
    font-weight: 560;
}}
QWidget#glassBackdrop QPushButton:hover {{
    background: rgba(255, 255, 255, 0.94);
    border-color: rgba(122, 2, 25, 0.38);
}}
QWidget#glassBackdrop QPushButton:pressed {{
    background: rgba(244, 232, 226, 0.95);
}}
QWidget#glassBackdrop QPushButton:checked {{
    background: rgba(253, 181, 21, 0.42);
    border-color: rgba(122, 2, 25, 0.45);
    color: {GLASS.maroon_dark};
    font-weight: 700;
}}
QWidget#glassBackdrop QPushButton:focus {{
    border: 2px solid rgba(122, 2, 25, 0.70);
}}
QWidget#glassBackdrop QPushButton:disabled {{
    color: rgba(56, 46, 48, 0.42);
    background: rgba(255, 255, 255, 0.30);
    border-color: rgba(255, 255, 255, 0.55);
}}
QWidget#glassBackdrop QPushButton#accent {{
    color: #fff9eb;
    background: qlineargradient(x1:0, y1:0, x2:1, y2:1, stop:0 #8E0420, stop:1 {GLASS.maroon_dark});
    border: 1px solid rgba(255, 255, 255, 0.45);
    font-weight: 700;
}}
QWidget#glassBackdrop QPushButton#accent:hover {{
    background: qlineargradient(x1:0, y1:0, x2:1, y2:1, stop:0 #A1062A, stop:1 #600014);
}}
QWidget#glassBackdrop QPushButton#accent:disabled {{
    color: rgba(255, 255, 255, 0.70);
    background: rgba(122, 2, 25, 0.35);
}}
QWidget#glassBackdrop QPushButton#danger {{
    color: #7f1d1d;
    background: rgba(254, 226, 226, 0.78);
    border: 1px solid rgba(185, 28, 28, 0.45);
    font-weight: 700;
}}
QWidget#glassBackdrop QPushButton#danger:hover {{
    background: rgba(254, 202, 202, 0.92);
}}
QWidget#glassBackdrop QLineEdit,
QWidget#glassBackdrop QComboBox,
QWidget#glassBackdrop QSpinBox,
QWidget#glassBackdrop QDoubleSpinBox {{
    background: rgba(255, 255, 255, 0.78);
    border: 1px solid rgba(122, 2, 25, 0.22);
    border-radius: 11px;
    padding: 6px 8px;
    color: {GLASS.ink};
    selection-background-color: {GLASS.maroon};
    selection-color: white;
}}
QWidget#glassBackdrop QComboBox {{
    padding-right: 34px;
}}
QWidget#glassBackdrop QSpinBox,
QWidget#glassBackdrop QDoubleSpinBox {{
    padding-right: 30px;
}}
QWidget#glassBackdrop QLineEdit:focus,
QWidget#glassBackdrop QComboBox:focus,
QWidget#glassBackdrop QSpinBox:focus,
QWidget#glassBackdrop QDoubleSpinBox:focus {{
    border: 2px solid rgba(122, 2, 25, 0.70);
    background: rgba(255, 255, 255, 0.94);
}}
QWidget#glassBackdrop QLineEdit:disabled,
QWidget#glassBackdrop QComboBox:disabled,
QWidget#glassBackdrop QSpinBox:disabled,
QWidget#glassBackdrop QDoubleSpinBox:disabled {{
    color: rgba(56, 46, 48, 0.45);
    background: rgba(255, 255, 255, 0.40);
}}
QWidget#glassBackdrop QGroupBox {{
    background: rgba(255, 255, 255, 0.30);
    border: 1px solid rgba(255, 255, 255, 0.85);
    border-radius: 14px;
    margin-top: 14px;
    padding: 10px 8px 8px 8px;
    font-weight: 680;
}}
QWidget#glassBackdrop QGroupBox::title {{
    subcontrol-origin: margin;
    subcontrol-position: top left;
    left: 12px;
    padding: 0 4px;
    color: {GLASS.maroon};
}}
QWidget#glassBackdrop QPlainTextEdit,
QWidget#glassBackdrop QTextEdit,
QWidget#glassBackdrop QListWidget,
QWidget#glassBackdrop QTreeWidget {{
    background: rgba(255, 255, 255, 0.55);
    border: 1px solid rgba(255, 255, 255, 0.90);
    border-bottom-color: rgba(122, 2, 25, 0.14);
    border-radius: 14px;
    padding: 4px;
}}
QWidget#glassBackdrop QPlainTextEdit#console {{
    background: rgba(32, 22, 26, 0.84);
    color: #fff2c9;
    border: 1px solid rgba(255, 255, 255, 0.30);
    border-radius: 14px;
    padding: 8px;
    font-family: Consolas, "Cascadia Mono", monospace;
    selection-background-color: {GLASS.maroon};
}}
QWidget#glassBackdrop QTableWidget,
QWidget#glassBackdrop QTableView {{
    background: rgba(255, 255, 255, 0.60);
    alternate-background-color: rgba(255, 248, 240, 0.55);
    border: 1px solid rgba(255, 255, 255, 0.90);
    border-radius: 14px;
    gridline-color: rgba(122, 2, 25, 0.10);
    selection-background-color: rgba(253, 181, 21, 0.45);
    selection-color: {GLASS.ink};
}}
QWidget#glassBackdrop QHeaderView::section {{
    background: rgba(255, 255, 255, 0.72);
    border: none;
    border-right: 1px solid rgba(122, 2, 25, 0.10);
    border-bottom: 1px solid rgba(122, 2, 25, 0.20);
    padding: 6px;
    color: #4d3a39;
    font-weight: 660;
}}
QWidget#glassBackdrop QTabWidget::pane {{
    background: rgba(255, 255, 255, 0.36);
    border: 1px solid rgba(255, 255, 255, 0.88);
    border-radius: 16px;
    top: -1px;
}}
QWidget#glassBackdrop QTabBar::tab {{
    background: rgba(255, 255, 255, 0.42);
    border: 1px solid rgba(255, 255, 255, 0.85);
    border-radius: 11px;
    padding: 7px 16px;
    margin: 2px 4px 4px 0px;
    color: {GLASS.muted};
    font-weight: 620;
}}
QWidget#glassBackdrop QTabBar::tab:selected {{
    background: rgba(255, 255, 255, 0.90);
    color: {GLASS.maroon};
    border-bottom: 2px solid {GLASS.maroon};
}}
QWidget#glassBackdrop QSplitter::handle {{
    background: transparent;
}}
QWidget#glassBackdrop QSplitter::handle:hover {{
    background: rgba(122, 2, 25, 0.16);
    border-radius: 3px;
}}
QWidget#glassBackdrop QCheckBox,
QWidget#glassBackdrop QRadioButton {{
    spacing: 7px;
}}
QWidget#glassBackdrop QCheckBox::indicator,
QWidget#glassBackdrop QRadioButton::indicator {{
    width: 15px;
    height: 15px;
    background: rgba(255, 255, 255, 0.85);
    border: 1px solid rgba(122, 2, 25, 0.45);
}}
QWidget#glassBackdrop QCheckBox::indicator {{
    border-radius: 4px;
}}
QWidget#glassBackdrop QRadioButton::indicator {{
    border-radius: 8px;
}}
QWidget#glassBackdrop QCheckBox::indicator:checked,
QWidget#glassBackdrop QRadioButton::indicator:checked {{
    background: qradialgradient(cx:0.5, cy:0.5, radius:0.6, fx:0.5, fy:0.5,
        stop:0 {GLASS.maroon}, stop:0.55 {GLASS.maroon}, stop:0.62 {GLASS.gold}, stop:1 {GLASS.gold});
    border: 1px solid {GLASS.maroon};
}}
QWidget#glassBackdrop QProgressBar {{
    background: rgba(255, 255, 255, 0.55);
    border: 1px solid rgba(255, 255, 255, 0.90);
    border-radius: 8px;
    text-align: center;
    min-height: 14px;
}}
QWidget#glassBackdrop QProgressBar::chunk {{
    background: qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 {GLASS.maroon}, stop:1 {GLASS.gold});
    border-radius: 7px;
}}
QWidget#glassBackdrop QScrollArea {{
    border: none;
}}
"""


def apply_glassmorphism_theme(app: QtWidgets.QApplication) -> None:
    """Apply the shared RapidPy theme plus the glassmorphism layer (idempotent)."""

    apply_liquid_glass_theme(app)
    if GLASSMORPHISM_QSS not in app.styleSheet():
        app.setStyleSheet(app.styleSheet() + GLASSMORPHISM_QSS)


def apply_glass_shadow(widget: QtWidgets.QWidget, *, blur: float = 36.0, offset_y: float = 10.0, alpha: int = 40) -> None:
    """Soft, warm elevation used for glass cards."""

    shadow = QtWidgets.QGraphicsDropShadowEffect(widget)
    shadow.setBlurRadius(blur)
    shadow.setOffset(0, offset_y)
    shadow.setColor(QtGui.QColor(60, 20, 30, alpha))
    widget.setGraphicsEffect(shadow)


def set_status_chip(label: QtWidgets.QLabel, text: str, level: str = "neutral") -> None:
    """Show state as words plus colour so it never relies on colour alone."""

    normalized = level if level in _STATUS_LEVELS else "neutral"
    label.setObjectName("statusChip")
    label.setProperty("status", normalized)
    label.setText(text)
    label.adjustSize()
    label.setAccessibleDescription(f"{normalized} state: {text}")
    style = label.style()
    style.unpolish(label)
    style.polish(label)


class GlassHeader(QtWidgets.QFrame):
    """Compact window header: app name, one-line purpose, badge and status chip."""

    def __init__(
        self,
        title: str,
        subtitle: str = "",
        *,
        badge: str = "",
        parent: QtWidgets.QWidget | None = None,
    ) -> None:
        super().__init__(parent)
        self.setObjectName("glassHeader")
        layout = QtWidgets.QHBoxLayout(self)
        layout.setContentsMargins(18, 10, 14, 10)
        layout.setSpacing(12)

        text_box = QtWidgets.QWidget()
        text_box.setSizePolicy(QtWidgets.QSizePolicy.Policy.Expanding, QtWidgets.QSizePolicy.Policy.Preferred)
        text_col = QtWidgets.QVBoxLayout(text_box)
        text_col.setContentsMargins(0, 0, 0, 0)
        text_col.setSpacing(0)
        left_top = QtCore.Qt.AlignmentFlag.AlignLeft | QtCore.Qt.AlignmentFlag.AlignVCenter
        self.title_label = QtWidgets.QLabel(title)
        self.title_label.setObjectName("glassHeaderTitle")
        self.title_label.setAlignment(left_top)
        text_col.addWidget(self.title_label)
        self.subtitle_label = QtWidgets.QLabel(subtitle)
        self.subtitle_label.setObjectName("glassHeaderSubtitle")
        self.subtitle_label.setAlignment(left_top)
        self.subtitle_label.setWordWrap(True)
        self.subtitle_label.setVisible(bool(subtitle))
        text_col.addWidget(self.subtitle_label)
        layout.addWidget(text_box, stretch=1)

        self.trailing = QtWidgets.QHBoxLayout()
        self.trailing.setSpacing(8)
        layout.addLayout(self.trailing)

        fixed = (QtWidgets.QSizePolicy.Policy.Fixed, QtWidgets.QSizePolicy.Policy.Fixed)
        self.status_chip = QtWidgets.QLabel()
        self.status_chip.setSizePolicy(*fixed)
        set_status_chip(self.status_chip, "Idle")
        self.status_chip.setAccessibleName(f"{title} status")
        layout.addWidget(self.status_chip)

        if badge:
            badge_label = QtWidgets.QLabel(badge)
            badge_label.setObjectName("glassHeaderBadge")
            badge_label.setSizePolicy(*fixed)
            layout.addWidget(badge_label)
        apply_glass_shadow(self, blur=28, offset_y=6, alpha=30)

    def set_status(self, text: str, level: str = "neutral") -> None:
        set_status_chip(self.status_chip, text, level)


def install_glass_shell(
    window: QtWidgets.QMainWindow,
    content: QtWidgets.QWidget,
    *,
    title: str,
    subtitle: str = "",
    badge: str = "RapidPy",
    margins: int = 14,
) -> GlassHeader:
    """Place ``content`` under a glass header on a painted backdrop.

    The backdrop becomes the window's central widget and the original content
    is re-parented beneath the header, so existing widget references stay valid.
    """

    backdrop = GlassBackdrop()
    outer = QtWidgets.QVBoxLayout(backdrop)
    outer.setContentsMargins(margins, margins, margins, margins)
    outer.setSpacing(10)
    header = GlassHeader(title, subtitle, badge=badge)
    outer.addWidget(header)
    content.setParent(backdrop)
    outer.addWidget(content, stretch=1)
    window.setCentralWidget(backdrop)
    return header


def style_glass_plot(plot_widget, *, x_label: str | None = None, y_label: str | None = None) -> None:
    """Give a pyqtgraph PlotWidget a frosted, light background that matches the cards."""

    plot_widget.setBackground(QtGui.QColor(255, 255, 255, 150))
    item = plot_widget.getPlotItem()
    axis_pen = QtGui.QPen(QtGui.QColor(GLASS.muted))
    text_pen = QtGui.QPen(QtGui.QColor(GLASS.ink))
    for name in ("left", "bottom", "right", "top"):
        axis = item.getAxis(name)
        axis.setPen(axis_pen)
        axis.setTextPen(text_pen)
    if x_label is not None:
        item.setLabel("bottom", x_label, color=GLASS.ink)
    if y_label is not None:
        item.setLabel("left", y_label, color=GLASS.ink)
    item.showGrid(x=True, y=True, alpha=0.18)


def fit_workspace_window(
    window: QtWidgets.QWidget,
    preferred: tuple[int, int],
    *,
    screen: QtGui.QScreen | None = None,
    fraction: float = 0.94,
) -> tuple[int, int]:
    """Size a data-dense instrument window to the visible working area.

    The shared bounds guard deliberately keeps generic windows compact; the
    instrument panels need room for their plots and controls, so they opt in
    to this handler instead.  The window may grow to the working area but its
    minimum never exceeds it, so nothing is forced off-screen or clipped.
    """

    screen = screen or window.screen() or QtWidgets.QApplication.primaryScreen()
    if screen is None:
        return window.width(), window.height()
    area = screen.availableGeometry()
    width = max(1, min(int(preferred[0]), int(area.width() * fraction)))
    height = max(1, min(int(preferred[1]), int(area.height() * fraction)))
    minimum = window.minimumSize()
    window.setMinimumSize(min(minimum.width(), area.width()), min(minimum.height(), area.height()))
    window.setMaximumSize(area.width(), area.height())
    if window.isMaximized() or window.isFullScreen():
        return area.width(), area.height()
    window.resize(width, height)
    frame = window.frameGeometry()
    x = max(area.left(), min(frame.left(), area.right() - frame.width() + 1))
    y = max(area.top(), min(frame.top(), area.bottom() - frame.height() + 1))
    window.move(x, y)
    return width, height
