"""The RAPID app icon, everywhere the main app shows its identity.

One source of truth -- the packaged ``rapid_main_assets`` icon -- feeds the
taskbar/title-bar icon, the toolbar brand, the sidebar footer, the About and
login dialogs, and the startup splash.
"""

from __future__ import annotations

from functools import lru_cache
from pathlib import Path

from PySide6 import QtCore, QtGui, QtWidgets

from rapidpy_common.ui import load_app_icon

from .startup import main_assets_dir, select_main_icon

#: Taskbar identity: Windows groups RapidPyMain windows under this ID and icon.
APP_USER_MODEL_ID = "UMN.IRM.RAPID.Main"
_ARTWORK = ("rapid_main_icon.png", "rapid_main_window_icon.png", "rapid_icon.png")


def icon_name() -> str:
    name, _path = select_main_icon(main_assets_dir())
    return name or "rapid_main_icon.png"


@lru_cache(maxsize=1)
def app_icon() -> QtGui.QIcon:
    return load_app_icon(icon_name(), main_assets_dir())


@lru_cache(maxsize=1)
def _artwork_path() -> Path | None:
    folder = main_assets_dir()
    for name in _ARTWORK:
        if (folder / name).is_file():
            return folder / name
    return None


_ARTWORK_CACHE: dict[str, QtGui.QPixmap] = {}


def _artwork() -> QtGui.QPixmap:
    path = _artwork_path()
    key = str(path)
    if key not in _ARTWORK_CACHE:
        _ARTWORK_CACHE[key] = QtGui.QPixmap(str(path)) if path is not None else QtGui.QPixmap()
    return _ARTWORK_CACHE[key]


def brand_pixmap(size: int, widget: QtWidgets.QWidget | None = None) -> QtGui.QPixmap:
    """The full-resolution artwork smoothly scaled for ``size`` logical pixels."""

    if widget is not None:
        ratio = widget.devicePixelRatioF()
    else:
        screen = QtWidgets.QApplication.primaryScreen() if QtWidgets.QApplication.instance() else None
        ratio = screen.devicePixelRatio() if screen is not None else 1.0
    physical = max(1, int(round(size * ratio)))
    source = _artwork()
    if source.isNull():
        source = app_icon().pixmap(physical, physical)
    if source.isNull():
        return QtGui.QPixmap()
    pixmap = source.scaled(
        physical, physical, QtCore.Qt.AspectRatioMode.KeepAspectRatio, QtCore.Qt.TransformationMode.SmoothTransformation
    )
    pixmap.setDevicePixelRatio(ratio)
    return pixmap


def brand_label(size: int, parent: QtWidgets.QWidget | None = None) -> QtWidgets.QLabel:
    label = QtWidgets.QLabel(parent)
    label.setObjectName("brandIcon")
    label.setPixmap(brand_pixmap(size, parent))
    label.setFixedSize(size, size)
    label.setScaledContents(False)
    label.setAlignment(QtCore.Qt.AlignmentFlag.AlignCenter)
    label.setAccessibleName("RAPID icon")
    label.setAttribute(QtCore.Qt.WidgetAttribute.WA_TranslucentBackground, True)
    return label


def splash_screen() -> QtWidgets.QSplashScreen | None:
    """A glass card with the icon, shown while the main window is built."""

    art = brand_pixmap(176)
    if art.isNull():
        return None
    ratio = art.devicePixelRatio()
    width, height = 380, 300
    canvas = QtGui.QPixmap(int(width * ratio), int(height * ratio))
    canvas.setDevicePixelRatio(ratio)
    canvas.fill(QtCore.Qt.GlobalColor.transparent)
    painter = QtGui.QPainter(canvas)
    painter.setRenderHint(QtGui.QPainter.RenderHint.Antialiasing, True)
    card = QtCore.QRectF(6, 6, width - 12, height - 12)
    gradient = QtGui.QLinearGradient(card.topLeft(), card.bottomRight())
    gradient.setColorAt(0.0, QtGui.QColor(250, 243, 241, 245))
    gradient.setColorAt(1.0, QtGui.QColor(243, 234, 222, 245))
    painter.setPen(QtGui.QPen(QtGui.QColor(0, 0, 0, 40), 1))
    painter.setBrush(gradient)
    painter.drawRoundedRect(card, 18, 18)
    painter.drawPixmap(QtCore.QPointF((width - 176) / 2, 26), art)
    font = QtGui.QFont(painter.font())
    font.setPixelSize(20)
    font.setBold(True)
    painter.setFont(font)
    painter.setPen(QtGui.QColor("#7A0219"))
    painter.drawText(QtCore.QRectF(0, 212, width, 30), QtCore.Qt.AlignmentFlag.AlignCenter, "RAPID v4")
    font.setPixelSize(12)
    font.setBold(False)
    painter.setFont(font)
    painter.setPen(QtGui.QColor(60, 50, 54, 170))
    painter.drawText(QtCore.QRectF(0, 240, width, 22), QtCore.Qt.AlignmentFlag.AlignCenter,
                     "Paleomagnetics Control System · starting…")
    painter.end()
    splash = QtWidgets.QSplashScreen(canvas, QtCore.Qt.WindowType.WindowStaysOnTopHint)
    splash.setAttribute(QtCore.Qt.WidgetAttribute.WA_TranslucentBackground, True)
    splash.setWindowIcon(app_icon())
    return splash
