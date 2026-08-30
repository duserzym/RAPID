"""
webcam_dialog.py — Modeless webcam monitor dialog for rapid_main.

Wraps WebcamWidget in a floating QDialog that stays open while the
user operates the magnetometer.  Cached like DebugConsoleDialog so
re-opening restores the last state.
"""
from __future__ import annotations

import sys
from pathlib import Path
from typing import Optional

from PySide6 import QtCore, QtWidgets
from rapidpy_common.ui import clamp_window_geometry

# Import shared WebcamWidget from the sibling webcam_viewer package.
# Falls back gracefully if webcam_viewer is not on the path.
try:
    _wv_path = Path(__file__).resolve().parents[3] / "webcam_viewer"
    if str(_wv_path) not in sys.path:
        sys.path.insert(0, str(_wv_path))
    from webcam_viewer.app import WebcamWidget  # type: ignore[import]
    _WEBCAM_AVAILABLE = True
except ImportError:
    _WEBCAM_AVAILABLE = False


class WebcamDialog(QtWidgets.QDialog):
    """
    Floating, modeless dialog showing the WebcamWidget live feed.

    Stays alive between hide/show calls (cache with ``_webcam_dlg`` in MainWindow).
    """

    def __init__(self, parent: Optional[QtWidgets.QWidget] = None) -> None:
        super().__init__(parent)
        self.setWindowTitle("Webcam Monitor — XY Stage")
        self.setWindowFlags(
            QtCore.Qt.WindowType.Window
            | QtCore.Qt.WindowType.WindowCloseButtonHint
            | QtCore.Qt.WindowType.WindowMinimizeButtonHint
        )
        self.resize(900, 620)
        self._build_ui()

    def showEvent(self, event: QtCore.QShowEvent) -> None:  # type: ignore[override]
        super().showEvent(event)
        QtCore.QTimer.singleShot(0, self._fit_to_screen)
        handle = self.windowHandle()
        if handle is not None and not getattr(self, "_screen_signal_connected", False):
            if handle.screen() is not None:
                handle.screen().availableGeometryChanged.connect(self._fit_to_screen)
            handle.screenChanged.connect(self._fit_to_screen)
            self._screen_signal_connected = True

    def _fit_to_screen(self, screen: QtCore.QObject | None = None) -> None:
        active_screen = (
            screen
            if isinstance(screen, QtCore.QScreen)
            else (self.screen() or QtWidgets.QApplication.primaryScreen())
        )
        if active_screen is None:
            return
        available = active_screen.availableGeometry()
        max_w, max_h = clamp_window_geometry(available, (self.width(), self.height()))
        min_size = self.minimumSize()
        if min_size.isValid() and not min_size.isNull():
            self.setMinimumSize(min(min_size.width(), max_w), min(min_size.height(), max_h))
        self.resize(min(self.width(), max_w), min(self.height(), max_h))
        frame = self.frameGeometry()
        frame.setSize(
            QtCore.QSize(
                min(frame.width(), max_w),
                min(frame.height(), max_h),
            )
        )
        if frame.width() > available.width() or frame.height() > available.height():
            frame.moveCenter(available.center())
        else:
            new_x = max(
                available.left(),
                min(frame.left(), available.right() - frame.width() + 1),
            )
            new_y = max(
                available.top(),
                min(frame.top(), available.bottom() - frame.height() + 1),
            )
            frame.moveTopLeft(QtCore.QPoint(new_x, new_y))
        self.setGeometry(frame)

    def _build_ui(self) -> None:
        layout = QtWidgets.QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)

        if _WEBCAM_AVAILABLE:
            self._webcam = WebcamWidget(self)
            layout.addWidget(self._webcam)
        else:
            msg = QtWidgets.QLabel(
                "webcam_viewer module not found.\n\n"
                "Ensure the webcam_viewer package is in your Python path:\n"
                "  RapidPy/webcam_viewer/\n\n"
                "Also install OpenCV:\n"
                "  pip install opencv-python"
            )
            msg.setAlignment(QtCore.Qt.AlignmentFlag.AlignCenter)
            msg.setStyleSheet("color: #888; font-size: 13px; padding: 40px;")
            layout.addWidget(msg)

    def closeEvent(self, event: QtWidgets.QCloseEvent) -> None:  # type: ignore[override]
        """Hide instead of destroy so camera connection is preserved."""
        event.ignore()
        self.hide()
