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
