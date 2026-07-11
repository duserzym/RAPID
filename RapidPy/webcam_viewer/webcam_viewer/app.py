"""
app.py — WebcamWidget and WebcamWindow for RAPID v4 XY-stage monitoring.

Supports:
  • USB webcams via OpenCV cv2.VideoCapture(index)
  • IP cameras via RTSP / HTTP URL (cv2.VideoCapture("rtsp://..."))
  • PySide6 QMediaDevices auto-enumeration for camera discovery
  • Graceful fallback if OpenCV is not installed (static placeholder frame)

Usage:
  from webcam_viewer.app import main
  main()
"""
from __future__ import annotations

import sys
from datetime import datetime
from pathlib import Path
from typing import Optional

from PySide6 import QtCore, QtGui, QtMultimedia, QtWidgets

# Optional OpenCV import — graceful degradation if not installed
try:
    import cv2  # type: ignore
    _OPENCV_AVAILABLE = True
except ImportError:
    _OPENCV_AVAILABLE = False


# ─────────────────────────────────────────────────────────────────────────────
# Helpers
# ─────────────────────────────────────────────────────────────────────────────

def _enumerate_cameras() -> list[tuple[str, str]]:
    """
    Return (display_name, identifier) pairs for available cameras.
    Uses QMediaDevices for system cameras; adds IP camera option.
    """
    entries: list[tuple[str, str]] = []
    for idx, dev in enumerate(QtMultimedia.QMediaDevices.videoInputs()):
        entries.append((dev.description(), str(idx)))
    if _OPENCV_AVAILABLE:
        # Probe a few additional USB indices not returned by QMediaDevices
        for extra_idx in range(max(len(entries), 0), max(len(entries), 0) + 4):
            cap = cv2.VideoCapture(extra_idx, cv2.CAP_DSHOW)
            if cap.isOpened():
                cap.release()
                # Only add if not already listed
                if str(extra_idx) not in [e[1] for e in entries]:
                    entries.append((f"USB Camera {extra_idx}", str(extra_idx)))
    if not entries:
        entries.append(("Default Camera (0)", "0"))
    entries.append(("IP / RTSP Camera…", "ip"))
    return entries


def _make_placeholder(width: int = 640, height: int = 480, msg: str = "") -> QtGui.QPixmap:
    """Return a dark placeholder pixmap with an optional message."""
    pix = QtGui.QPixmap(width, height)
    pix.fill(QtGui.QColor("#1a1a1a"))
    painter = QtGui.QPainter(pix)
    painter.setPen(QtGui.QPen(QtGui.QColor("#888888")))
    painter.setFont(QtGui.QFont("Segoe UI", 14))
    painter.drawText(pix.rect(), QtCore.Qt.AlignmentFlag.AlignCenter, msg or "No Signal")
    painter.end()
    return pix


# ─────────────────────────────────────────────────────────────────────────────
# WebcamWidget
# ─────────────────────────────────────────────────────────────────────────────

class WebcamWidget(QtWidgets.QWidget):
    """
    Self-contained webcam display widget.

    Embeds camera controls (selector, FPS, snapshot) and a live video feed.
    Works standalone or embedded in any QWidget/QDialog.
    """

    #: Emitted when connection status changes: True = connected, False = failed
    connection_changed = QtCore.Signal(bool)

    def __init__(self, parent: Optional[QtWidgets.QWidget] = None) -> None:
        super().__init__(parent)
        self._cap: Optional["cv2.VideoCapture"] = None  # type: ignore[name-defined]
        self._timer = QtCore.QTimer(self)
        self._timer.timeout.connect(self._grab_frame)
        self._snapshot_dir = Path.home() / "RAPID_snapshots"
        self._connected = False

        self._build_ui()
        self._populate_cameras()

    # ── UI Construction ───────────────────────────────────────────────────────

    def _build_ui(self) -> None:
        root = QtWidgets.QVBoxLayout(self)
        root.setContentsMargins(8, 8, 8, 8)
        root.setSpacing(6)

        # ── Controls row ──────────────────────────────────────────────────────
        ctrl = QtWidgets.QHBoxLayout()

        self._cam_combo = QtWidgets.QComboBox()
        self._cam_combo.setMinimumWidth(200)
        self._cam_combo.currentIndexChanged.connect(self._on_camera_selected)
        ctrl.addWidget(QtWidgets.QLabel("Camera:"))
        ctrl.addWidget(self._cam_combo, 2)

        self._url_edit = QtWidgets.QLineEdit()
        self._url_edit.setPlaceholderText("rtsp://user:pass@host/stream")
        self._url_edit.setVisible(False)
        ctrl.addWidget(self._url_edit, 3)

        ctrl.addSpacing(12)
        ctrl.addWidget(QtWidgets.QLabel("FPS:"))
        self._fps_combo = QtWidgets.QComboBox()
        for fps in ("1", "5", "10", "15", "30"):
            self._fps_combo.addItem(fps)
        self._fps_combo.setCurrentText("10")
        self._fps_combo.currentTextChanged.connect(self._on_fps_changed)
        ctrl.addWidget(self._fps_combo)

        ctrl.addSpacing(12)
        self._connect_btn = QtWidgets.QPushButton("Connect")
        self._connect_btn.setFixedWidth(90)
        self._connect_btn.clicked.connect(self._toggle_connection)
        ctrl.addWidget(self._connect_btn)

        self._snapshot_btn = QtWidgets.QPushButton("Snapshot")
        self._snapshot_btn.setFixedWidth(90)
        self._snapshot_btn.setEnabled(False)
        self._snapshot_btn.clicked.connect(self._take_snapshot)
        ctrl.addWidget(self._snapshot_btn)

        ctrl.addStretch()
        root.addLayout(ctrl)

        # ── Video display ─────────────────────────────────────────────────────
        self._video_lbl = QtWidgets.QLabel()
        self._video_lbl.setAlignment(QtCore.Qt.AlignmentFlag.AlignCenter)
        self._video_lbl.setMinimumSize(320, 240)
        self._video_lbl.setStyleSheet(
            "background: #111; border: 1px solid rgba(122,2,25,0.2); border-radius: 4px;"
        )
        self._video_lbl.setPixmap(_make_placeholder())
        self._video_lbl.setSizePolicy(
            QtWidgets.QSizePolicy.Policy.Expanding,
            QtWidgets.QSizePolicy.Policy.Expanding,
        )
        root.addWidget(self._video_lbl, 1)

        # ── Status bar ────────────────────────────────────────────────────────
        status_row = QtWidgets.QHBoxLayout()
        self._status_led = QtWidgets.QLabel("●")
        self._status_led.setStyleSheet("color: #cc3333; font-size: 14px;")
        self._status_label = QtWidgets.QLabel("Disconnected")
        self._status_label.setStyleSheet("color: #666; font-size: 12px;")

        self._snap_dir_btn = QtWidgets.QPushButton("📁 Snapshot Dir")
        self._snap_dir_btn.setFixedHeight(22)
        self._snap_dir_btn.clicked.connect(self._choose_snapshot_dir)
        self._snap_dir_lbl = QtWidgets.QLabel(str(self._snapshot_dir))
        self._snap_dir_lbl.setStyleSheet("color: #888; font-size: 11px;")

        status_row.addWidget(self._status_led)
        status_row.addWidget(self._status_label)
        status_row.addStretch()
        status_row.addWidget(self._snap_dir_btn)
        status_row.addWidget(self._snap_dir_lbl)
        root.addLayout(status_row)

    def _populate_cameras(self) -> None:
        self._cam_combo.blockSignals(True)
        self._cam_combo.clear()
        for name, identifier in _enumerate_cameras():
            self._cam_combo.addItem(name, identifier)
        self._cam_combo.blockSignals(False)

    # ── Camera control ────────────────────────────────────────────────────────

    def _on_camera_selected(self, index: int) -> None:
        identifier = self._cam_combo.itemData(index)
        self._url_edit.setVisible(identifier == "ip")
        if self._connected:
            self._disconnect()

    def _on_fps_changed(self, fps_str: str) -> None:
        if self._timer.isActive():
            interval = max(33, int(1000 / max(1, int(fps_str))))
            self._timer.start(interval)

    def _toggle_connection(self) -> None:
        if self._connected:
            self._disconnect()
        else:
            self._connect()

    def _connect(self) -> None:
        if not _OPENCV_AVAILABLE:
            self._set_status(False, "OpenCV not installed — install opencv-python")
            return

        identifier = self._cam_combo.currentData()
        if identifier == "ip":
            source: str | int = self._url_edit.text().strip()
            if not source:
                self._set_status(False, "Enter an RTSP/HTTP URL")
                return
        else:
            try:
                source = int(identifier)
            except (TypeError, ValueError):
                source = 0

        self._cap = cv2.VideoCapture(source, cv2.CAP_DSHOW if isinstance(source, int) else cv2.CAP_ANY)  # type: ignore[attr-defined]
        if not self._cap.isOpened():
            self._cap = None
            self._set_status(False, "Failed to open camera")
            return

        fps = int(self._fps_combo.currentText())
        interval = max(33, int(1000 / max(1, fps)))
        self._timer.start(interval)
        self._set_status(True, f"Connected — {self._cam_combo.currentText()}")
        self._snapshot_btn.setEnabled(True)
        self._connect_btn.setText("Disconnect")

    def _disconnect(self) -> None:
        self._timer.stop()
        if self._cap is not None:
            self._cap.release()
            self._cap = None
        self._video_lbl.setPixmap(_make_placeholder(msg="Disconnected"))
        self._set_status(False, "Disconnected")
        self._snapshot_btn.setEnabled(False)
        self._connect_btn.setText("Connect")

    def _set_status(self, ok: bool, msg: str) -> None:
        self._connected = ok
        color = "#22bb33" if ok else "#cc3333"
        self._status_led.setStyleSheet(f"color: {color}; font-size: 14px;")
        self._status_label.setText(msg)
        self.connection_changed.emit(ok)

    # ── Frame capture ─────────────────────────────────────────────────────────

    def _grab_frame(self) -> None:
        if self._cap is None or not self._cap.isOpened():
            self._set_status(False, "Camera disconnected")
            self._timer.stop()
            return
        ret, frame = self._cap.read()
        if not ret:
            return
        self._display_frame(frame)

    def _display_frame(self, frame: "cv2.Mat") -> None:  # type: ignore[name-defined]
        import cv2
        rgb = cv2.cvtColor(frame, cv2.COLOR_BGR2RGB)
        h, w, ch = rgb.shape
        qimg = QtGui.QImage(rgb.data, w, h, ch * w, QtGui.QImage.Format.Format_RGB888)
        pix = QtGui.QPixmap.fromImage(qimg)
        scaled = pix.scaled(
            self._video_lbl.size(),
            QtCore.Qt.AspectRatioMode.KeepAspectRatio,
            QtCore.Qt.TransformationMode.SmoothTransformation,
        )
        self._video_lbl.setPixmap(scaled)

    # ── Snapshot ──────────────────────────────────────────────────────────────

    def _take_snapshot(self) -> None:
        if self._cap is None or not self._cap.isOpened():
            return
        import cv2
        ret, frame = self._cap.read()
        if not ret:
            return
        self._snapshot_dir.mkdir(parents=True, exist_ok=True)
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        out_path = self._snapshot_dir / f"rapid_snapshot_{ts}.png"
        cv2.imwrite(str(out_path), frame)
        self._status_label.setText(f"Snapshot saved: {out_path.name}")

    def _choose_snapshot_dir(self) -> None:
        d = QtWidgets.QFileDialog.getExistingDirectory(
            self, "Choose Snapshot Directory", str(self._snapshot_dir)
        )
        if d:
            self._snapshot_dir = Path(d)
            self._snap_dir_lbl.setText(str(self._snapshot_dir))

    # ── Cleanup ───────────────────────────────────────────────────────────────

    def closeEvent(self, event: QtGui.QCloseEvent) -> None:  # type: ignore[override]
        self._disconnect()
        super().closeEvent(event)


# ─────────────────────────────────────────────────────────────────────────────
# WebcamWindow (standalone app window)
# ─────────────────────────────────────────────────────────────────────────────

class WebcamWindow(QtWidgets.QMainWindow):
    """Top-level window wrapping WebcamWidget for standalone use."""

    def __init__(self) -> None:
        super().__init__()
        self.setWindowTitle("RAPID — XY Stage Webcam Monitor")
        self.resize(900, 620)
        self._widget = WebcamWidget()
        self.setCentralWidget(self._widget)
        self._widget.connection_changed.connect(self._on_connection_changed)

    def _on_connection_changed(self, connected: bool) -> None:
        status = "Connected" if connected else "Disconnected"
        self.setWindowTitle(f"RAPID — XY Stage Webcam Monitor  [{status}]")


# ─────────────────────────────────────────────────────────────────────────────
# Entry point
# ─────────────────────────────────────────────────────────────────────────────

def main() -> None:
    """Launch the standalone webcam viewer application."""
    # Import theme and icon helpers if available, and keep fallback styling
    # intact if shared resources are unavailable.
    app = QtWidgets.QApplication(sys.argv)
    assets_dir = Path(__file__).resolve().parent / "assets"
    _has_window_icon = False
    try:
        sys.path.insert(0, str(Path(__file__).resolve().parents[2]))
        from rapidpy_common.ui import apply_liquid_glass_theme, set_app_icon
        apply_liquid_glass_theme(app)
        set_app_icon(app, "webcam_viewer_icon.png", assets_dir)
        _has_window_icon = True
    except Exception:
        pass

    win = WebcamWindow()
    if _has_window_icon:
        set_app_icon(win, "webcam_viewer_icon.png", assets_dir)
    win.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
