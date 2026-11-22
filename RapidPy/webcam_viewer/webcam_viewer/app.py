"""
app.py — WebcamWidget and WebcamWindow for RAPID v4 XY-stage monitoring.

USB capture and network playback use the bundled Qt Multimedia backend.

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

# ─────────────────────────────────────────────────────────────────────────────
# Helpers
# ─────────────────────────────────────────────────────────────────────────────

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
        self._devices = QtMultimedia.QMediaDevices(self)
        self._camera = None
        self._session = None
        self._player = None
        self._sink = None
        self._image = QtGui.QImage()
        self._frame_time = QtCore.QElapsedTimer()
        self._frame_deadline = QtCore.QTimer(self)
        self._frame_deadline.setSingleShot(True)
        self._frame_deadline.setInterval(15_000)
        self._frame_deadline.timeout.connect(self._frame_timed_out)
        self._snapshot_dir = Path.home() / "RAPID_snapshots"
        self._connected = False

        self._build_ui()
        self._devices.videoInputsChanged.connect(self._populate_cameras)
        self._populate_cameras()

    # ── UI Construction ───────────────────────────────────────────────────────

    def _build_ui(self) -> None:
        root = QtWidgets.QVBoxLayout(self)
        root.setContentsMargins(8, 8, 8, 8)
        root.setSpacing(6)

        source_controls = QtWidgets.QHBoxLayout()

        self._cam_combo = QtWidgets.QComboBox()
        self._cam_combo.setMinimumWidth(160)
        self._cam_combo.currentIndexChanged.connect(self._on_camera_selected)
        source_controls.addWidget(QtWidgets.QLabel("Camera:"))
        source_controls.addWidget(self._cam_combo, 2)

        self._url_edit = QtWidgets.QLineEdit()
        self._url_edit.setPlaceholderText("rtsp://camera.local/stream")
        self._url_edit.setVisible(False)
        source_controls.addWidget(self._url_edit, 3)
        root.addLayout(source_controls)

        ctrl = QtWidgets.QHBoxLayout()
        ctrl.addWidget(QtWidgets.QLabel("FPS:"))
        self._fps_combo = QtWidgets.QComboBox()
        for fps in ("1", "5", "10", "15", "30"):
            self._fps_combo.addItem(fps)
        self._fps_combo.setCurrentText("10")
        self._fps_combo.setToolTip("Maximum preview refresh rate")
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
            QtWidgets.QSizePolicy.Policy.Ignored,
            QtWidgets.QSizePolicy.Policy.Ignored,
        )
        self._video_lbl.installEventFilter(self)
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
        self._snap_dir_lbl.setWordWrap(True)
        self._snap_dir_lbl.setStyleSheet("color: #888; font-size: 11px;")

        status_row.addWidget(self._status_led)
        status_row.addWidget(self._status_label)
        status_row.addStretch()
        status_row.addWidget(self._snap_dir_btn)
        status_row.addWidget(self._snap_dir_lbl)
        root.addLayout(status_row)

    def _populate_cameras(self) -> None:
        selected = self._cam_combo.currentData()
        devices = self._devices.videoInputs()
        self._cam_combo.blockSignals(True)
        self._cam_combo.clear()
        match = -1
        for index, device in enumerate(devices):
            self._cam_combo.addItem(device.description(), device)
            if isinstance(selected, QtMultimedia.QCameraDevice) and device.id() == selected.id():
                match = index
        if not devices:
            self._cam_combo.addItem("No USB camera detected")
        self._cam_combo.addItem("IP / RTSP Camera…", "ip")
        if selected == "ip":
            match = self._cam_combo.count() - 1
        self._cam_combo.setCurrentIndex(max(0, match))
        self._cam_combo.blockSignals(False)
        self._url_edit.setVisible(selected == "ip")
        if self._camera is not None and match < 0:
            self.disconnect_camera("USB camera disconnected")

    # ── Camera control ────────────────────────────────────────────────────────

    def _on_camera_selected(self, index: int) -> None:
        identifier = self._cam_combo.itemData(index)
        self._url_edit.setVisible(identifier == "ip")
        if self._sink is not None:
            self.disconnect_camera()

    def _toggle_connection(self) -> None:
        if self._sink is not None:
            self.disconnect_camera()
        else:
            self._connect()

    def _connect(self) -> None:
        source = self._cam_combo.currentData()
        if source == "ip":
            url = QtCore.QUrl(self._url_edit.text().strip(), QtCore.QUrl.ParsingMode.StrictMode)
            if not url.isValid() or url.scheme() not in {"rtsp", "http", "https"} or not url.host():
                self._set_status(False, "Enter a valid RTSP or HTTP camera URL")
                return
            self._open_stream(url)
        elif isinstance(source, QtMultimedia.QCameraDevice):
            self._open_camera(source)
        else:
            self._set_status(False, "Select a USB or network camera")

    def _begin_connection(self):
        if self._sink is not None:
            self.disconnect_camera()
        sink = self._sink = QtMultimedia.QVideoSink(self)
        sink.videoFrameChanged.connect(self._receive_frame, QtCore.Qt.ConnectionType.QueuedConnection)
        self._frame_deadline.start()
        self._set_status(False, "Waiting for video…")
        self._connect_btn.setText("Disconnect")
        return sink

    def _open_camera(self, device):
        sink = self._begin_connection()
        camera = self._camera = QtMultimedia.QCamera(device, self)
        session = self._session = QtMultimedia.QMediaCaptureSession(self)
        session.setCamera(camera)
        session.setVideoSink(sink)
        camera.errorOccurred.connect(
            lambda error, _detail: self._source_ended(sink, "USB camera could not provide video")
            if error != QtMultimedia.QCamera.Error.NoError else None
        )
        camera.activeChanged.connect(
            lambda active: self._source_ended(sink, "USB camera stopped")
            if not active and self._connected else None
        )
        camera.start()

    def _open_stream(self, url):
        sink = self._begin_connection()
        player = self._player = QtMultimedia.QMediaPlayer(self)
        player.setVideoSink(sink)
        player.errorOccurred.connect(
            lambda error, _detail: self._source_ended(sink, "Camera stream could not be played")
            if error != QtMultimedia.QMediaPlayer.Error.NoError else None
        )
        player.mediaStatusChanged.connect(
            lambda status: self._source_ended(sink, "Camera stream ended")
            if status == QtMultimedia.QMediaPlayer.MediaStatus.EndOfMedia else None
        )
        player.setSource(url)
        if self._player is player:
            player.play()

    def _source_ended(self, sink, message):
        if self._sink is sink:
            self.disconnect_camera(message)

    def _frame_timed_out(self):
        self.disconnect_camera("No video frames received")

    def disconnect_camera(self, message="Disconnected") -> None:
        self._frame_deadline.stop()
        sink, self._sink = self._sink, None
        camera, self._camera = self._camera, None
        session, self._session = self._session, None
        player, self._player = self._player, None
        if session is not None:
            session.setVideoSink(None)
            session.setCamera(None)
            session.deleteLater()
        if camera is not None:
            camera.stop()
            camera.deleteLater()
        if player is not None:
            player.setVideoSink(None)
            player.stop()
            player.setSource(QtCore.QUrl())
            player.deleteLater()
        if sink is not None:
            sink.deleteLater()
        self._image = QtGui.QImage()
        self._frame_time.invalidate()
        self._video_lbl.setPixmap(_make_placeholder(msg=message))
        self._set_status(False, message)
        self._snapshot_btn.setEnabled(False)
        self._connect_btn.setText("Connect")

    def _set_status(self, ok: bool, msg: str) -> None:
        self._connected = ok
        color = "#22bb33" if ok else "#cc3333"
        self._status_led.setStyleSheet(f"color: {color}; font-size: 14px;")
        self._status_label.setText(msg)
        self.connection_changed.emit(ok)

    # ── Frame capture ─────────────────────────────────────────────────────────

    @QtCore.Slot(QtMultimedia.QVideoFrame)
    def _receive_frame(self, frame) -> None:
        if self._sink is None or self.sender() is not self._sink or not frame.isValid():
            return
        if self._frame_time.isValid() and self._frame_time.elapsed() < 1000 / int(self._fps_combo.currentText()):
            return
        image = frame.toImage()
        if image.isNull():
            return
        rotation = frame.rotation().value
        if rotation:
            image = image.transformed(QtGui.QTransform().rotate(rotation))
        if frame.mirrored():
            image = image.mirrored(True, False)
        self._image = image.copy()
        self._frame_time.start()
        self._frame_deadline.start()
        self._render_frame()
        if not self._connected:
            self._set_status(True, "Video connected")
            self._snapshot_btn.setEnabled(True)

    def _render_frame(self):
        if self._image.isNull():
            return
        scaled = QtGui.QPixmap.fromImage(self._image).scaled(
            self._video_lbl.contentsRect().size(),
            QtCore.Qt.AspectRatioMode.KeepAspectRatio,
            QtCore.Qt.TransformationMode.SmoothTransformation,
        )
        self._video_lbl.setPixmap(scaled)

    def eventFilter(self, watched, event):
        if watched is self._video_lbl and event.type() == QtCore.QEvent.Type.Resize:
            self._render_frame()
        return super().eventFilter(watched, event)

    # ── Snapshot ──────────────────────────────────────────────────────────────

    def _take_snapshot(self) -> None:
        if self._image.isNull() or not self._connected:
            return
        ts = datetime.now().strftime("%Y%m%d_%H%M%S_%f")
        out_path = self._snapshot_dir / f"rapid_snapshot_{ts}.png"
        output = QtCore.QSaveFile(str(out_path))
        try:
            self._snapshot_dir.mkdir(parents=True, exist_ok=True)
            if not output.open(QtCore.QIODevice.OpenModeFlag.WriteOnly):
                raise OSError(output.errorString())
            if not self._image.save(output, "PNG") or not output.commit():
                raise OSError(output.errorString())
        except OSError:
            if output.isOpen():
                output.cancelWriting()
            self._status_label.setText("Could not save snapshot")
        else:
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
        self.disconnect_camera()
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

    def closeEvent(self, event: QtGui.QCloseEvent) -> None:
        self._widget.disconnect_camera()
        super().closeEvent(event)


# ─────────────────────────────────────────────────────────────────────────────
# Entry point
# ─────────────────────────────────────────────────────────────────────────────

def main() -> None:
    """Launch the standalone webcam viewer application."""
    # Import shared theme/icon helpers and keep taskbar icon assignment explicit.
    sys.path.insert(0, str(Path(__file__).resolve().parents[2]))
    from rapidpy_common.resources import asset_directory
    from rapidpy_common.ui import apply_liquid_glass_theme, apply_window_bounds_guard, set_app_icon

    app = QtWidgets.QApplication(sys.argv)
    apply_window_bounds_guard(app)
    assets_dir = asset_directory(
        "webcam_viewer_assets",
        Path(__file__).resolve().parent.parent / "assets",
    )
    apply_liquid_glass_theme(app)
    set_app_icon(app, "webcam_viewer_icon.png", assets_dir)

    win = WebcamWindow()
    set_app_icon(win, "webcam_viewer_icon.png", assets_dir)
    win.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
