from __future__ import annotations

import csv
import sys
import time
from collections import deque
from datetime import datetime
from pathlib import Path

import pyqtgraph as pg
from PySide6 import QtCore, QtGui, QtWidgets
from serial.tools import list_ports

from .config import AppConfig, load_config, save_config
from .models import MeasurementSample
from .squid_serial import SquidCommunicationError, SquidSerialClient


class AcquisitionWorker(QtCore.QObject):
    sample_ready = QtCore.Signal(object)
    status = QtCore.Signal(str)
    failed = QtCore.Signal(str)
    finished = QtCore.Signal()

    def __init__(
        self,
        client: SquidSerialClient,
        interval_s: float,
        spacing_mode: str,
        parent: QtCore.QObject | None = None,
    ) -> None:
        super().__init__(parent)
        self._client = client
        self._interval_s = max(interval_s, 0.05)
        self._spacing_mode = spacing_mode
        self._running = True

    @QtCore.Slot()
    def run(self) -> None:
        started = time.monotonic()
        wait_s = self._interval_s
        self.status.emit("Acquisition started.")

        while self._running:
            time.sleep(wait_s)
            if not self._running:
                break

            elapsed = time.monotonic() - started
            try:
                x, y, z = self._client.read_xyz_volts()
            except (SquidCommunicationError, OSError) as exc:
                self.failed.emit(str(exc))
                break

            self.sample_ready.emit(
                MeasurementSample(
                    time_s=elapsed,
                    x_volts=x,
                    y_volts=y,
                    z_volts=z,
                )
            )

            if self._spacing_mode == "Log":
                wait_s = max(wait_s * self._interval_s, 0.05)
            else:
                wait_s = self._interval_s

        self.status.emit("Acquisition stopped.")
        self.finished.emit()

    def stop(self) -> None:
        self._running = False


class MainWindow(QtWidgets.QMainWindow):
    MAX_POINTS = 2500

    def __init__(self) -> None:
        super().__init__()
        self.setWindowTitle("RapidPy VRM Logger")
        self.resize(1180, 760)
        self._assets_dir = Path(__file__).resolve().parent.parent / "assets"
        self._icon_png = self._assets_dir / "vrm_icon.png"
        if self._icon_png.exists():
            self.setWindowIcon(QtGui.QIcon(str(self._icon_png)))

        self._config = load_config()
        self._client = SquidSerialClient()
        self._thread: QtCore.QThread | None = None
        self._worker: AcquisitionWorker | None = None
        self._csv_handle = None
        self._csv_writer = None

        self._time = deque(maxlen=self.MAX_POINTS)
        self._x_vals = deque(maxlen=self.MAX_POINTS)
        self._y_vals = deque(maxlen=self.MAX_POINTS)
        self._z_vals = deque(maxlen=self.MAX_POINTS)

        self._build_ui()
        self._apply_style()
        self._load_into_widgets()

    def _build_ui(self) -> None:
        root = QtWidgets.QWidget()
        self.setCentralWidget(root)

        main_layout = QtWidgets.QHBoxLayout(root)
        main_layout.setContentsMargins(16, 16, 16, 16)
        main_layout.setSpacing(14)

        controls_card = QtWidgets.QFrame()
        controls_card.setObjectName("card")
        controls_card.setMinimumWidth(360)
        controls_layout = QtWidgets.QVBoxLayout(controls_card)
        controls_layout.setContentsMargins(18, 18, 18, 18)
        controls_layout.setSpacing(12)

        title = QtWidgets.QLabel("VRM Decay Logger")
        title.setObjectName("title")
        subtitle = QtWidgets.QLabel("Three-axis SQUID live view + CSV logging")
        subtitle.setObjectName("subtitle")
        controls_layout.addWidget(title)
        controls_layout.addWidget(subtitle)

        self.port_combo = QtWidgets.QComboBox()
        self.refresh_btn = QtWidgets.QPushButton("Refresh Ports")
        self.connect_btn = QtWidgets.QPushButton("Connect")
        self.connect_btn.setObjectName("accent")
        self.disconnect_btn = QtWidgets.QPushButton("Disconnect")

        self.interval_spin = QtWidgets.QDoubleSpinBox()
        self.interval_spin.setRange(0.05, 3600.0)
        self.interval_spin.setDecimals(3)
        self.interval_spin.setSuffix(" s")

        self.spacing_combo = QtWidgets.QComboBox()
        self.spacing_combo.addItems(["Linear", "Log"])

        self.unit_combo = QtWidgets.QComboBox()
        self.unit_combo.addItems(["Volts", "Moment"])

        self.cal_x = QtWidgets.QDoubleSpinBox()
        self.cal_y = QtWidgets.QDoubleSpinBox()
        self.cal_z = QtWidgets.QDoubleSpinBox()
        for spin in (self.cal_x, self.cal_y, self.cal_z):
            spin.setRange(-1_000_000.0, 1_000_000.0)
            spin.setDecimals(6)

        self.file_edit = QtWidgets.QLineEdit()
        self.browse_btn = QtWidgets.QPushButton("Choose CSV")

        self.start_btn = QtWidgets.QPushButton("Start Logging")
        self.start_btn.setObjectName("accent")
        self.stop_btn = QtWidgets.QPushButton("Stop")
        self.stop_btn.setEnabled(False)

        form = QtWidgets.QFormLayout()
        form.setLabelAlignment(QtCore.Qt.AlignLeft)
        form.setFormAlignment(QtCore.Qt.AlignTop)
        form.setHorizontalSpacing(10)
        form.setVerticalSpacing(10)

        port_row = QtWidgets.QHBoxLayout()
        port_row.addWidget(self.port_combo, stretch=1)
        port_row.addWidget(self.refresh_btn)
        form.addRow("Serial Port", port_row)

        conn_row = QtWidgets.QHBoxLayout()
        conn_row.addWidget(self.connect_btn)
        conn_row.addWidget(self.disconnect_btn)
        form.addRow("Connection", conn_row)

        form.addRow("Interval", self.interval_spin)
        form.addRow("Step Scale", self.spacing_combo)
        form.addRow("Display Units", self.unit_combo)
        form.addRow("X Calibration", self.cal_x)
        form.addRow("Y Calibration", self.cal_y)
        form.addRow("Z Calibration", self.cal_z)

        file_row = QtWidgets.QHBoxLayout()
        file_row.addWidget(self.file_edit, stretch=1)
        file_row.addWidget(self.browse_btn)
        form.addRow("CSV Output", file_row)

        controls_layout.addLayout(form)

        actions = QtWidgets.QHBoxLayout()
        actions.addWidget(self.start_btn)
        actions.addWidget(self.stop_btn)
        controls_layout.addLayout(actions)

        self.status_label = QtWidgets.QLabel("Idle")
        self.status_label.setObjectName("status")
        controls_layout.addWidget(self.status_label)

        console_title = QtWidgets.QLabel("Console")
        console_title.setObjectName("consoleTitle")
        controls_layout.addWidget(console_title)
        self.console_output = QtWidgets.QPlainTextEdit()
        self.console_output.setReadOnly(True)
        self.console_output.setMinimumHeight(170)
        self.console_output.setObjectName("console")
        controls_layout.addWidget(self.console_output)

        view_card = QtWidgets.QFrame()
        view_card.setObjectName("card")
        view_layout = QtWidgets.QVBoxLayout(view_card)
        view_layout.setContentsMargins(14, 14, 14, 14)
        view_layout.setSpacing(8)

        self.plot = pg.PlotWidget()
        self.plot.setBackground("#f8f3e8")
        self.plot.showGrid(x=True, y=True, alpha=0.2)
        self.plot.setLabel("left", "Signal")
        self.plot.setLabel("bottom", "Time (s)")
        self.plot.getAxis("left").setPen(pg.mkPen("#766865", width=1))
        self.plot.getAxis("left").setTextPen(pg.mkPen("#615250"))
        self.plot.getAxis("bottom").setPen(pg.mkPen("#766865", width=1))
        self.plot.getAxis("bottom").setTextPen(pg.mkPen("#615250"))
        self.plot.addLegend(offset=(8, 8))
        view_box = self.plot.getViewBox()
        view_box.enableAutoRange(axis=pg.ViewBox.XAxis, enable=True)
        view_box.enableAutoRange(axis=pg.ViewBox.YAxis, enable=True)

        self.curve_x = self.plot.plot([], [], pen=pg.mkPen("#7A0219", width=2.2), name="X")
        self.curve_y = self.plot.plot([], [], pen=pg.mkPen("#FFCD34", width=2.2), name="Y")
        self.curve_z = self.plot.plot([], [], pen=pg.mkPen("#3a3231", width=2.2), name="Z")

        value_bar = QtWidgets.QHBoxLayout()
        self.value_x = QtWidgets.QLabel("X: --")
        self.value_y = QtWidgets.QLabel("Y: --")
        self.value_z = QtWidgets.QLabel("Z: --")
        for widget in (self.value_x, self.value_y, self.value_z):
            widget.setObjectName("valuePill")
            value_bar.addWidget(widget)

        view_layout.addWidget(self.plot, stretch=1)
        view_layout.addLayout(value_bar)

        main_layout.addWidget(controls_card)
        main_layout.addWidget(view_card, stretch=1)

        # Subtle card shadow to mimic layered liquid-glass depth.
        self._apply_card_shadow(controls_card)
        self._apply_card_shadow(view_card)

        self.refresh_btn.clicked.connect(self._refresh_ports)
        self.connect_btn.clicked.connect(self._connect)
        self.disconnect_btn.clicked.connect(self._disconnect)
        self.browse_btn.clicked.connect(self._pick_output_file)
        self.start_btn.clicked.connect(self._start_logging)
        self.stop_btn.clicked.connect(self._stop_logging)
        self.unit_combo.currentTextChanged.connect(self._refresh_plot_units)

    def _apply_style(self) -> None:
        app = QtWidgets.QApplication.instance()
        if app is None:
            return

        common_assets = Path(__file__).resolve().parents[2] / "rapidpy_common" / "assets"
        arrow_down = (common_assets / "arrow_down.svg").as_posix()
        arrow_up = (common_assets / "arrow_up.svg").as_posix()

        app.setStyle("Fusion")
        font = QtGui.QFont("SF Pro Text", 10)
        if not QtGui.QFontInfo(font).exactMatch():
            font = QtGui.QFont("Avenir Next", 10)
        if not QtGui.QFontInfo(font).exactMatch():
            font = QtGui.QFont("Segoe UI", 10)
        app.setFont(font)

        style = (
            """
            QWidget {
                background: #f3eee2;
                color: #2f2827;
            }
            QFrame#card {
                background: rgba(255, 255, 255, 0.72);
                border: 1px solid rgba(255, 255, 255, 0.65);
                border-radius: 24px;
            }
            QLabel#title {
                font-size: 24px;
                font-weight: 760;
                color: #7A0219;
            }
            QLabel#subtitle {
                color: #61534d;
                margin-bottom: 4px;
            }
            QLabel#status {
                background: rgba(255, 255, 255, 0.68);
                border: 1px solid rgba(122, 2, 25, 0.18);
                border-radius: 14px;
                padding: 9px;
                color: #4d3a39;
            }
            QLabel#consoleTitle {
                color: #7A0219;
                font-weight: 720;
            }
            QLabel#valuePill {
                background: rgba(255, 255, 255, 0.82);
                border: 1px solid rgba(122, 2, 25, 0.16);
                border-radius: 16px;
                padding: 10px 12px;
                font-weight: 650;
            }
            QPlainTextEdit#console {
                background: rgba(28, 20, 19, 0.88);
                color: #fff2c9;
                border-radius: 14px;
                border: 1px solid rgba(255, 205, 52, 0.32);
                padding: 8px;
                selection-background-color: #7A0219;
            }
            QPushButton {
                background: rgba(255, 255, 255, 0.76);
                border: 1px solid rgba(255, 255, 255, 0.75);
                border-radius: 14px;
                padding: 9px 14px;
            }
            QPushButton:hover {
                background: rgba(255, 255, 255, 0.92);
            }
            QPushButton:pressed {
                background: rgba(232, 226, 216, 0.95);
            }
            QPushButton#accent {
                background: qlineargradient(x1:0, y1:0, x2:1, y2:1, stop:0 #7A0219, stop:1 #5a0013);
                color: #fff9eb;
                border: 1px solid rgba(255, 255, 255, 0.26);
                font-weight: 680;
            }
            QPushButton#accent:hover {
                background: qlineargradient(x1:0, y1:0, x2:1, y2:1, stop:0 #8a0220, stop:1 #650016);
            }
            QPushButton#accent:pressed {
                background: #5a0013;
            }
            QLineEdit, QComboBox, QDoubleSpinBox {
                border: 1px solid rgba(255, 255, 255, 0.82);
                background: rgba(255, 255, 255, 0.72);
                border-radius: 12px;
                padding: 7px;
                selection-background-color: #7A0219;
                selection-color: #ffffff;
            }
            QComboBox {
                padding-right: 34px;
            }
            QComboBox::drop-down {
                subcontrol-origin: padding;
                subcontrol-position: top right;
                width: 28px;
                margin: 3px;
                border: none;
                border-radius: 10px;
                background: rgba(122, 2, 25, 0.12);
            }
            QComboBox::drop-down:hover {
                background: rgba(122, 2, 25, 0.2);
            }
            QComboBox::drop-down:pressed {
                background: rgba(122, 2, 25, 0.28);
            }
            QComboBox::down-arrow {
                image: url(__ARROW_DOWN__);
                width: 14px;
                height: 14px;
            }
            QAbstractSpinBox {
                padding-right: 52px;
            }
            QAbstractSpinBox::up-button,
            QAbstractSpinBox::down-button {
                width: 24px;
                border: none;
                border-radius: 9px;
                background: rgba(122, 2, 25, 0.12);
                margin-right: 3px;
            }
            QAbstractSpinBox::up-button {
                subcontrol-origin: border;
                subcontrol-position: top right;
                margin-top: 3px;
                margin-bottom: 1px;
            }
            QAbstractSpinBox::down-button {
                subcontrol-origin: border;
                subcontrol-position: bottom right;
                margin-top: 1px;
                margin-bottom: 3px;
            }
            QAbstractSpinBox::up-button:hover,
            QAbstractSpinBox::down-button:hover {
                background: rgba(122, 2, 25, 0.2);
            }
            QAbstractSpinBox::up-button:pressed,
            QAbstractSpinBox::down-button:pressed {
                background: rgba(122, 2, 25, 0.28);
            }
            QAbstractSpinBox::up-arrow {
                image: url(__ARROW_UP__);
                width: 13px;
                height: 13px;
            }
            QAbstractSpinBox::down-arrow {
                image: url(__ARROW_DOWN__);
                width: 13px;
                height: 13px;
            }
            QScrollBar:vertical {
                background: rgba(255, 255, 255, 0.2);
                width: 10px;
                border-radius: 5px;
                margin: 4px 2px 4px 2px;
            }
            QScrollBar::handle:vertical {
                background: rgba(255, 205, 52, 0.72);
                border-radius: 5px;
                min-height: 24px;
            }
            QScrollBar::add-line:vertical,
            QScrollBar::sub-line:vertical {
                height: 0px;
            }
            """
        )
        self.setStyleSheet(style.replace("__ARROW_DOWN__", arrow_down).replace("__ARROW_UP__", arrow_up))

    def _apply_card_shadow(self, card: QtWidgets.QFrame) -> None:
        shadow = QtWidgets.QGraphicsDropShadowEffect(self)
        shadow.setBlurRadius(34)
        shadow.setOffset(0, 10)
        shadow.setColor(QtGui.QColor(35, 25, 25, 48))
        card.setGraphicsEffect(shadow)

    def _load_into_widgets(self) -> None:
        self._refresh_ports()
        self.interval_spin.setValue(self._config.interval_s)
        self.spacing_combo.setCurrentText(self._config.spacing_mode)
        self.unit_combo.setCurrentText(self._config.display_unit)
        self.cal_x.setValue(self._config.calibration_x)
        self.cal_y.setValue(self._config.calibration_y)
        self.cal_z.setValue(self._config.calibration_z)
        self.file_edit.setText(self._config.output_file)

        if self._config.window_geometry:
            self.restoreGeometry(QtCore.QByteArray.fromHex(self._config.window_geometry.encode("ascii")))

    def _refresh_ports(self) -> None:
        current = self.port_combo.currentText()
        self.port_combo.clear()
        ports = sorted(p.device for p in list_ports.comports())
        self.port_combo.addItems(ports)

        preferred = self._config.port or current
        if preferred and preferred in ports:
            self.port_combo.setCurrentText(preferred)

    def _set_status(self, message: str) -> None:
        self.status_label.setText(message)
        self._append_console(message)

    def _append_console(self, message: str) -> None:
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.console_output.appendPlainText(f"[{timestamp}] {message}")

    def _connect(self) -> None:
        port = self.port_combo.currentText().strip()
        if not port:
            QtWidgets.QMessageBox.warning(self, "Missing Port", "Choose a serial port first.")
            return

        try:
            self._client.connect(port=port, timeout=1.0)
        except Exception as exc:
            QtWidgets.QMessageBox.critical(self, "Connection Error", str(exc))
            return

        self._set_status(f"Connected to {port} at 1200,N,8,1")

    def _disconnect(self) -> None:
        if self._thread is not None:
            QtWidgets.QMessageBox.information(
                self,
                "Acquisition Running",
                "Stop acquisition before disconnecting.",
            )
            return
        self._client.disconnect()
        self._set_status("Disconnected")

    def _pick_output_file(self) -> None:
        current = self.file_edit.text().strip() or "vrm_log.csv"
        path, _ = QtWidgets.QFileDialog.getSaveFileName(
            self,
            "Choose CSV Output",
            current,
            "CSV Files (*.csv)",
        )
        if path:
            self.file_edit.setText(path)

    def _start_logging(self) -> None:
        if not self._client.is_connected:
            QtWidgets.QMessageBox.warning(self, "Not Connected", "Connect to SQUID serial first.")
            return

        output_path = Path(self.file_edit.text().strip())
        if not output_path.name:
            QtWidgets.QMessageBox.warning(self, "Output Missing", "Choose a CSV output path.")
            return

        try:
            output_path.parent.mkdir(parents=True, exist_ok=True)
            is_new_file = not output_path.exists() or output_path.stat().st_size == 0
            self._csv_handle = output_path.open("a", newline="", encoding="utf-8")
            self._csv_writer = csv.writer(self._csv_handle)
            if is_new_file:
                self._csv_writer.writerow(
                    [
                        "time_s",
                        "x_volts",
                        "y_volts",
                        "z_volts",
                        "x_display",
                        "y_display",
                        "z_display",
                        "display_unit",
                    ]
                )
        except OSError as exc:
            QtWidgets.QMessageBox.critical(self, "File Error", str(exc))
            return

        self._time.clear()
        self._x_vals.clear()
        self._y_vals.clear()
        self._z_vals.clear()

        self.start_btn.setEnabled(False)
        self.stop_btn.setEnabled(True)

        self._thread = QtCore.QThread(self)
        self._worker = AcquisitionWorker(
            client=self._client,
            interval_s=self.interval_spin.value(),
            spacing_mode=self.spacing_combo.currentText(),
        )
        self._worker.moveToThread(self._thread)

        self._thread.started.connect(self._worker.run)
        self._worker.sample_ready.connect(self._handle_sample)
        self._worker.failed.connect(self._handle_worker_error)
        self._worker.status.connect(self._set_status)
        self._worker.finished.connect(self._worker_finished)
        self._worker.finished.connect(self._thread.quit)
        self._worker.finished.connect(self._worker.deleteLater)
        self._thread.finished.connect(self._thread.deleteLater)

        self._thread.start()
        self._append_console("Live plotting and CSV logging active.")

    def _stop_logging(self) -> None:
        if self._worker is not None:
            self._worker.stop()

    def _worker_finished(self) -> None:
        self.start_btn.setEnabled(True)
        self.stop_btn.setEnabled(False)

        if self._csv_handle is not None:
            self._csv_handle.close()
            self._csv_handle = None
            self._csv_writer = None

        self._worker = None
        self._thread = None

    def _display_values(
        self, x: float, y: float, z: float, unit: str
    ) -> tuple[float, float, float]:
        if unit == "Moment":
            x *= self.cal_x.value()
            y *= self.cal_y.value()
            z *= self.cal_z.value()
            suffix = "emu"
        else:
            suffix = "V"

        self.value_x.setText(f"X: {x:+.6g} {suffix}")
        self.value_y.setText(f"Y: {y:+.6g} {suffix}")
        self.value_z.setText(f"Z: {z:+.6g} {suffix}")
        return x, y, z

    @QtCore.Slot(object)
    def _handle_sample(self, sample: MeasurementSample) -> None:
        unit = self.unit_combo.currentText()
        x_disp, y_disp, z_disp = self._display_values(
            sample.x_volts,
            sample.y_volts,
            sample.z_volts,
            unit,
        )

        self._time.append(sample.time_s)
        self._x_vals.append(x_disp)
        self._y_vals.append(y_disp)
        self._z_vals.append(z_disp)

        self.curve_x.setData(list(self._time), list(self._x_vals))
        self.curve_y.setData(list(self._time), list(self._y_vals))
        self.curve_z.setData(list(self._time), list(self._z_vals))
        view_box = self.plot.getViewBox()
        view_box.enableAutoRange(axis=pg.ViewBox.XAxis, enable=True)
        view_box.enableAutoRange(axis=pg.ViewBox.YAxis, enable=True)

        if self._csv_writer is not None:
            self._csv_writer.writerow(
                [
                    f"{sample.time_s:.6f}",
                    f"{sample.x_volts:.9g}",
                    f"{sample.y_volts:.9g}",
                    f"{sample.z_volts:.9g}",
                    f"{x_disp:.9g}",
                    f"{y_disp:.9g}",
                    f"{z_disp:.9g}",
                    unit,
                ]
            )
            self._csv_handle.flush()

    @QtCore.Slot(str)
    def _handle_worker_error(self, message: str) -> None:
        self._set_status(f"Error: {message}")
        QtWidgets.QMessageBox.critical(self, "Acquisition Error", message)
        self._stop_logging()

    def _refresh_plot_units(self) -> None:
        if self.unit_combo.currentText() == "Moment":
            self.plot.setLabel("left", "Moment (emu)")
        else:
            self.plot.setLabel("left", "Voltage (V)")

    def closeEvent(self, event: QtGui.QCloseEvent) -> None:  # noqa: N802
        if self._worker is not None:
            self._worker.stop()
            while self._thread is not None and self._thread.isRunning():
                QtWidgets.QApplication.processEvents()
                time.sleep(0.02)

        if self._csv_handle is not None:
            self._csv_handle.close()

        self._client.disconnect()
        self._save_from_widgets()
        event.accept()

    def _save_from_widgets(self) -> None:
        self._config = AppConfig(
            port=self.port_combo.currentText().strip(),
            interval_s=self.interval_spin.value(),
            spacing_mode=self.spacing_combo.currentText(),
            output_file=self.file_edit.text().strip(),
            display_unit=self.unit_combo.currentText(),
            calibration_x=self.cal_x.value(),
            calibration_y=self.cal_y.value(),
            calibration_z=self.cal_z.value(),
            window_geometry=self.saveGeometry().toHex().data().decode("ascii"),
        )
        save_config(self._config)


def main() -> int:
    app = QtWidgets.QApplication(sys.argv)
    pg.setConfigOptions(antialias=True)
    window = MainWindow()
    window.show()
    return app.exec()
