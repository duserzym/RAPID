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

from .config import AppConfig, _auto_find_ini, load_config, read_calibration_from_ini, save_config
from .models import MeasurementSample
from .squid_serial import SquidCommunicationError, SquidSerialClient


class AbsoluteTimeAxis(pg.AxisItem):
    """Top x-axis that displays wall-clock datetime from elapsed-second x values."""

    def __init__(self) -> None:
        super().__init__(orientation="top")
        self._session_start: float = 0.0
        tick_font = QtGui.QFont()
        tick_font.setPointSize(8)
        self.setStyle(tickFont=tick_font, autoExpandTextSpace=False)

    def set_session_start(self, epoch: float) -> None:
        self._session_start = epoch
        self.picture = None  # force redraw
        self.update()

    def tickStrings(self, values: list, scale: float, spacing: float) -> list[str]:
        if not self._session_start:
            return [""] * len(values)
        try:
            span = abs(self.range[1] - self.range[0])
        except Exception:
            span = spacing * max(len(values), 1)
        out = []
        for v in values:
            dt = datetime.fromtimestamp(self._session_start + v)
            if span < 120:          # < 2 min  → HH:MM:SS
                out.append(dt.strftime("%H:%M:%S"))
            elif span < 7_200:      # < 2 h    → HH:MM
                out.append(dt.strftime("%H:%M"))
            elif span < 172_800:    # < 2 days → HH:MM\nMon DD
                out.append(dt.strftime("%H:%M\n%b %d"))
            else:                   # ≥ 2 days → Mon DD\nYYYY
                out.append(dt.strftime("%b %d\n%Y"))
        return out


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
        compact_font = QtGui.QFont(self.font())
        compact_size = compact_font.pointSizeF()
        if compact_size > 0:
            compact_font.setPointSizeF(max(8.0, compact_size - 1.5))
            self.setFont(compact_font)
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

        self._baseline_raw: tuple[float, float, float] | None = None
        self._session_start_epoch: float = 0.0

        self._time = deque(maxlen=self.MAX_POINTS)
        self._x_vals = deque(maxlen=self.MAX_POINTS)
        self._y_vals = deque(maxlen=self.MAX_POINTS)
        self._z_vals = deque(maxlen=self.MAX_POINTS)

        self._build_ui()
        self._apply_style()
        self._load_into_widgets()

        # resize() must come after _build_ui() so it isn't overridden by
        # the window's computed minimum size hint during the first show.
        self.setMinimumSize(1160, 760)
        self.resize(1320, 864)

    def _build_ui(self) -> None:
        root = QtWidgets.QWidget()
        self.setCentralWidget(root)

        outer_layout = QtWidgets.QHBoxLayout(root)
        outer_layout.setContentsMargins(10, 10, 10, 10)
        outer_layout.setSpacing(0)

        splitter = QtWidgets.QSplitter(QtCore.Qt.Orientation.Horizontal)
        splitter.setChildrenCollapsible(False)
        splitter.setHandleWidth(8)
        outer_layout.addWidget(splitter)

        left_scroll = QtWidgets.QScrollArea()
        left_scroll.setObjectName("panelScroll")
        left_scroll.setWidgetResizable(True)
        left_scroll.setFrameShape(QtWidgets.QFrame.Shape.NoFrame)
        left_scroll.setHorizontalScrollBarPolicy(QtCore.Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        left_scroll.setVerticalScrollBarPolicy(QtCore.Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        # Cap the minimum height so this scroll area does not force the window
        # to be as tall as all its content. The form will scroll instead.
        left_scroll.setMinimumHeight(100)
        left_scroll.setSizePolicy(
            QtWidgets.QSizePolicy.Policy.Preferred,
            QtWidgets.QSizePolicy.Policy.Expanding,
        )

        left_host = QtWidgets.QWidget()
        left_host_layout = QtWidgets.QVBoxLayout(left_host)
        left_host_layout.setContentsMargins(0, 0, 6, 0)
        left_host_layout.setSpacing(0)

        controls_card = QtWidgets.QFrame()
        controls_card.setObjectName("card")
        controls_card.setMinimumWidth(260)
        controls_card.setSizePolicy(
            QtWidgets.QSizePolicy.Policy.Preferred,
            QtWidgets.QSizePolicy.Policy.Expanding,
        )
        controls_layout = QtWidgets.QVBoxLayout(controls_card)
        controls_layout.setContentsMargins(12, 12, 12, 12)
        controls_layout.setSpacing(6)

        title = QtWidgets.QLabel("VRM Decay Logger")
        title.setObjectName("title")
        subtitle = QtWidgets.QLabel("Three-axis SQUID live view + CSV logging")
        subtitle.setObjectName("subtitle")
        subtitle.setWordWrap(True)
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

        self.range_fact_spin = QtWidgets.QDoubleSpinBox()
        self.range_fact_spin.setRange(1e-12, 1.0)
        self.range_fact_spin.setDecimals(8)
        self.range_fact_spin.setSingleStep(1e-6)
        self.ini_path_edit = QtWidgets.QLineEdit()
        self.ini_path_edit.setPlaceholderText("Path to Paleomag_v3.INI")
        self.load_ini_btn = QtWidgets.QPushButton("Load from INI")

        self.baseline_btn = QtWidgets.QPushButton("Take Baseline")
        self.baseline_btn.setObjectName("accent")
        self.baseline_x_label = QtWidgets.QLabel("X: --")
        self.baseline_y_label = QtWidgets.QLabel("Y: --")
        self.baseline_z_label = QtWidgets.QLabel("Z: --")

        self.file_edit = QtWidgets.QLineEdit()
        self.browse_btn = QtWidgets.QPushButton("Choose CSV")

        self.start_btn = QtWidgets.QPushButton("Start Logging")
        self.start_btn.setObjectName("accent")
        self.stop_btn = QtWidgets.QPushButton("Stop")
        self.stop_btn.setEnabled(False)

        form = QtWidgets.QFormLayout()
        form.setLabelAlignment(QtCore.Qt.AlignLeft)
        form.setFormAlignment(QtCore.Qt.AlignTop)
        form.setHorizontalSpacing(8)
        form.setVerticalSpacing(5)

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
        form.addRow("Range Factor", self.range_fact_spin)

        ini_row = QtWidgets.QHBoxLayout()
        ini_row.addWidget(self.ini_path_edit, stretch=1)
        ini_row.addWidget(self.load_ini_btn)
        form.addRow("Cal INI File", ini_row)

        file_row = QtWidgets.QHBoxLayout()
        file_row.addWidget(self.file_edit, stretch=1)
        file_row.addWidget(self.browse_btn)
        form.addRow("CSV Output", file_row)

        controls_layout.addLayout(form)

        actions = QtWidgets.QHBoxLayout()
        actions.addWidget(self.start_btn)
        actions.addWidget(self.stop_btn)
        controls_layout.addLayout(actions)

        # ---- Baseline section ----
        sep = QtWidgets.QFrame()
        sep.setFrameShape(QtWidgets.QFrame.Shape.HLine)
        sep.setStyleSheet("color: rgba(122,2,25,0.2);")
        controls_layout.addWidget(sep)

        baseline_header = QtWidgets.QHBoxLayout()
        baseline_lbl = QtWidgets.QLabel("Baseline")
        baseline_lbl.setObjectName("consoleTitle")
        baseline_header.addWidget(baseline_lbl)
        baseline_header.addStretch()
        baseline_header.addWidget(self.baseline_btn)
        controls_layout.addLayout(baseline_header)

        baseline_bar = QtWidgets.QHBoxLayout()
        for lbl in (self.baseline_x_label, self.baseline_y_label, self.baseline_z_label):
            lbl.setObjectName("valuePill")
            lbl.setSizePolicy(
                QtWidgets.QSizePolicy.Policy.Expanding,
                QtWidgets.QSizePolicy.Policy.Preferred,
            )
            baseline_bar.addWidget(lbl)
        controls_layout.addLayout(baseline_bar)

        self.status_label = QtWidgets.QLabel("Idle")
        self.status_label.setObjectName("status")
        controls_layout.addWidget(self.status_label)

        console_title = QtWidgets.QLabel("Console")
        console_title.setObjectName("consoleTitle")
        controls_layout.addWidget(console_title)
        self.console_output = QtWidgets.QPlainTextEdit()
        self.console_output.setReadOnly(True)
        self.console_output.setMinimumHeight(80)
        self.console_output.setObjectName("console")
        self.console_output.setSizePolicy(
            QtWidgets.QSizePolicy.Policy.Preferred,
            QtWidgets.QSizePolicy.Policy.Expanding,
        )
        controls_layout.addWidget(self.console_output, stretch=1)

        view_card = QtWidgets.QFrame()
        view_card.setObjectName("card")
        view_card.setSizePolicy(
            QtWidgets.QSizePolicy.Policy.Expanding,
            QtWidgets.QSizePolicy.Policy.Expanding,
        )
        view_layout = QtWidgets.QVBoxLayout(view_card)
        view_layout.setContentsMargins(10, 10, 10, 10)
        view_layout.setSpacing(6)

        self._abs_axis = AbsoluteTimeAxis()
        self.plot = pg.PlotWidget(axisItems={"top": self._abs_axis})
        self.plot.setBackground("#f8f3e8")
        self.plot.showGrid(x=True, y=True, alpha=0.2)
        self.plot.getPlotItem().showAxis("top")
        _lbl_style = {"font-size": "13pt", "color": "#615250"}
        self.plot.setLabel("left", "Signal", **_lbl_style)
        self.plot.setLabel("bottom", "Time (s)", **_lbl_style)
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

        # Snap-to-point tooltip on hover
        self._mouse_proxy = pg.SignalProxy(
            self.plot.scene().sigMouseMoved, rateLimit=30, slot=self._on_mouse_moved
        )

        value_bar = QtWidgets.QHBoxLayout()
        self.value_x = QtWidgets.QLabel("X: --")
        self.value_y = QtWidgets.QLabel("Y: --")
        self.value_z = QtWidgets.QLabel("Z: --")
        for widget in (self.value_x, self.value_y, self.value_z):
            widget.setObjectName("valuePill")
            widget.setSizePolicy(
                QtWidgets.QSizePolicy.Policy.Expanding,
                QtWidgets.QSizePolicy.Policy.Preferred,
            )
            value_bar.addWidget(widget)

        view_layout.addWidget(self.plot, stretch=1)
        view_layout.addLayout(value_bar)

        left_host_layout.addWidget(controls_card)
        left_scroll.setWidget(left_host)
        splitter.addWidget(left_scroll)
        splitter.addWidget(view_card)
        splitter.setStretchFactor(0, 0)
        splitter.setStretchFactor(1, 1)
        splitter.setSizes([560, 760])

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
        self.load_ini_btn.clicked.connect(self._load_ini_calibration)
        self.baseline_btn.clicked.connect(self._take_baseline)
        self.start_btn.setEnabled(False)  # Requires baseline first

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
            QScrollArea#panelScroll {
                background: transparent;
                border: none;
            }
            QScrollArea#panelScroll > QWidget > QWidget {
                background: transparent;
            }
            QSplitter::handle {
                background: rgba(122, 2, 25, 0.08);
                border-radius: 4px;
            }
            QSplitter::handle:hover {
                background: rgba(122, 2, 25, 0.18);
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
        self.range_fact_spin.setValue(self._config.range_fact)
        self.ini_path_edit.setText(self._config.ini_path)
        self.file_edit.setText(self._config.output_file)

        # Auto-detect and apply INI calibration on first launch
        if not self._config.ini_path:
            auto_ini = _auto_find_ini()
            if auto_ini is not None:
                self.ini_path_edit.setText(str(auto_ini))
                self._apply_ini_calibration(auto_ini)

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
        self._baseline_raw = None
        self.baseline_x_label.setText("X: --")
        self.baseline_y_label.setText("Y: --")
        self.baseline_z_label.setText("Z: --")
        self.start_btn.setEnabled(False)
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

    def _take_baseline(self) -> None:
        """Take a single SQUID reading and store it as the baseline."""
        if not self._client.is_connected:
            QtWidgets.QMessageBox.warning(self, "Not Connected", "Connect to SQUID serial port first.")
            return
        if self._worker is not None:
            QtWidgets.QMessageBox.warning(self, "Acquisition Running",
                "Stop acquisition before taking a new baseline.")
            return

        self.baseline_btn.setEnabled(False)
        self._set_status("Taking baseline reading…")
        QtWidgets.QApplication.processEvents()
        try:
            x, y, z = self._client.read_xyz_volts()
        except Exception as exc:
            self._set_status(f"Baseline error: {exc}")
            self.baseline_btn.setEnabled(True)
            return

        self._baseline_raw = (x, y, z)
        self._update_baseline_labels()
        unit = self.unit_combo.currentText()
        self._set_status(
            f"Baseline taken — X:{x:+.4g}  Y:{y:+.4g}  Z:{z:+.4g} V  (displayed in {unit})"
        )
        self.baseline_btn.setEnabled(True)
        self.start_btn.setEnabled(True)

    def _load_ini_calibration(self) -> None:
        """Browse for (or use existing) INI path and load calibration."""
        current = self.ini_path_edit.text().strip()
        if not current or not Path(current).exists():
            path, _ = QtWidgets.QFileDialog.getOpenFileName(
                self, "Open Paleomag INI File", current or "", "INI Files (*.INI *.ini)"
            )
            if not path:
                return
            self.ini_path_edit.setText(path)
        self._apply_ini_calibration(Path(self.ini_path_edit.text().strip()))

    def _apply_ini_calibration(self, ini_path: Path) -> None:
        result = read_calibration_from_ini(ini_path)
        if result is None:
            QtWidgets.QMessageBox.warning(self, "INI Load Failed",
                f"Could not read [MagnetometerCalibration] from:\n{ini_path}")
            return
        xcal, ycal, zcal, rfact = result
        self.cal_x.setValue(xcal)
        self.cal_y.setValue(ycal)
        self.cal_z.setValue(zcal)
        self.range_fact_spin.setValue(rfact)
        self._append_console(
            f"Loaded calibration from INI: X={xcal}, Y={ycal}, Z={zcal}, RangeFact={rfact:.2e}"
        )

    def _start_logging(self) -> None:
        if not self._client.is_connected:
            QtWidgets.QMessageBox.warning(self, "Not Connected", "Connect to SQUID serial first.")
            return

        if self._baseline_raw is None:
            QtWidgets.QMessageBox.warning(self, "No Baseline",
                "Take a baseline reading before starting VRM logging.")
            return

        output_path = Path(self.file_edit.text().strip())
        if not output_path.name:
            QtWidgets.QMessageBox.warning(self, "Output Missing", "Choose a CSV output path.")
            return

        # --- CSV exists check ---
        append_to_existing = False
        overwrite = False
        if output_path.exists() and output_path.stat().st_size > 0:
            dlg = QtWidgets.QMessageBox(self)
            dlg.setWindowTitle("CSV File Already Exists")
            dlg.setText(f"<b>{output_path.name}</b> already exists.")
            dlg.setInformativeText("Append, overwrite, or choose a different file name?")
            append_btn = dlg.addButton("Append", QtWidgets.QMessageBox.ButtonRole.AcceptRole)
            overwrite_btn = dlg.addButton("Overwrite", QtWidgets.QMessageBox.ButtonRole.DestructiveRole)
            new_btn = dlg.addButton("Choose New Name", QtWidgets.QMessageBox.ButtonRole.ActionRole)
            dlg.addButton("Cancel", QtWidgets.QMessageBox.ButtonRole.RejectRole)
            dlg.setDefaultButton(append_btn)
            dlg.exec()
            clicked = dlg.clickedButton()
            if clicked is None or clicked.text() == "Cancel":
                return
            if clicked is new_btn:
                path, _ = QtWidgets.QFileDialog.getSaveFileName(
                    self, "Choose New CSV Output", str(output_path), "CSV Files (*.csv)"
                )
                if not path:
                    return
                self.file_edit.setText(path)
                output_path = Path(path)
            elif clicked is overwrite_btn:
                overwrite = True
            else:
                append_to_existing = True

        try:
            output_path.parent.mkdir(parents=True, exist_ok=True)
            is_new_file = not append_to_existing
            open_mode = "w" if overwrite else "a"
            self._csv_handle = output_path.open(open_mode, newline="", encoding="utf-8")
            self._csv_writer = csv.writer(self._csv_handle)
            if is_new_file:
                self._csv_writer.writerow(
                    [
                        "time_s",
                        "datetime_local",
                        "x_volts",
                        "y_volts",
                        "z_volts",
                        "x_baseline_v",
                        "y_baseline_v",
                        "z_baseline_v",
                        "x_display",
                        "y_display",
                        "z_display",
                        "display_unit",
                    ]
                )
        except OSError as exc:
            QtWidgets.QMessageBox.critical(self, "File Error", str(exc))
            return

        self._session_start_epoch = time.time()
        self._abs_axis.set_session_start(self._session_start_epoch)

        self._time.clear()
        self._x_vals.clear()
        self._y_vals.clear()
        self._z_vals.clear()

        self.start_btn.setEnabled(False)
        self.stop_btn.setEnabled(True)
        self.baseline_btn.setEnabled(False)

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
        # Baseline must be refreshed before the next logging run.
        self._baseline_raw = None
        self.baseline_x_label.setText("X: -- (take new baseline)")
        self.baseline_y_label.setText("Y: --")
        self.baseline_z_label.setText("Z: --")
        self.baseline_btn.setEnabled(True)
        self.start_btn.setEnabled(False)  # Disabled until fresh baseline taken
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
            rf = self.range_fact_spin.value()
            x = x * self.cal_x.value() * rf
            y = y * self.cal_y.value() * rf
            z = z * self.cal_z.value() * rf
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
        # Subtract baseline (always in volts before any unit conversion)
        bx, by, bz = self._baseline_raw if self._baseline_raw else (0.0, 0.0, 0.0)
        x_disp, y_disp, z_disp = self._display_values(
            sample.x_volts - bx,
            sample.y_volts - by,
            sample.z_volts - bz,
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
            abs_dt = datetime.fromtimestamp(self._session_start_epoch + sample.time_s)
            self._csv_writer.writerow(
                [
                    f"{sample.time_s:.6f}",
                    abs_dt.strftime("%Y-%m-%d %H:%M:%S.%f")[:-3],  # ms precision
                    f"{sample.x_volts:.9g}",
                    f"{sample.y_volts:.9g}",
                    f"{sample.z_volts:.9g}",
                    f"{bx:.9g}",
                    f"{by:.9g}",
                    f"{bz:.9g}",
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
        lbl_style = {"font-size": "13pt", "color": "#615250"}
        if self.unit_combo.currentText() == "Moment":
            self.plot.setLabel("left", "Moment (emu)", **lbl_style)
        else:
            self.plot.setLabel("left", "Voltage (V)", **lbl_style)
        self._update_baseline_labels()

    def _update_baseline_labels(self) -> None:
        """Redisplay baseline in whichever unit is currently selected."""
        if self._baseline_raw is None:
            return
        x, y, z = self._baseline_raw
        unit = self.unit_combo.currentText()
        if unit == "Moment":
            rf = self.range_fact_spin.value()
            xd = x * self.cal_x.value() * rf
            yd = y * self.cal_y.value() * rf
            zd = z * self.cal_z.value() * rf
            suffix = "emu"
        else:
            xd, yd, zd = x, y, z
            suffix = "V"
        self.baseline_x_label.setText(f"X: {xd:+.6g} {suffix}")
        self.baseline_y_label.setText(f"Y: {yd:+.6g} {suffix}")
        self.baseline_z_label.setText(f"Z: {zd:+.6g} {suffix}")

    def _on_mouse_moved(self, evt: tuple) -> None:
        pos = evt[0]
        if not self.plot.sceneBoundingRect().contains(pos):
            QtWidgets.QToolTip.hideText()
            return
        times = list(self._time)
        if not times:
            QtWidgets.QToolTip.hideText()
            return
        mouse_pt = self.plot.getViewBox().mapSceneToView(pos)
        t_cursor = mouse_pt.x()
        idx = min(range(len(times)), key=lambda i: abs(times[i] - t_cursor))
        # Only show tooltip when cursor is within 2% of plot width from a point
        vb_rect = self.plot.getViewBox().viewRect()
        snap_tol = vb_rect.width() * 0.02
        if abs(times[idx] - t_cursor) > snap_tol:
            QtWidgets.QToolTip.hideText()
            return
        xs = list(self._x_vals)
        ys = list(self._y_vals)
        zs = list(self._z_vals)
        unit = self.unit_combo.currentText()
        suffix = "emu" if unit == "Moment" else "V"
        tip = (
            f"t = {times[idx]:.3f} s\n"
            f"X: {xs[idx]:+.5g}\n"
            f"Y: {ys[idx]:+.5g}\n"
            f"Z: {zs[idx]:+.5g} {suffix}"
        )
        global_pos = self.plot.mapToGlobal(
            self.plot.mapFromScene(pos).toPoint()
        )
        QtWidgets.QToolTip.showText(global_pos, tip, self.plot)

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
            range_fact=self.range_fact_spin.value(),
            ini_path=self.ini_path_edit.text().strip(),
            window_geometry=self.saveGeometry().toHex().data().decode("ascii"),
        )
        save_config(self._config)


def main() -> int:
    app = QtWidgets.QApplication(sys.argv)
    pg.setConfigOptions(antialias=True)
    window = MainWindow()
    window.show()
    return app.exec()
