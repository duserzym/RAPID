from __future__ import annotations

from collections import deque
from collections.abc import Callable, Mapping
import sys
import threading
import time
from pathlib import Path
from typing import Any

from PySide6 import QtCore, QtGui, QtWidgets

try:
    import pyqtgraph as pg
except ImportError:  # pragma: no cover - operator-facing fallback
    pg = None


def _bootstrap_common_imports() -> None:
    root = Path(__file__).resolve().parents[2]
    if str(root) not in sys.path:
        sys.path.insert(0, str(root))


_bootstrap_common_imports()
from rapidpy_common.hardware import (  # noqa: E402
    MotorAxisConfig,
    MotorSerialClient,
    MotorTelemetry,
    MoveResult,
    convert_position_to_hole,
)
from rapidpy_common.ui import (  # noqa: E402
    apply_card_shadow,
    apply_liquid_glass_theme,
    workspace_window_size,
    apply_window_bounds_guard,
    set_app_icon,
)


class TelemetryThread(QtCore.QThread):
    """Continuously sample one axis without blocking Qt's GUI thread."""

    sample_ready = QtCore.Signal(object)
    poll_failed = QtCore.Signal(str)

    def __init__(self, client: MotorSerialClient, parent: QtCore.QObject | None = None) -> None:
        super().__init__(parent)
        self._client = client
        self._axis: MotorAxisConfig | None = None
        self._axis_lock = threading.Lock()
        self._stop_event = threading.Event()
        self._poll_interval_s = 0.10

    def set_axis(self, axis: MotorAxisConfig) -> None:
        with self._axis_lock:
            self._axis = axis

    def stop_monitoring(self) -> None:
        self._stop_event.set()

    def run(self) -> None:
        previous_error = ""
        while not self._stop_event.is_set():
            with self._axis_lock:
                axis = self._axis
            if axis is not None and self._client.is_connected:
                try:
                    sample = self._client.read_telemetry(axis)
                    self.sample_ready.emit(sample)
                    previous_error = ""
                except Exception as exc:  # serial failures are surfaced in the UI
                    message = str(exc)
                    if message != previous_error:
                        self.poll_failed.emit(message)
                        previous_error = message
            self._stop_event.wait(self._poll_interval_s)


class MotorCommandThread(QtCore.QThread):
    """Run one potentially blocking motor operation while telemetry continues."""

    succeeded = QtCore.Signal(object)
    failed = QtCore.Signal(str)

    def __init__(self, operation: Callable[[], Any], parent: QtCore.QObject | None = None) -> None:
        super().__init__(parent)
        self._operation = operation

    def run(self) -> None:
        try:
            self.succeeded.emit(self._operation())
        except Exception as exc:
            self.failed.emit(str(exc))


class MainWindow(QtWidgets.QMainWindow):
    TRACE_LIMIT = 2400
    _TORQUE_FULL_SCALE = 32767.0
    _TORQUE_FIELD_ALIASES = (
        "actual_torque",
        "feedback_torque",
        "feedback_torque_raw",
        "actual_torque_raw",
        "actual_torque_percent",
        "torque",
        "torque_current",
        "torque_feedback",
        "torque_adc",
        "torque_feedback_raw",
        "torque_raw",
        "torque_value",
        "current_torque",
        "motor_torque",
        "motor_torque_raw",
        "load_torque",
        "feedback_load",
    )
    _TIMESTAMP_FIELD_ALIASES = (
        "timestamp",
        "time",
        "ts",
        "time_ms",
        "sample_time",
    )
    _AXIS_FIELD_ALIASES = (
        "axis_name",
        "axis",
        "axis_id",
        "motor",
        "motor_axis",
        "axis_label",
    )
    _INPUT_SIGNAL_FIELD_ALIASES = (
        "target_position",
        "command",
        "command_position",
        "commanded_position",
        "input_position",
        "input_position_counts",
        "command_position_counts",
        "position_command",
        "command_position_setpoint",
        "command_pos",
        "input_signal",
        "position_cmd",
        "target",
        "setpoint",
        "input_cmd",
        "position_setpoint",
    )
    _OUTPUT_SIGNAL_FIELD_ALIASES = (
        "actual_position",
        "feedback_position",
        "feedback",
        "output_position",
        "output_signal",
        "position_feedback",
        "actual",
        "position",
        "actual_position_counts",
        "feedback_position_counts",
        "output_position_counts",
        "position_output",
        "measured_position",
        "motor_feedback",
        "encoder_position",
        "actual_feedback",
        "output_actual",
        "actual_pos",
        "encoder_counts",
    )
    _VALUE_GRID_COLUMNS = 2
    _VALUE_GRID_COMPACT_COLUMNS = 2
    _VALUE_GRID_NARROW_COLUMNS = 1
    _COMPACT_SPLIT_WIDTH = 1200
    _COMPACT_SPLIT_HEIGHT = 700
    _NARROW_SPLIT_WIDTH = 560
    _DEFAULT_WINDOW_SIZE = (1100, 760)
    _SPLIT_LEFT_WIDTH_MIN = 160
    _SPLIT_LEFT_WIDTH_MAX = 340

    def __init__(self) -> None:
        super().__init__()
        self.setWindowTitle("RapidPy DC Motor Control")
        self.resize(*self._DEFAULT_WINDOW_SIZE)
        self.client = MotorSerialClient()
        self._values_per_row = self._VALUE_GRID_COLUMNS
        self.axes = {
            "Changer (X)": MotorAxisConfig("ChangerX", 1, 1),
            "Turning": MotorAxisConfig("Turning", 2, 2),
            "Up/Down": MotorAxisConfig("UpDown", 3, 3),
            "Changer (Y)": MotorAxisConfig("ChangerY", 4, 4),
        }
        self._command_thread: MotorCommandThread | None = None
        self._last_sample: object | None = None
        self._trace_origin: float | None = None
        self._sample_tick: int = 0
        self._last_render_tick: int = 0
        self._torque_seen = False
        self._traces: dict[str, deque[float]] = {
            key: deque(maxlen=self.TRACE_LIMIT)
            for key in (
                "time",
                "target",
                "actual",
                "error",
                "velocity_1",
                "velocity_2",
                "io_delta",
                "torque",
            )
        }
        self._build_ui()

        self._telemetry_thread = TelemetryThread(self.client, self)
        self._telemetry_thread.set_axis(self._selected_axis())
        self._telemetry_thread.sample_ready.connect(self._on_telemetry)
        self._telemetry_thread.poll_failed.connect(self._on_poll_error)
        self._telemetry_thread.start()

        self._plot_timer = QtCore.QTimer(self)
        self._plot_timer.setInterval(100)
        self._plot_timer.timeout.connect(self._refresh_plots)
        self._plot_timer.start()

    def _build_ui(self) -> None:
        root = QtWidgets.QWidget(self)
        self.setCentralWidget(root)
        layout = QtWidgets.QVBoxLayout(root)
        layout.setContentsMargins(14, 14, 14, 14)
        layout.setSpacing(12)

        header = QtWidgets.QHBoxLayout()
        heading = QtWidgets.QVBoxLayout()
        title = QtWidgets.QLabel("Quicksilver Motor Lab")
        title.setObjectName("title")
        subtitle = QtWidgets.QLabel(
            "Command input/output and encoder feedback live trace for the active motor"
        )
        subtitle.setObjectName("subtitle")
        subtitle.setWordWrap(True)
        subtitle.setSizePolicy(
            QtWidgets.QSizePolicy.Policy.Ignored,
            QtWidgets.QSizePolicy.Policy.Preferred,
        )
        heading.addWidget(title)
        heading.addWidget(subtitle)
        header.addLayout(heading)
        header.addStretch(1)
        self.status = QtWidgets.QLabel("Disconnected")
        self.status.setObjectName("valuePill")
        header.addWidget(self.status)
        layout.addLayout(header)

        splitter = QtWidgets.QSplitter(QtCore.Qt.Orientation.Horizontal)
        splitter.setChildrenCollapsible(False)
        splitter.setHandleWidth(6)
        splitter.addWidget(self._build_control_scroll())
        splitter.addWidget(self._build_telemetry_panel())
        splitter.setStretchFactor(0, 0)
        splitter.setStretchFactor(1, 1)
        splitter.setSizes([
            self._SPLIT_LEFT_WIDTH_MAX,
            max(1, self._DEFAULT_WINDOW_SIZE[0] - self._SPLIT_LEFT_WIDTH_MAX),
        ])
        layout.addWidget(splitter, 1)
        self._splitter = splitter

        self.connect_btn.clicked.connect(self._connect)
        self.disconnect_btn.clicked.connect(self._disconnect)
        self.axis_combo.currentTextChanged.connect(self._axis_changed)
        self.move_btn.clicked.connect(self._move)
        self.spin_btn.clicked.connect(self._spin)
        self.goto_hole_btn.clicked.connect(self._goto_hole)
        self.read_hole_btn.clicked.connect(self._read_hole)
        self.home_top_btn.clicked.connect(self._home_to_top)
        self.home_center_btn.clicked.connect(self._home_xy_center)
        self.corner_btn.clicked.connect(self._move_xy_corner)
        self.pickup_btn.clicked.connect(self._sample_pickup)
        self.dropoff_btn.clicked.connect(self._sample_dropoff)
        self.clear_plot_btn.clicked.connect(self._clear_traces)

    def _build_control_scroll(self) -> QtWidgets.QScrollArea:
        scroll = QtWidgets.QScrollArea()
        scroll.setObjectName("panelScroll")
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QtWidgets.QFrame.Shape.NoFrame)
        scroll.setHorizontalScrollBarPolicy(QtCore.Qt.ScrollBarPolicy.ScrollBarAlwaysOff)

        left = QtWidgets.QFrame()
        left.setObjectName("card")
        left.setSizePolicy(QtWidgets.QSizePolicy.Policy.Preferred, QtWidgets.QSizePolicy.Policy.Expanding)
        l = QtWidgets.QVBoxLayout(left)
        l.setContentsMargins(18, 18, 18, 18)
        l.setSpacing(12)

        section = QtWidgets.QLabel("Connection")
        section.setObjectName("title")
        section.setStyleSheet("font-size:18px;")
        l.addWidget(section)

        conn_row = QtWidgets.QHBoxLayout()
        self.port_edit = QtWidgets.QLineEdit("COM4")
        self.port_edit.setPlaceholderText("COM port")
        self.connect_btn = QtWidgets.QPushButton("Connect")
        self.connect_btn.setObjectName("accent")
        self.disconnect_btn = QtWidgets.QPushButton("Disconnect")
        conn_row.addWidget(self.port_edit, 1)
        conn_row.addWidget(self.connect_btn)
        conn_row.addWidget(self.disconnect_btn)
        l.addLayout(conn_row)

        self.axis_combo = QtWidgets.QComboBox()
        self.axis_combo.addItems(list(self.axes.keys()))
        self.pos_spin = QtWidgets.QSpinBox()
        self.pos_spin.setRange(-5_000_000, 5_000_000)
        self.speed_spin = QtWidgets.QSpinBox()
        self.speed_spin.setRange(50, 500_000_000)
        self.speed_spin.setValue(1200)
        self.move_btn = QtWidgets.QPushButton("Move Axis")
        self.move_btn.setObjectName("accent")

        form = QtWidgets.QFormLayout()
        form.setFieldGrowthPolicy(QtWidgets.QFormLayout.FieldGrowthPolicy.AllNonFixedFieldsGrow)
        form.addRow("Active Axis", self.axis_combo)
        form.addRow("Target Position", self.pos_spin)
        form.addRow("Velocity", self.speed_spin)
        form.addRow("", self.move_btn)
        l.addLayout(form)

        divider = QtWidgets.QFrame()
        divider.setFrameShape(QtWidgets.QFrame.Shape.HLine)
        l.addWidget(divider)

        self.spin_rps = QtWidgets.QDoubleSpinBox()
        self.spin_rps.setRange(0.01, 20.0)
        self.spin_rps.setValue(1.0)
        self.spin_btn = QtWidgets.QPushButton("Spin Turning Motor")
        spin_form = QtWidgets.QFormLayout()
        spin_form.addRow("Spin (rps)", self.spin_rps)
        spin_form.addRow("", self.spin_btn)
        l.addLayout(spin_form)

        self.hole_spin = QtWidgets.QDoubleSpinBox()
        self.hole_spin.setRange(1.0, 101.0)
        self.hole_spin.setDecimals(1)
        self.goto_hole_btn = QtWidgets.QPushButton("Changer Motor To Hole")
        self.read_hole_btn = QtWidgets.QPushButton("Read Changer Hole")
        hole_form = QtWidgets.QFormLayout()
        hole_form.addRow("Target Hole", self.hole_spin)
        hole_form.addRow("", self.goto_hole_btn)
        hole_form.addRow("", self.read_hole_btn)
        l.addLayout(hole_form)

        ops = QtWidgets.QGridLayout()
        self.home_top_btn = QtWidgets.QPushButton("Home Z To Top")
        self.home_center_btn = QtWidgets.QPushButton("Home XY To Center")
        self.corner_btn = QtWidgets.QPushButton("Move XY To Corner")
        self.pickup_btn = QtWidgets.QPushButton("Sample Pickup")
        self.dropoff_btn = QtWidgets.QPushButton("Sample Dropoff")
        ops.addWidget(self.home_top_btn, 0, 0)
        ops.addWidget(self.home_center_btn, 0, 1)
        ops.addWidget(self.corner_btn, 1, 0)
        ops.addWidget(self.pickup_btn, 1, 1)
        ops.addWidget(self.dropoff_btn, 2, 0, 1, 2)
        l.addLayout(ops)
        l.addStretch(1)

        scroll.setWidget(left)
        apply_card_shadow(left)
        return scroll

    def _build_telemetry_panel(self) -> QtWidgets.QFrame:
        panel = QtWidgets.QFrame()
        panel.setObjectName("card")
        r = QtWidgets.QVBoxLayout(panel)
        r.setContentsMargins(16, 16, 16, 16)
        r.setSpacing(10)

        telemetry_header = QtWidgets.QVBoxLayout()
        legend = QtWidgets.QLabel("Live Controller Feedback")
        legend.setObjectName("title")
        legend.setStyleSheet("font-size:20px;")
        telemetry_header.addWidget(legend)
        trace_controls = QtWidgets.QHBoxLayout()
        trace_controls.addStretch(1)
        self.pause_plot = QtWidgets.QCheckBox("Pause traces")
        self.history_combo = QtWidgets.QComboBox()
        self.history_combo.addItems(["30 s", "60 s", "120 s", "240 s"])
        self.history_combo.setCurrentText("60 s")
        self.clear_plot_btn = QtWidgets.QPushButton("Clear")
        trace_controls.addWidget(self.pause_plot)
        trace_controls.addWidget(self.history_combo)
        trace_controls.addWidget(self.clear_plot_btn)
        telemetry_header.addLayout(trace_controls)
        r.addLayout(telemetry_header)

        values = QtWidgets.QGridLayout()
        values.setHorizontalSpacing(8)
        self._values_grid = values
        self.target_value = self._make_value("Input command --")
        self.actual_value = self._make_value("Output actual --")
        self.io_delta_value = self._make_value("Input/output delta --")
        self.error_value = self._make_value("Error --")
        self.velocity_value = self._make_value("Velocity1 --")
        self.velocity2_value = self._make_value("Velocity2 --")
        self.torque_value = self._make_value("Torque -- (N/A)")
        self._value_grid_tiles = (
            self.target_value,
            self.actual_value,
            self.io_delta_value,
            self.error_value,
            self.velocity_value,
            self.velocity2_value,
            self.torque_value,
        )
        self._apply_value_grid_columns(self._values_per_row)
        r.addLayout(values)

        self._build_plots(r)

        self.console = QtWidgets.QPlainTextEdit()
        self.console.setObjectName("console")
        self.console.setReadOnly(True)
        self.console.setMaximumBlockCount(500)
        self.console.setMaximumHeight(125)
        self.console.setPlaceholderText("Connection and motion events")
        r.addWidget(self.console)
        apply_card_shadow(panel)
        return panel

    def _apply_value_grid_columns(self, columns: int) -> None:
        if not hasattr(self, "_values_grid") or not hasattr(self, "_value_grid_tiles"):
            return
        columns = max(1, min(int(columns), max(1, len(self._value_grid_tiles))))

        layout = self._values_grid
        previous_columns = getattr(self, "_values_per_row", None)
        if columns == previous_columns and layout.count() == len(self._value_grid_tiles):
            return
        self._values_per_row = columns
        while layout.count():
            item = layout.takeAt(0)
            widget = item.widget() if item is not None else None
            if widget is not None:
                layout.removeWidget(widget)

        for index, widget in enumerate(self._value_grid_tiles):
            layout.addWidget(widget, index // columns, index % columns)

    @staticmethod
    def _make_value(text: str) -> QtWidgets.QLabel:
        label = QtWidgets.QLabel(text)
        label.setObjectName("valuePill")
        label.setAlignment(QtCore.Qt.AlignmentFlag.AlignCenter)
        label.setSizePolicy(QtWidgets.QSizePolicy.Policy.Expanding, QtWidgets.QSizePolicy.Policy.Fixed)
        return label

    @staticmethod
    def _read_sample_field(sample: object, field: str) -> object | None:
        if isinstance(sample, Mapping):
            if field in sample:
                return sample[field]
            field_key = field.lower()
            for key in sample:
                try:
                    if str(key).lower() == field_key:
                        return sample[key]  # type: ignore[index]
                except Exception:
                    continue
            return None

        if hasattr(sample, field):
            try:
                return getattr(sample, field)
            except Exception:
                return None
        return None

    @staticmethod
    def _coerce_float(value: object | None) -> float | None:
        if value is None:
            return None
        text = str(value).strip()
        if not text:
            return None
        text = text.replace(",", "")
        try:
            value_float = float(text)
        except (TypeError, ValueError):
            return None
        if value_float != value_float:
            return None
        return value_float

    @classmethod
    def _coerce_int(cls, value: object | None) -> int | None:
        number = cls._coerce_float(value)
        if number is None:
            return None
        return int(number)

    @staticmethod
    def _axis_token(value: object | None) -> str:
        text = ("" if value is None else str(value)).strip().lower()
        if not text:
            return ""
        return "".join(char for char in text if char.isalnum())

    def _extract_axis_token(self, sample: object) -> str:
        for field in self._AXIS_FIELD_ALIASES:
            axis_value = self._read_sample_field(sample, field)
            token = self._axis_token(axis_value)
            if token:
                return token
        return self._axis_token(self._selected_axis().name)

    def _axis_matches(self, sample: object) -> bool:
        return self._extract_axis_token(sample) == self._axis_token(self._selected_axis().name)

    @classmethod
    def _extract_timestamp(cls, sample: object) -> float | None:
        for field in cls._TIMESTAMP_FIELD_ALIASES:
            value = cls._coerce_float(cls._read_sample_field(sample, field))
            if value is not None:
                return value / 1000.0 if field == "time_ms" else value
        return None

    def _build_plots(self, layout: QtWidgets.QVBoxLayout) -> None:
        self._curves: dict[str, Any] = {}
        self._plots: list[Any] = []
        if pg is None:
            fallback = QtWidgets.QLabel("Live plotting requires pyqtgraph. Install the app requirements.")
            fallback.setAlignment(QtCore.Qt.AlignmentFlag.AlignCenter)
            fallback.setObjectName("valuePill")
            fallback.setWordWrap(True)
            fallback.setSizePolicy(
                QtWidgets.QSizePolicy.Policy.Ignored,
                QtWidgets.QSizePolicy.Policy.Expanding,
            )
            layout.addWidget(fallback, 1)
            return

        graph = pg.GraphicsLayoutWidget()
        graph.setBackground("#fffdf8")
        # Five stacked traces need ~130 px each to stay readable; on short
        # screens the graph scrolls instead of collapsing the plots together.
        graph.setMinimumHeight(5 * 130)

        position_plot = graph.addPlot(row=0, col=0, title="Command input vs encoder output")
        error_plot = graph.addPlot(row=1, col=0, title="Position error")
        velocity_plot = graph.addPlot(row=2, col=0, title="Actual velocity filters")
        delta_plot = graph.addPlot(row=3, col=0, title="Input-output delta")
        torque_plot = graph.addPlot(row=4, col=0, title="Actual torque feedback")
        self._plots = [position_plot, error_plot, velocity_plot, delta_plot, torque_plot]
        for plot in self._plots:
            plot.showGrid(x=True, y=True, alpha=0.16)
            plot.setDownsampling(auto=True, mode="peak")
            plot.setClipToView(True)
            plot.getAxis("left").setTextPen("#61534d")
            plot.getAxis("bottom").setTextPen("#61534d")
        for plot in self._plots[:-1]:
            plot.hideAxis("bottom")
        torque_plot.setLabel("bottom", "Elapsed", units="s")
        position_plot.setLabel("left", "Position", units="counts")
        error_plot.setLabel("left", "Error", units="counts")
        velocity_plot.setLabel("left", "Velocity", units="SAV")
        delta_plot.setLabel("left", "Input-Output", units="counts")
        torque_plot.setLabel("left", "Torque", units="controller")
        error_plot.setXLink(position_plot)
        velocity_plot.setXLink(position_plot)
        delta_plot.setXLink(position_plot)
        torque_plot.setXLink(position_plot)

        position_plot.addLegend(offset=(8, 8))
        velocity_plot.addLegend(offset=(8, 8))
        delta_plot.addLegend(offset=(8, 8))
        self._curves["target"] = position_plot.plot(
            pen=pg.mkPen("#d69f00", width=2), name="Target / input command"
        )
        self._curves["actual"] = position_plot.plot(
            pen=pg.mkPen("#7a0219", width=2), name="Actual / output feedback"
        )
        self._curves["error"] = error_plot.plot(pen=pg.mkPen("#d05f2d", width=2))
        self._curves["velocity_1"] = velocity_plot.plot(
            pen=pg.mkPen("#31566d", width=2), name="Filter 1"
        )
        self._curves["velocity_2"] = velocity_plot.plot(
            pen=pg.mkPen("#70a2b8", width=1), name="Filter 2"
        )
        self._curves["io_delta"] = delta_plot.plot(
            pen=pg.mkPen("#8f4ba8", width=2), name="Actual - Command"
        )
        self._curves["torque"] = torque_plot.plot(pen=pg.mkPen("#28755f", width=2))
        graph_scroll = QtWidgets.QScrollArea()
        graph_scroll.setObjectName("panelScroll")
        graph_scroll.setWidgetResizable(True)
        graph_scroll.setFrameShape(QtWidgets.QFrame.Shape.NoFrame)
        graph_scroll.setHorizontalScrollBarPolicy(QtCore.Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        graph_scroll.setWidget(graph)
        layout.addWidget(graph_scroll, 1)
        self._graph = graph
        self._graph_scroll = graph_scroll

    def _extract_torque(self, sample: object) -> float | None:
        for attr in self._TORQUE_FIELD_ALIASES:
            value = self._coerce_float(self._read_sample_field(sample, attr))
            if value is not None:
                return value
        return None

    @classmethod
    def _extract_signal_position(cls, sample: object, aliases: tuple[str, ...]) -> int | None:
        for alias in aliases:
            value = cls._coerce_int(cls._read_sample_field(sample, alias))
            if value is not None:
                return value
        return None

    def _selected_axis(self) -> MotorAxisConfig:
        return self.axes[self.axis_combo.currentText()]

    def _axis_changed(self, _name: str) -> None:
        axis = self._selected_axis()
        self._telemetry_thread.set_axis(axis)
        self._clear_traces()
        self._log(f"Telemetry axis: {self.axis_combo.currentText()} (@{axis.address})")

    def _connect(self) -> None:
        if self._command_is_running():
            self._warn_busy()
            return
        try:
            self.client.connect(self.port_edit.text().strip(), baudrate=57600)
        except Exception as exc:
            QtWidgets.QMessageBox.critical(self, "Connection Error", str(exc))
            return
        self._clear_traces()
        self._log(f"Connected to motor serial on {self.port_edit.text().strip()}; telemetry active")

    def _disconnect(self) -> None:
        if self._command_is_running():
            self._warn_busy()
            return
        self.client.disconnect()
        self._last_sample = None
        self._log("Disconnected")

    def _log(self, text: str) -> None:
        self.console.appendPlainText(text)
        self.status.setText(text)

    def _require_connection(self) -> bool:
        if self.client.is_connected:
            return True
        QtWidgets.QMessageBox.warning(
            self, "Not Connected", "Connect motor serial before issuing motor commands."
        )
        return False

    def _command_is_running(self) -> bool:
        return self._command_thread is not None and self._command_thread.isRunning()

    def _warn_busy(self) -> None:
        QtWidgets.QMessageBox.information(
            self, "Motor Busy", "Wait for the active motor operation to finish before starting another."
        )

    def _set_motion_enabled(self, enabled: bool) -> None:
        for button in (
            self.move_btn,
            self.spin_btn,
            self.goto_hole_btn,
            self.read_hole_btn,
            self.home_top_btn,
            self.home_center_btn,
            self.corner_btn,
            self.pickup_btn,
            self.dropoff_btn,
            self.connect_btn,
            self.disconnect_btn,
        ):
            button.setEnabled(enabled)

    def _run_command(
        self,
        label: str,
        operation: Callable[[], Any],
        formatter: Callable[[Any], str],
    ) -> None:
        if not self._require_connection():
            return
        if self._command_is_running():
            self._warn_busy()
            return
        self._set_motion_enabled(False)
        self._log(f"{label} started")
        thread = MotorCommandThread(operation, self)
        thread.succeeded.connect(lambda result: self._log(formatter(result)))
        thread.failed.connect(lambda message: self._command_failed(label, message))
        thread.finished.connect(self._command_finished)
        self._command_thread = thread
        thread.start()

    def _command_failed(self, label: str, message: str) -> None:
        self._log(f"{label} failed: {message}")
        QtWidgets.QMessageBox.warning(self, f"{label} Error", message)

    def _command_finished(self) -> None:
        self._set_motion_enabled(True)
        if self._command_thread is not None:
            self._command_thread.deleteLater()
            self._command_thread = None

    def _move(self) -> None:
        axis = self._selected_axis()
        target = self.pos_spin.value()
        speed = self.speed_spin.value()
        self._run_command(
            f"{axis.name} move",
            lambda: self.client.move_motor(axis, target, speed, wait_for_stop=True),
            lambda result: self._format_move(axis.name, result),
        )

    @staticmethod
    def _format_move(label: str, result: MoveResult) -> str:
        return (
            f"{label}: target={result.target:,}, final={result.final_position:,}, "
            f"success={result.success}"
        )

    def _spin(self) -> None:
        rps = self.spin_rps.value()
        turning = self.axes["Turning"]
        self.axis_combo.setCurrentText("Turning")
        self._run_command(
            "Turning spin",
            lambda: self.client.turning_motor_spin(turning, speed_rps=rps, duration_s=60.0),
            lambda result: self._format_move("Turning spin armed", result),
        )

    def _goto_hole(self) -> None:
        hole = self.hole_spin.value()
        axis = self.axes["Changer (X)"]
        self.axis_combo.setCurrentText("Changer (X)")
        self._run_command(
            f"Changer move to hole {hole:.2f}",
            lambda: self.client.changer_motor_to_hole(axis, hole, wait_for_stop=True),
            lambda result: self._format_move(f"Changer hole {hole:.2f}", result),
        )

    def _read_hole(self) -> None:
        axis = self.axes["Changer (X)"]
        self.axis_combo.setCurrentText("Changer (X)")
        self._run_command(
            "Read changer hole",
            lambda: self.client.read_position(axis),
            lambda position: (
                f"Changer position {position:,} = hole "
                f"{convert_position_to_hole(position, slot_min=1, slot_max=101, one_step=-1000):.2f}"
            ),
        )

    def _home_to_top(self) -> None:
        axis = self.axes["Up/Down"]
        self.axis_combo.setCurrentText("Up/Down")
        self._run_command(
            "Home Z to top",
            lambda: self.client.home_to_top(axis),
            lambda result: self._format_move("Home Z", result),
        )

    def _home_xy_center(self) -> None:
        self._run_command(
            "Home XY to center",
            lambda: self.client.home_xy_to_center(
                self.axes["Changer (X)"], self.axes["Changer (Y)"], self.axes["Up/Down"]
            ),
            lambda results: (
                f"Home XY: x={results[0].final_position:,}, y={results[1].final_position:,}, "
                f"success={results[0].success and results[1].success}"
            ),
        )

    def _move_xy_corner(self) -> None:
        self._run_command(
            "Move XY to corner",
            lambda: self.client.move_xy_to_corner(
                self.axes["Changer (X)"], self.axes["Changer (Y)"], self.axes["Up/Down"]
            ),
            lambda results: (
                f"XY corner: x={results[0].final_position:,}, y={results[1].final_position:,}"
            ),
        )

    def _sample_pickup(self) -> None:
        axis = self.axes["Up/Down"]
        self.axis_combo.setCurrentText("Up/Down")
        self._run_command(
            "Sample pickup",
            lambda: self.client.sample_pickup(axis),
            lambda result: self._format_move("Sample pickup", result),
        )

    def _sample_dropoff(self) -> None:
        axis = self.axes["Up/Down"]
        self.axis_combo.setCurrentText("Up/Down")
        self._run_command(
            "Sample dropoff",
            lambda: self.client.sample_dropoff(axis, use_xy_table=True),
            lambda result: self._format_move("Sample dropoff", result),
        )

    @QtCore.Slot(object)
    def _on_telemetry(self, sample: object) -> None:
        if not self._axis_matches(sample):
            return

        timestamp = self._extract_timestamp(sample)
        if timestamp is None:
            timestamp = time.monotonic()
        if timestamp is None:
            return

        self._sample_tick += 1

        target_position = self._extract_signal_position(sample, self._INPUT_SIGNAL_FIELD_ALIASES)
        actual_position = self._extract_signal_position(sample, self._OUTPUT_SIGNAL_FIELD_ALIASES)
        if target_position is None or actual_position is None:
            return

        self._last_sample = sample
        position_error = self._coerce_int(self._read_sample_field(sample, "position_error"))
        velocity_1 = self._coerce_int(self._read_sample_field(sample, "velocity_1"))
        velocity_2 = self._coerce_int(self._read_sample_field(sample, "velocity_2"))

        self.target_value.setText(f"Input command {target_position:,}")
        self.actual_value.setText(f"Output feedback {actual_position:,}")
        io_delta = int(actual_position - target_position)
        self.io_delta_value.setText(f"Input-output delta {io_delta:+,}")

        if position_error is None:
            self.error_value.setText("Error --")
            self._traces["error"].append(float("nan"))
        else:
            self.error_value.setText(f"Error {position_error:+,}")
            self._traces["error"].append(float(position_error))

        if velocity_1 is None:
            self.velocity_value.setText("Velocity1 --")
            self._traces["velocity_1"].append(float("nan"))
        else:
            self.velocity_value.setText(f"Velocity1 {velocity_1:+,} SAV")
            self._traces["velocity_1"].append(float(velocity_1))

        if velocity_2 is None:
            self.velocity2_value.setText("Velocity2 --")
            self._traces["velocity_2"].append(float("nan"))
        else:
            self.velocity2_value.setText(f"Velocity2 {velocity_2:+,} SAV")
            self._traces["velocity_2"].append(float(velocity_2))

        actual_torque = self._extract_torque(sample)
        if actual_torque is None:
            self.torque_value.setText("Torque -- (N/A)")
            self._traces["torque"].append(float("nan"))
        else:
            self._torque_seen = True
            self._traces["torque"].append(float(actual_torque))
            torque_percent = (float(actual_torque) / self._TORQUE_FULL_SCALE) * 100.0
            self.torque_value.setText(
                f"Torque {int(actual_torque):+,} ({torque_percent:+.1f}% FS)"
            )

        if self._trace_origin is None:
            self._trace_origin = timestamp
        elapsed = timestamp - self._trace_origin
        self._traces["time"].append(elapsed)
        self._traces["target"].append(float(target_position))
        self._traces["actual"].append(float(actual_position))
        self._traces["io_delta"].append(float(io_delta))
        if not self._command_is_running():
            self.status.setText(
                f"Live @{self._selected_axis().address} | {actual_position:,} counts"
            )

    @QtCore.Slot(str)
    def _on_poll_error(self, message: str) -> None:
        self._log(f"Telemetry read error: {message}")

    def _history_seconds(self) -> float:
        try:
            return float(self.history_combo.currentText().split()[0])
        except (ValueError, IndexError):
            return 60.0

    def _refresh_plots(self) -> None:
        if (
            pg is None
            or self.pause_plot.isChecked()
            or self._sample_tick == self._last_render_tick
            or not self._traces["time"]
        ):
            return
        self._last_render_tick = self._sample_tick
        history = self._history_seconds()
        latest = self._traces["time"][-1]
        cutoff = latest - history
        while self._traces["time"] and self._traces["time"][0] < cutoff:
            for trace in self._traces.values():
                trace.popleft()
        times = list(self._traces["time"])
        for name, curve in self._curves.items():
            trace = self._traces[name]
            if name == "torque" and not self._torque_seen:
                continue
            if not trace:
                curve.setData([], [])
                continue
            curve.setData(times[-len(trace) :], list(trace), skipFiniteCheck=True)
        if latest > history:
            self._plots[0].setXRange(latest - history, latest, padding=0.01)

    def _clear_traces(self) -> None:
        for trace in self._traces.values():
            trace.clear()
        self._trace_origin = None
        self._sample_tick = 0
        self._last_render_tick = 0
        self._torque_seen = False
        if pg is not None:
            for curve in self._curves.values():
                curve.setData([], [])

        self.target_value.setText("Input command --")
        self.actual_value.setText("Output actual --")
        self.error_value.setText("Error --")
        self.io_delta_value.setText("Input-output delta --")
        self.velocity_value.setText("Velocity1 --")
        self.velocity2_value.setText("Velocity2 --")
        self.torque_value.setText("Torque -- (N/A)")

    def showEvent(self, event: QtGui.QShowEvent) -> None:
        super().showEvent(event)
        QtCore.QTimer.singleShot(0, self._fit_to_screen)
        handle = self.windowHandle()
        if handle is not None and not getattr(self, "_screen_signal_connected", False):
            handle.screenChanged.connect(self._fit_to_screen)
            self._screen_signal_connected = True

    @QtCore.Slot()
    def _fit_to_screen(self, screen: QtGui.QScreen | None = None) -> None:
        active_screen = screen or self.screen() or QtWidgets.QApplication.primaryScreen()
        if active_screen is None:
            return
        available = active_screen.availableGeometry()
        fitted_w, fitted_h = workspace_window_size(available, (self.width(), self.height()))
        if self.isMaximized():
            self.showNormal()
        # The window may grow to the work area; only the startup size is fitted.
        self.setMaximumSize(available.width(), available.height())

        min_size = self.minimumSize()
        if min_size.isValid() and not min_size.isNull():
            self.setMinimumSize(min(min_size.width(), fitted_w), min(min_size.height(), fitted_h))

        # Stack the panels only when the screen is narrow; a short but wide
        # screen keeps them side by side so the traces keep their height.
        compact = available.width() < self._COMPACT_SPLIT_WIDTH
        if available.width() < self._NARROW_SPLIT_WIDTH:
            wanted_columns = self._VALUE_GRID_NARROW_COLUMNS
        elif compact:
            wanted_columns = self._VALUE_GRID_COMPACT_COLUMNS
        else:
            wanted_columns = self._VALUE_GRID_COLUMNS
        self._apply_value_grid_columns(wanted_columns)
        wanted_orientation = (
            QtCore.Qt.Orientation.Vertical if compact else QtCore.Qt.Orientation.Horizontal
        )
        if self._splitter.orientation() != wanted_orientation:
            self._splitter.setOrientation(wanted_orientation)

        if self._splitter.orientation() == QtCore.Qt.Orientation.Horizontal:
            left = max(
                self._SPLIT_LEFT_WIDTH_MIN,
                min(self._SPLIT_LEFT_WIDTH_MAX, fitted_w // 3),
            )
            self._splitter.setSizes([left, max(1, fitted_w - left)])
        else:
            top = max(
                self._SPLIT_LEFT_WIDTH_MIN,
                min(self._SPLIT_LEFT_WIDTH_MAX, fitted_h // 3),
            )
            self._splitter.setSizes([top, max(1, fitted_h - top)])

        self.resize(
            min(self.width(), fitted_w),
            min(self.height(), fitted_h),
        )
        self.centralWidget().layout().activate()
        frame = self.frameGeometry()
        frame.setSize(QtCore.QSize(
            min(frame.width(), fitted_w),
            min(frame.height(), fitted_h),
        ))
        if not available.intersects(frame):
            frame.moveCenter(available.center())
        else:
            x = max(
                available.left(),
                min(frame.left(), available.right() - frame.width() + 1),
            )
            y = max(
                available.top(),
                min(frame.top(), available.bottom() - frame.height() + 1),
            )
            frame.moveTopLeft(QtCore.QPoint(x, y))
        self.setGeometry(frame)

    def closeEvent(self, event: QtGui.QCloseEvent) -> None:
        if self._command_is_running():
            QtWidgets.QMessageBox.warning(
                self,
                "Motor Operation Active",
                "Wait for the active motor operation to finish before closing the controller.",
            )
            event.ignore()
            return
        self._plot_timer.stop()
        self._telemetry_thread.stop_monitoring()
        self._telemetry_thread.wait(1000)
        self.client.disconnect()
        event.accept()


def main() -> int:
    app = QtWidgets.QApplication(sys.argv)
    apply_window_bounds_guard(app)
    apply_liquid_glass_theme(app)
    set_app_icon(app, "dc_motor_control_icon.png", Path(__file__).resolve().parent.parent / "assets")
    window = MainWindow()
    set_app_icon(window, "dc_motor_control_icon.png", Path(__file__).resolve().parent.parent / "assets")
    window.show()
    return app.exec()
