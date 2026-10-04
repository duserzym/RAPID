from __future__ import annotations

from collections import deque
import math
import threading
import time
from collections.abc import Mapping

from PySide6 import QtCore, QtGui, QtWidgets

from rapid_main.diagnostic_services import (
    DCMotorBackend,
    build_backend_or_unavailable,
    build_dcmotor_backend,
)
from rapid_main.glass_theme import set_semantic_status
from rapidpy_common.hardware import MotorTelemetry

try:
    import pyqtgraph as pg
except ImportError:  # pragma: no cover - optional dependency path
    pg = None


class MotorCommandWorker(QtCore.QObject):
    succeeded = QtCore.Signal(str, object)
    failed = QtCore.Signal(str, str)
    settled = QtCore.Signal()

    def __init__(self, backend, label, action):
        super().__init__()
        self.backend, self.label, self.action = backend, label, action
        self._stop_event = threading.Event()

    def stop(self):
        self._stop_event.set()
        stop = getattr(self.backend, 'request_stop', None)
        if callable(stop):
            stop()

    @QtCore.Slot()
    def run(self):
        check = getattr(self.backend, 'set_cancel_check', None)
        try:
            if self._stop_event.is_set():
                raise InterruptedError('Command cancelled before initialization.')
            if callable(check):
                check(self._stop_event.is_set)
            result = self.action()
            if self._stop_event.is_set():
                raise InterruptedError('Command stopped; cleanup has settled.')
            self.succeeded.emit(self.label, result)
        except Exception as exc:
            self.failed.emit(self.label, str(exc))
        finally:
            if callable(check):
                check(None)
            self.settled.emit()


class TelemetryThread(QtCore.QThread):
    """Continuously sample one axis while the rest of the dialog stays responsive."""

    sample_ready = QtCore.Signal(object)
    poll_failed = QtCore.Signal(str)

    def __init__(self, backend: DCMotorBackend, parent: QtCore.QObject | None = None) -> None:
        super().__init__(parent)
        self._backend = backend
        self._axis: str | None = None
        self._axis_lock = threading.Lock()
        self._stop_event = threading.Event()
        self._poll_interval_s = 0.10

    def set_axis(self, axis: str) -> None:
        with self._axis_lock:
            self._axis = axis

    def stop_monitoring(self) -> None:
        self._stop_event.set()

    def run(self) -> None:
        previous_error = ""
        while not self._stop_event.is_set():
            with self._axis_lock:
                axis = self._axis

            if (
                axis is not None
                and hasattr(self._backend, "read_telemetry")
                and self._backend.is_connected()
            ):
                try:
                    sample = self._backend.read_telemetry(axis)  # type: ignore[attr-defined]
                    self.sample_ready.emit(sample)
                    previous_error = ""
                except Exception as exc:
                    message = str(exc)
                    if message != previous_error:
                        self.poll_failed.emit(message)
                        previous_error = message
            self._stop_event.wait(self._poll_interval_s)


class DCMotorDialog(QtWidgets.QDialog):
    """In-process DC motor control dialog used from rapid_main diagnostics."""

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
        "torque_raw_count",
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
        "motor_command",
        "setpoint_position",
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
    _VALUE_GRID_COLUMNS = 3
    _VALUE_GRID_COMPACT_COLUMNS = 2
    _VALUE_GRID_NARROW_COLUMNS = 1
    _COMPACT_SPLIT_WIDTH = 1250
    _COMPACT_SPLIT_HEIGHT = 700
    _NARROW_SPLIT_WIDTH = 560
    _LOG_EVERY_N_SAMPLES = 12
    _MIN_WINDOW_SIZE = (500, 500)
    _SPLIT_LEFT_WIDTH_MIN = 160
    _SPLIT_LEFT_WIDTH_MAX = 240

    def __init__(
        self,
        parent: QtWidgets.QWidget | None = None,
        *,
        backend: DCMotorBackend | None = None,
        port: str = "COM3",
        baud: int = 9600,
    ) -> None:
        super().__init__(parent)
        self.setObjectName("glassDialog")
        self.setWindowTitle("DC Motor Control")
        self.setAccessibleName("QuickSilver DC motor diagnostics and control")
        self.setMinimumWidth(self._MIN_WINDOW_SIZE[0])
        self.setMinimumHeight(self._MIN_WINDOW_SIZE[1])
        self.resize(*self._MIN_WINDOW_SIZE)
        self.setWindowFlags(self.windowFlags() & ~QtCore.Qt.WindowContextHelpButtonHint)

        # Hardware mode fails closed: a missing adapter becomes a visible
        # "unavailable" backend, never a silent simulator.
        self._backend = backend or build_backend_or_unavailable(
            "DC motors",
            build_dcmotor_backend,
            port=port,
            baud=int(baud),
            nocomm=False,
        )
        availability_check = getattr(self._backend, "is_available", None)
        self._backend_available = (
            bool(availability_check()) if callable(availability_check) else True
        )
        self._axes = self._discover_axes()
        self._connected = False
        self._default_port = port
        self._default_baud = baud
        self._torque_seen = False
        self._trace_origin: float | None = None
        self._sample_tick = 0
        self._last_render_tick = 0
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

        self._plot_timer = QtCore.QTimer(self)
        self._command_thread = None
        self._command_worker = None
        self._close_after_cleanup = False
        self._last_command_error = ''
        self._close_timer = QtCore.QTimer(self)
        self._close_timer.setSingleShot(True)
        self._close_timer.timeout.connect(self._resume_close)
        self._command_settle_timer = QtCore.QTimer(self)
        self._command_settle_timer.setSingleShot(True)
        self._command_settle_timer.timeout.connect(self._command_settled)
        self._plot_timer.setInterval(100)
        self._plot_timer.timeout.connect(self._refresh_plots)
        self._values_per_row = self._VALUE_GRID_COLUMNS
        self._build_ui()
        self._refresh_axis_summary()

        self._telemetry_thread = TelemetryThread(self._backend, self)
        self._telemetry_thread.sample_ready.connect(self._on_telemetry)
        self._telemetry_thread.poll_failed.connect(self._on_poll_error)
        self._telemetry_thread.finished.connect(self._resume_close)
        self._telemetry_thread.start()
        self._plot_timer.start()

        self._telemetry_thread.set_axis(self._selected_axis())

    def _discover_axes(self) -> tuple[str, ...]:
        try:
            axes = self._backend.available_axes()
        except Exception:
            axes = (
                "Changer (X)",
                "Turning",
                "Up/Down",
                "Changer (Y)",
            )
        return tuple(axes) if axes else (
            "Changer (X)",
            "Turning",
            "Up/Down",
            "Changer (Y)",
        )

    def _build_ui(self) -> None:
        root = QtWidgets.QVBoxLayout(self)
        root.setContentsMargins(16, 16, 16, 16)
        root.setSpacing(10)

        header = QtWidgets.QVBoxLayout()
        heading = QtWidgets.QVBoxLayout()
        title = QtWidgets.QLabel("Quicksilver Motor Control")
        title.setObjectName("dialogTitle")
        title.setAccessibleName("QuickSilver motor control")
        subtitle = QtWidgets.QLabel(
            "Live command, encoder, velocity, and torque feedback for any selected QuickSilver axis."
        )
        subtitle.setObjectName("dialogSubtitle")
        subtitle.setWordWrap(True)
        heading.addWidget(title)
        heading.addWidget(subtitle)
        header.addLayout(heading)
        summary = QtWidgets.QHBoxLayout()
        self._status = QtWidgets.QLabel()
        self._status.setWordWrap(True)
        self._status.setMaximumHeight(64)
        self._axis_summary = QtWidgets.QLabel("Monitoring axis: --")
        self._axis_summary.setWordWrap(True)
        self._axis_summary.setObjectName("valuePill")
        self._axis_summary.setAccessibleName("Monitored motor axis")
        summary.addWidget(self._status, 1)
        summary.addWidget(self._axis_summary, 1)
        header.addLayout(summary)
        root.addLayout(header)
        safety = QtWidgets.QHBoxLayout()
        self.stop_motion_btn = QtWidgets.QPushButton('Stop Active Motion')
        self.stop_motion_btn.setEnabled(False)
        self.stop_motion_btn.clicked.connect(self._stop_command)
        self.recover_stop_btn = QtWidgets.QPushButton('Recover: Verify Stopped')
        self.recover_stop_btn.clicked.connect(self._recover_stop)
        safety.addWidget(self.stop_motion_btn)
        safety.addWidget(self.recover_stop_btn)
        root.addLayout(safety)

        splitter = QtWidgets.QSplitter(QtCore.Qt.Orientation.Horizontal)
        splitter.setChildrenCollapsible(False)
        splitter.setHandleWidth(6)
        self._control_panel = self._build_control_panel()
        self._telemetry_panel = self._build_telemetry_panel()
        splitter.addWidget(self._control_panel)
        splitter.addWidget(self._telemetry_panel)
        splitter.setStretchFactor(0, 0)
        splitter.setStretchFactor(1, 1)
        splitter.setSizes([self._SPLIT_LEFT_WIDTH_MAX, max(1, self._MIN_WINDOW_SIZE[0] - self._SPLIT_LEFT_WIDTH_MAX)])
        self._splitter = splitter
        self._workspace_stack = QtWidgets.QStackedWidget()
        self._compact_tabs = QtWidgets.QTabWidget()
        self._workspace_stack.addWidget(splitter)
        self._workspace_stack.addWidget(self._compact_tabs)
        root.addWidget(self._workspace_stack, 1)

        self.connect_btn.clicked.connect(self._connect)
        self.disconnect_btn.clicked.connect(self._disconnect)
        self.axis_combo.currentTextChanged.connect(self._axis_changed)
        self.move_btn.clicked.connect(self._move_axis)
        self.spin_btn.clicked.connect(self._spin_turning)
        self.goto_hole_btn.clicked.connect(self._goto_hole)
        self.read_hole_btn.clicked.connect(self._read_hole)
        self.home_top_btn.clicked.connect(self._home_to_top)
        self.home_xy_btn.clicked.connect(self._home_xy_to_center)
        self.corner_btn.clicked.connect(self._move_xy_to_corner)
        self.pickup_btn.clicked.connect(self._sample_pickup)
        self.dropoff_btn.clicked.connect(self._sample_dropoff)
        self.clear_plot_btn.clicked.connect(self._clear_traces)
        self._configure_accessibility()

    def _configure_accessibility(self) -> None:
        controls = (
            (self.stop_motion_btn, 'Stop active motor motion', 'Requests cancellation and waits for acknowledged stop and zero velocity readback.'),
            (self.recover_stop_btn, 'Recover motor stopped state', 'Verifies the original axes stopped in place; never homes or replays motion.'),
            (self.port_edit, "Motor controller serial port", "Serial port used only when Connect is pressed."),
            (self.baud_combo, "Motor controller baud rate", "Baud rate used only when Connect is pressed."),
            (self.connect_btn, "Connect motor controller", "Opens the configured motor-controller connection; it does not move an axis."),
            (self.disconnect_btn, "Disconnect motor controller", "Stops live telemetry and closes the motor-controller connection."),
            (self.axis_combo, "Active motor axis", "Selects the axis monitored and used by Move Axis."),
            (self.pos_spin, "Target motor position", "Target encoder position for the selected axis."),
            (self.speed_spin, "Motor velocity", "Requested controller velocity for Move Axis."),
            (self.move_btn, "Move selected motor axis", "Moves physical hardware when a live backend is connected."),
            (self.spin_rps, "Turning motor rotations per second", "Requested turning speed."),
            (self.spin_btn, "Spin turning motor", "Moves physical turning hardware for the configured duration when connected."),
            (self.hole_spin, "Target sample changer hole", "Target changer position from hole 1 through 101."),
            (self.goto_hole_btn, "Move changer to target hole", "Moves physical changer hardware when connected."),
            (self.read_hole_btn, "Read current changer hole", "Reads changer encoder position without commanding motion."),
            (self.home_top_btn, "Home vertical axis to top", "Homes physical vertical hardware when connected."),
            (self.home_xy_btn, "Home XY axes to center", "Homes physical XY hardware when connected."),
            (self.corner_btn, "Move XY axes to corner", "Moves physical XY hardware when connected."),
            (self.pickup_btn, "Run sample pickup", "Runs the connected vertical sample-pickup motion."),
            (self.dropoff_btn, "Run sample dropoff", "Runs the connected sample-dropoff motion."),
            (self.pause_plot, "Pause motor telemetry plots", "Pauses plot rendering without stopping telemetry acquisition."),
            (self.history_combo, "Motor telemetry history duration", "Selects the visible trace duration."),
            (self.clear_plot_btn, "Clear motor telemetry", "Clears current traces and live readouts."),
            (self._console, "Motor connection and motion event log", "Read-only diagnostic events; not proof of hardware acceptance by itself."),
        )
        for widget, name, description in controls:
            widget.setAccessibleName(name)
            widget.setAccessibleDescription(description)

        readouts = {
            self.target_value: "Motor input command",
            self.actual_value: "Motor output feedback",
            self.io_delta_value: "Motor input-output delta",
            self.error_value: "Motor position error",
            self.velocity_value: "Motor velocity filter one",
            self.velocity2_value: "Motor velocity filter two",
            self.torque_value: "Motor torque feedback",
        }
        for widget, name in readouts.items():
            widget.setAccessibleName(name)

    def showEvent(self, event: QtGui.QShowEvent) -> None:
        super().showEvent(event)
        QtCore.QTimer.singleShot(0, self._fit_to_screen)
        handle = self.windowHandle()
        if handle is not None and not getattr(self, "_screen_signal_connected", False):
            if handle.screen() is not None:
                handle.screen().availableGeometryChanged.connect(
                    lambda *_sig_args: self._fit_to_screen(),
                )
            handle.screenChanged.connect(self._fit_to_screen)
            self._screen_signal_connected = True

    def resizeEvent(self, event: QtGui.QResizeEvent) -> None:
        super().resizeEvent(event)
        if hasattr(self, '_workspace_stack') and not getattr(self, '_fitting_geometry', False):
            self._fitting_geometry = True
            try:
                self._fit_to_screen()
            finally:
                self._fitting_geometry = False

    @QtCore.Slot(object)
    def _fit_to_screen(self, screen: QtGui.QScreen | None = None) -> None:
        active_screen = screen or self.screen() or QtWidgets.QApplication.primaryScreen()
        if active_screen is None:
            return

        available = active_screen.availableGeometry()
        max_w = min(self.width(), max(1, available.width() - 32))
        max_h = min(self.height(), max(1, available.height() - 32))
        min_size = self.minimumSize()
        if min_size.isValid() and not min_size.isNull():
            self.setMinimumSize(min(min_size.width(), max_w), min(min_size.height(), max_h))
        compact = (
            max_w < self._COMPACT_SPLIT_WIDTH
            or max_h < self._COMPACT_SPLIT_HEIGHT
        )
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

        if compact and self._workspace_stack.currentWidget() != self._compact_tabs:
            self._compact_tabs.addTab(self._control_panel, 'Controls')
            self._compact_tabs.addTab(self._telemetry_panel, 'Feedback and Plots')
            self._workspace_stack.setCurrentWidget(self._compact_tabs)
        elif not compact and self._workspace_stack.currentWidget() != self._splitter:
            while self._compact_tabs.count():
                self._compact_tabs.removeTab(0)
            self._splitter.addWidget(self._control_panel)
            self._splitter.addWidget(self._telemetry_panel)
            self._workspace_stack.setCurrentWidget(self._splitter)

        if self._splitter.orientation() == QtCore.Qt.Orientation.Horizontal:
            left = max(
                self._SPLIT_LEFT_WIDTH_MIN,
                min(self._SPLIT_LEFT_WIDTH_MAX, max_w // 3),
            )
            self._splitter.setSizes([left, max(1, max_w - left)])
        else:
            left = max(
                self._SPLIT_LEFT_WIDTH_MIN,
                min(self._SPLIT_LEFT_WIDTH_MAX, max_h // 3),
            )
            self._splitter.setSizes([left, max(1, max_h - left)])

        self.resize(
            min(self.width(), max_w),
            min(self.height(), max_h),
        )

        frame = self.frameGeometry()
        if frame.width() > available.width() or frame.height() > available.height():
            frame.moveCenter(available.center())
            self.move(frame.topLeft())
        else:
            new_x = max(available.left(), min(frame.left(), available.right() - frame.width() + 1))
            new_y = max(available.top(), min(frame.top(), available.bottom() - frame.height() + 1))
            frame.moveTopLeft(QtCore.QPoint(new_x, new_y))
            self.move(frame.topLeft())

    def _build_control_panel(self) -> QtWidgets.QWidget:
        scroll = QtWidgets.QScrollArea()
        scroll.setObjectName("panelScroll")
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QtWidgets.QFrame.Shape.NoFrame)
        scroll.setHorizontalScrollBarPolicy(QtCore.Qt.ScrollBarPolicy.ScrollBarAlwaysOff)

        left = QtWidgets.QFrame()
        left.setObjectName("card")
        left.setSizePolicy(
            QtWidgets.QSizePolicy.Policy.Preferred,
            QtWidgets.QSizePolicy.Policy.Expanding,
        )
        l = QtWidgets.QVBoxLayout(left)
        l.setContentsMargins(16, 16, 16, 16)
        l.setSpacing(10)

        conn_card = self._build_section_card("Connection")
        conn_layout = conn_card.layout()
        if conn_layout is None:
            return scroll

        motion_card = self._build_section_card("Motion Commands")
        motion_layout = motion_card.layout()
        if motion_layout is None:
            return scroll

        sample_card = self._build_section_card("Sample Transport")
        sample_layout = sample_card.layout()
        if sample_layout is None:
            return scroll

        conn_row = QtWidgets.QHBoxLayout()
        self.port_edit = QtWidgets.QLineEdit(self._default_port)
        self.port_edit.setPlaceholderText("COM port")
        self.connect_btn = QtWidgets.QPushButton("Connect")
        self.connect_btn.setObjectName("accent")
        self.disconnect_btn = QtWidgets.QPushButton("Disconnect")
        conn_row.addWidget(self.port_edit, 1)
        conn_row.addWidget(self.connect_btn)
        conn_row.addWidget(self.disconnect_btn)
        conn_layout.addLayout(conn_row)

        self.baud_combo = QtWidgets.QComboBox()
        self.baud_combo.addItems(["9600", "19200", "38400", "57600", "115200"])
        self.baud_combo.setCurrentText(str(self._default_baud))
        self._baud_row = QtWidgets.QHBoxLayout()
        self._baud_row.addWidget(QtWidgets.QLabel("Baud"))
        self._baud_row.addWidget(self.baud_combo)
        self._baud_row.addStretch(1)
        conn_layout.addLayout(self._baud_row)
        wiring = getattr(self._backend, "wiring_summary", "")
        if isinstance(wiring, str) and wiring:
            self.port_edit.setReadOnly(True)
            self.baud_combo.setCurrentText(str(self._backend._station_baud))
            self.baud_combo.setEnabled(False)
            wiring_label = QtWidgets.QLabel(wiring)
            wiring_label.setWordWrap(True)
            conn_layout.addWidget(wiring_label)

        self.axis_combo = QtWidgets.QComboBox()
        self.axis_combo.addItems(self._axes)
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
        motion_layout.addLayout(form)

        self.spin_rps = QtWidgets.QDoubleSpinBox()
        self.spin_rps.setRange(0.01, 20.0)
        self.spin_rps.setValue(1.0)
        self.spin_btn = QtWidgets.QPushButton("Spin Turning Motor")
        spin_form = QtWidgets.QFormLayout()
        spin_form.addRow("Spin (rps)", self.spin_rps)
        spin_form.addRow("", self.spin_btn)
        motion_layout.addLayout(spin_form)

        self.hole_spin = QtWidgets.QDoubleSpinBox()
        self.hole_spin.setRange(1.0, 101.0)
        self.hole_spin.setDecimals(1)
        self.goto_hole_btn = QtWidgets.QPushButton("Changer Motor To Hole")
        self.read_hole_btn = QtWidgets.QPushButton("Read Changer Hole")
        hole_form = QtWidgets.QFormLayout()
        hole_form.addRow("Target Hole", self.hole_spin)
        hole_form.addRow("", self.goto_hole_btn)
        hole_form.addRow("", self.read_hole_btn)
        motion_layout.addLayout(hole_form)

        self.home_top_btn = QtWidgets.QPushButton("Home Z to Top")
        self.home_xy_btn = QtWidgets.QPushButton("Home XY to Center")
        self.corner_btn = QtWidgets.QPushButton("Move XY to Corner")
        self.pickup_btn = QtWidgets.QPushButton("Sample Pickup")
        self.dropoff_btn = QtWidgets.QPushButton("Sample Dropoff")

        ops = QtWidgets.QGridLayout()
        ops.addWidget(self.home_top_btn, 0, 0)
        ops.addWidget(self.home_xy_btn, 0, 1)
        ops.addWidget(self.corner_btn, 1, 0)
        ops.addWidget(self.pickup_btn, 1, 1)
        ops.addWidget(self.dropoff_btn, 2, 0, 1, 2)
        sample_layout.addLayout(ops)
        l.addWidget(conn_card)
        l.addWidget(motion_card)
        l.addWidget(sample_card)
        l.addStretch(1)

        scroll.setWidget(left)
        return scroll

    def _build_telemetry_panel(self) -> QtWidgets.QScrollArea:
        panel = QtWidgets.QFrame()
        panel.setObjectName("card")
        panel.setSizePolicy(
            QtWidgets.QSizePolicy.Policy.Expanding,
            QtWidgets.QSizePolicy.Policy.Expanding,
        )
        r = QtWidgets.QVBoxLayout(panel)
        r.setContentsMargins(16, 16, 16, 16)
        r.setSpacing(10)

        legend = QtWidgets.QLabel("Live Controller Feedback")
        legend.setObjectName("dialogTitle")
        r.addWidget(legend)

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
        r.addLayout(trace_controls)

        values = QtWidgets.QGridLayout()
        values.setHorizontalSpacing(8)
        values.setVerticalSpacing(8)
        self._values_grid = values
        self.target_value = self._make_value("Input command --")
        self.actual_value = self._make_value("Output actual --")
        self.io_delta_value = self._make_value("Input-output delta --")
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

        self._console = QtWidgets.QPlainTextEdit()
        self._console.setObjectName("console")
        self._console.setReadOnly(True)
        self._console.setMaximumBlockCount(300)
        self._console.setMaximumHeight(125)
        self._console.setPlaceholderText("Connection, telemetry, and motion events")
        r.addWidget(self._console)

        self._refresh_connection_state()
        scroll = QtWidgets.QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QtWidgets.QFrame.Shape.NoFrame)
        scroll.setWidget(panel)
        return scroll

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

        self._values_grid = layout

    def _build_section_card(self, title: str) -> QtWidgets.QGroupBox:
        card = QtWidgets.QGroupBox(title)
        card.setObjectName("card")
        layout = QtWidgets.QVBoxLayout(card)
        layout.setContentsMargins(12, 12, 12, 12)
        layout.setSpacing(8)
        return card

    @staticmethod
    def _make_value(text: str) -> QtWidgets.QLabel:
        label = QtWidgets.QLabel(text)
        label.setObjectName("valuePill")
        label.setAlignment(QtCore.Qt.AlignmentFlag.AlignCenter)
        label.setSizePolicy(
            QtWidgets.QSizePolicy.Policy.Expanding,
            QtWidgets.QSizePolicy.Policy.Fixed,
        )
        return label

    def _build_plots(self, layout: QtWidgets.QVBoxLayout) -> None:
        self._curves: dict[str, object] = {}
        self._plots: list[object] = []
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
        graph.setMinimumHeight(650)

        position_plot = graph.addPlot(row=0, col=0, title="Input command vs output feedback")
        error_plot = graph.addPlot(row=1, col=0, title="Position error")
        velocity_plot = graph.addPlot(row=2, col=0, title="Velocity filters")
        delta_plot = graph.addPlot(row=3, col=0, title="Input-output residual")
        torque_plot = graph.addPlot(row=4, col=0, title="Torque feedback")
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
            pen=pg.mkPen("#d69f00", width=2),
            name="Target / input command",
        )
        self._curves["actual"] = position_plot.plot(
            pen=pg.mkPen("#7a0219", width=2),
            name="Actual / output feedback",
        )
        self._curves["error"] = error_plot.plot(pen=pg.mkPen("#d05f2d", width=2))
        self._curves["velocity_1"] = velocity_plot.plot(
            pen=pg.mkPen("#31566d", width=2),
            name="Filter 1",
        )
        self._curves["velocity_2"] = velocity_plot.plot(
            pen=pg.mkPen("#70a2b8", width=1),
            name="Filter 2",
        )
        self._curves["io_delta"] = delta_plot.plot(
            pen=pg.mkPen("#8f4ba8", width=2),
            name="Actual - Command",
        )
        self._curves["torque"] = torque_plot.plot(pen=pg.mkPen("#28755f", width=2))
        layout.addWidget(graph, 1)
        self._graph = graph

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
        if math.isnan(value_float):
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
        return self._axis_token(self._selected_axis())

    def _axis_matches(self, sample: object) -> bool:
        return self._extract_axis_token(sample) == self._axis_token(self._selected_axis())

    @classmethod
    def _extract_timestamp(cls, sample: object) -> float | None:
        for field in cls._TIMESTAMP_FIELD_ALIASES:
            value = cls._coerce_float(cls._read_sample_field(sample, field))
            if value is not None:
                return value / 1000.0 if field == "time_ms" else value
        return None

    def _extract_torque(self, sample: object) -> float | None:
        for field in self._TORQUE_FIELD_ALIASES:
            value = self._coerce_float(self._read_sample_field(sample, field))
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

    def _annotate_status(self, text: str) -> str:
        status = text.strip()
        if self._backend.simulated and "sim" not in status.lower():
            status = f"{status} (simulated)"
        return status

    def _set_status(self, text: str, level: str | None = None) -> None:
        annotated = self._annotate_status(text)
        normalized = annotated.lower()
        if level is None:
            if not self._backend_available:
                level = "unavailable"
            elif self._backend.simulated:
                level = "simulated"
            elif "error" in normalized or "failed" in normalized:
                level = "error"
            elif self._connected and self._backend.is_connected():
                level = "ready"
            else:
                level = "neutral"
        set_semantic_status(
            self._status,
            annotated,
            level,
            accessible_name="DC motor connection status",
        )

    def _selected_axis(self) -> str:
        current = self.axis_combo.currentText().strip()
        if current:
            return current
        if self._axes:
            return self._axes[0]
        return ""

    def _refresh_axis_combo(self, *, preserve_axis: str | None = None) -> None:
        self.axis_combo.blockSignals(True)
        self.axis_combo.clear()
        self.axis_combo.addItems(self._axes)
        selected = preserve_axis or self._selected_axis()
        if selected in self._axes:
            self.axis_combo.setCurrentText(selected)
        elif self.axis_combo.count():
            self.axis_combo.setCurrentIndex(0)
        self._telemetry_thread.set_axis(self._selected_axis())
        self.axis_combo.blockSignals(False)
        self._refresh_axis_summary()

    def _axis_changed(self, _name: str) -> None:
        self._telemetry_thread.set_axis(self._selected_axis())
        self._clear_traces()
        self._log(f"Telemetry axis: {self._selected_axis()}")
        self._refresh_axis_summary()

    def _start_command(self, label, action, *, requires_connected=True):
        if self._close_after_cleanup or self._command_thread is not None:
            self._log('Wait for the active motor command to settle.')
            return
        if requires_connected and not self._connected_required():
            return
        worker = MotorCommandWorker(self._backend, label, action)
        thread = QtCore.QThread(self)
        worker.moveToThread(thread)
        thread.started.connect(worker.run)
        worker.succeeded.connect(self._command_succeeded)
        worker.failed.connect(self._command_failed)
        worker.settled.connect(worker.deleteLater)
        worker.settled.connect(thread.quit)
        thread.finished.connect(self._command_settled)
        self._command_worker, self._command_thread = worker, thread
        self._log(label + ' running; Stop remains available.')
        self._refresh_connection_state()
        thread.start()

    def _stop_command(self):
        if self._command_worker is not None:
            self._command_worker.stop()
        elif getattr(self._backend, 'operation_active', False) is True:
            self._backend.request_stop()

    def _recover_stop(self):
        action = getattr(self._backend, 'recover_stop', None)
        if not callable(action):
            self._log('No-comm diagnostics have no physical stop recovery.')
            return
        self._start_command('Verify stopped recovery', action)

    def _command_succeeded(self, label, result):
        self._last_command_error = ''
        self._connected = bool(self._backend.is_connected())
        if label == 'Connect':
            self._axes = self._discover_axes()
            self._refresh_axis_combo(preserve_axis=self._selected_axis())
        self._log(label + ' finished after cleanup' + (': ' + str(result) if result is not None else '.'))

    def _command_failed(self, label, message):
        self._last_command_error = message
        self._connected = bool(self._backend.is_connected())
        self._log(label + ' failed or stopped: ' + message)
        self._set_status(message, 'error')

    @QtCore.Slot()
    def _command_settled(self):
        if self._command_thread is not None and self._command_thread.isRunning():
            self._command_settle_timer.start(10)
            return
        self._command_settle_timer.stop()
        if self._command_thread is not None:
            self._command_thread.deleteLater()
        self._command_thread = None
        self._command_worker = None
        self._refresh_connection_state()
        self._resume_close()

    @QtCore.Slot()
    def _resume_close(self):
        if self._close_after_cleanup:
            if (self._command_thread is not None
                    or self._telemetry_thread.isRunning() or getattr(self._backend, 'operation_active', False) is True):
                self._close_timer.start(100)
            else:
                self.close()

    def _connect(self) -> None:
        port, baud = self.port_edit.text().strip(), int(self.baud_combo.currentText())
        self._start_command('Connect', lambda: self._backend.connect(port, baud), requires_connected=False)

    def _disconnect(self) -> None:
        if self._close_after_cleanup:
            if self._backend.is_connected():
                self._backend.disconnect()
            self._connected = False
            return
        self._start_command('Disconnect', self._backend.disconnect, requires_connected=False)

    def _log(self, text: str) -> None:
        self._console.appendPlainText(text)
        self._set_status(text)

    def _connected_required(self) -> bool:
        if not self._connected or not self._backend.is_connected():
            self._log("Not connected to motor controller.")
            return False
        return True

    def _refresh_connection_state(self) -> None:
        self.connect_btn.setEnabled(self._backend_available and not self._connected)
        self.disconnect_btn.setEnabled(self._connected)
        self.port_edit.setEnabled(self._backend_available and not self._connected)
        self.baud_combo.setEnabled(self._backend_available and not self._connected)
        self.axis_combo.setEnabled(self._connected)
        for widget in (
            self.move_btn,
            self.spin_btn,
            self.goto_hole_btn,
            self.read_hole_btn,
            self.home_top_btn,
            self.home_xy_btn,
            self.corner_btn,
            self.pickup_btn,
            self.dropoff_btn,
        ):
            widget.setEnabled(self._connected)
        if not self._backend_available:
            self._set_status(self._backend.status(), "unavailable")
        elif self._connected and self._backend.is_connected():
            self._set_status(self._backend.status())
        else:
            self._set_status("Disconnected")
        active = getattr(self._backend, 'operation_active', False) is True
        recovery_required = getattr(self._backend, 'can_recover_pending', False) is True
        busy = self._command_thread is not None or self._close_after_cleanup or active
        if busy:
            for widget in (self.connect_btn, self.disconnect_btn, self.port_edit, self.baud_combo,
                           self.axis_combo, self.move_btn, self.spin_btn, self.goto_hole_btn,
                           self.read_hole_btn, self.home_top_btn, self.home_xy_btn, self.corner_btn,
                           self.pickup_btn, self.dropoff_btn):
                widget.setEnabled(False)
        if recovery_required:
            for widget in (self.move_btn, self.spin_btn, self.goto_hole_btn, self.home_top_btn,
                           self.home_xy_btn, self.corner_btn, self.pickup_btn, self.dropoff_btn):
                widget.setEnabled(False)
        self.stop_motion_btn.setEnabled(self._command_thread is not None or active)
        self.recover_stop_btn.setEnabled(self._connected and not busy)
        if self._last_command_error:
            self._set_status(self._last_command_error, 'error')
        elif recovery_required:
            self._set_status(self._backend.status(), 'error')
        self._refresh_axis_summary()

    def _refresh_axis_summary(self) -> None:
        if not self._connected:
            self._axis_summary.setText("Monitoring axis: -- (connect to enable)")
            return
        selected = self._selected_axis()
        if selected:
            self._axis_summary.setText(f"Monitoring axis: {selected}")
            return
        self._axis_summary.setText("Monitoring axis: --")

    def _move_axis(self) -> None:
        axis, target, speed = self._selected_axis(), self.pos_spin.value(), self.speed_spin.value()
        self._start_command('Move ' + axis, lambda: self._backend.move_motor(axis, target=target, speed=speed, wait_for_stop=True))

    def _spin_turning(self) -> None:
        self.axis_combo.setCurrentText('Turning')
        rps = self.spin_rps.value()
        self._start_command('Turning spin', lambda: self._backend.spin_turning(speed_rps=rps, duration_s=60.))

    def _goto_hole(self) -> None:
        hole = self.hole_spin.value()
        self.axis_combo.setCurrentText('Changer (X)')
        self._start_command('Go to changer hole', lambda: self._backend.goto_hole(hole=hole))

    def _read_hole(self) -> None:
        self.axis_combo.setCurrentText('Changer (X)')
        self._start_command('Read changer hole', lambda: self._backend.read_hole('Changer (X)'))

    def _home_to_top(self) -> None:
        self.axis_combo.setCurrentText('Up/Down')
        self._start_command('Home to top', self._backend.home_to_top)

    def _home_xy_to_center(self) -> None:
        self._start_command('Home XY to center', self._backend.home_xy_to_center)

    def _move_xy_to_corner(self) -> None:
        self._start_command('Move XY to corner', self._backend.move_xy_to_corner)

    def _sample_pickup(self) -> None:
        self._start_command('Sample pickup', self._backend.sample_pickup)

    def _sample_dropoff(self) -> None:
        self._start_command('Sample dropoff', lambda: self._backend.sample_dropoff(use_xy_table=True))

    @QtCore.Slot(object)
    def _on_telemetry(self, sample: object) -> None:
        if not self._axis_matches(sample):
            return
        timestamp = self._extract_timestamp(sample)
        if timestamp is None:
            timestamp = time.monotonic()
        target_position = self._extract_signal_position(sample, self._INPUT_SIGNAL_FIELD_ALIASES)
        actual_position = self._extract_signal_position(sample, self._OUTPUT_SIGNAL_FIELD_ALIASES)
        position_error = self._coerce_int(self._read_sample_field(sample, "position_error"))
        velocity_1 = self._coerce_int(self._read_sample_field(sample, "velocity_1"))
        velocity_2 = self._coerce_int(self._read_sample_field(sample, "velocity_2"))
        if target_position is None or actual_position is None:
            return

        if position_error is None:
            self.error_value.setText("Error --")
        else:
            self.error_value.setText(f"Error {position_error:+,}")

        if velocity_1 is None:
            self.velocity_value.setText("Velocity1 --")
        else:
            self.velocity_value.setText(f"Velocity1 {velocity_1:+,} SAV")

        if velocity_2 is None:
            self.velocity2_value.setText("Velocity2 --")
        else:
            self.velocity2_value.setText(f"Velocity2 {velocity_2:+,} SAV")

        self._trace_origin = self._trace_origin or timestamp
        elapsed = timestamp - self._trace_origin
        self._sample_tick += 1
        io_delta = int(actual_position - target_position)
        self.target_value.setText(f"Input command {target_position:,}")
        self.actual_value.setText(f"Output feedback {actual_position:,}")
        self.io_delta_value.setText(f"Input-output delta {io_delta:+,}")
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

        self._traces["time"].append(elapsed)
        self._traces["target"].append(float(target_position))
        self._traces["actual"].append(float(actual_position))
        self._traces["io_delta"].append(float(io_delta))
        if position_error is None:
            self._traces["error"].append(float("nan"))
        else:
            self._traces["error"].append(float(position_error))
        if velocity_1 is None:
            self._traces["velocity_1"].append(float("nan"))
        else:
            self._traces["velocity_1"].append(float(velocity_1))
        if velocity_2 is None:
            self._traces["velocity_2"].append(float("nan"))
        else:
            self._traces["velocity_2"].append(float(velocity_2))
        if self._sample_tick % self._LOG_EVERY_N_SAMPLES == 0:
            self._log(
                f"Live @{self._selected_axis()} | "
                f"in={target_position:,} out={actual_position:,}"
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
            curve.setData(times, list(trace), skipFiniteCheck=True)
        if latest > history and self._plots:
            for plot in self._plots:
                plot.setXRange(latest - history, latest, padding=0.01)

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
        self.actual_value.setText("Output feedback --")
        self.io_delta_value.setText("Input-output delta --")
        self.error_value.setText("Error --")
        self.velocity_value.setText("Velocity1 --")
        self.velocity2_value.setText("Velocity2 --")
        self.torque_value.setText("Torque -- (N/A)")

    def closeEvent(self, event: QtGui.QCloseEvent) -> None:
        self._close_after_cleanup = True
        self._plot_timer.stop()
        self._stop_command()
        self._telemetry_thread.stop_monitoring()
        if (self._command_thread is not None
                or self._telemetry_thread.isRunning() or getattr(self._backend, 'operation_active', False) is True):
            event.ignore()
            self._close_timer.start(100)
            return
        self._close_timer.stop()
        self._disconnect()
        super().closeEvent(event)
