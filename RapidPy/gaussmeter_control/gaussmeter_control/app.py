from __future__ import annotations

import csv
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
import sys
import time

import pyqtgraph as pg
from PySide6 import QtCore, QtGui, QtWidgets
from serial.tools import list_ports


def _bootstrap_common_imports() -> None:
    root = Path(__file__).resolve().parents[2]
    if str(root) not in sys.path:
        sys.path.insert(0, str(root))


_bootstrap_common_imports()
from rapidpy_common.gaussmeter import (  # noqa: E402
    AUTO_PORT,
    BASE_UNIT_LABELS,
    DEFAULT_MODE,
    GaussmeterClient,
    GaussmeterConnectionError,
    GaussmeterDriverError,
    GaussmeterError,
    GaussmeterReading,
    MODE_LABELS,
    fwbell_driver_status,
    gaussmeter_driver_status,
    serial_port_name_to_number,
)
from rapidpy_common.ui import apply_card_shadow, apply_liquid_glass_theme  # noqa: E402


RANGE_OPTIONS = (
    ("Auto", 4),
    ("Range 0", 0),
    ("Range 1", 1),
    ("Range 2", 2),
    ("Range 3", 3),
)

SAMPLE_EXPORT_HEADERS = (
    "index",
    "elapsed_s",
    "captured_at",
    "value",
    "raw_value",
    "units_label",
    "base_units_label",
    "mode_index",
    "mode_label",
    "units_index",
    "range_index",
    "driver_path",
)


@dataclass(slots=True)
class SessionSample:
    index: int
    elapsed_s: float
    captured_at: datetime
    reading: GaussmeterReading
    driver_path: str


class MainWindow(QtWidgets.QMainWindow):
    def __init__(self) -> None:
        super().__init__()
        self.setWindowTitle("RapidPy Gaussmeter Control")
        self.resize(1360, 860)
        self.setMinimumSize(1040, 720)
        compact_font = QtGui.QFont(self.font())
        compact_size = compact_font.pointSizeF()
        if compact_size > 0:
            compact_font.setPointSizeF(max(9.0, compact_size - 0.4))
            self.setFont(compact_font)
        self._client: GaussmeterClient | None = None
        self._poll_timer = QtCore.QTimer(self)
        self._poll_timer.timeout.connect(self._poll_once)
        self._sample_timer = QtCore.QTimer(self)
        self._sample_timer.timeout.connect(self._sample_session_step)
        self._session_samples: list[SessionSample] = []
        self._session_anchor_s: float | None = None
        self._sampling_active = False
        self._sampling_target_count = 0
        self._sampling_captured_count = 0
        self._sampling_alarm_enabled = False
        self._building = False
        self._build_ui()
        self._wire_events()
        self._refresh_ports()
        self._refresh_driver_status()
        self._set_connected_state(False)

    def _build_ui(self) -> None:
        root = QtWidgets.QWidget(self)
        self.setCentralWidget(root)
        layout = QtWidgets.QHBoxLayout(root)
        layout.setContentsMargins(16, 16, 16, 16)
        layout.setSpacing(14)

        splitter = QtWidgets.QSplitter(QtCore.Qt.Orientation.Horizontal)
        splitter.setChildrenCollapsible(False)
        splitter.setHandleWidth(8)
        layout.addWidget(splitter)

        left_scroll = QtWidgets.QScrollArea()
        left_scroll.setObjectName("panelScroll")
        left_scroll.setWidgetResizable(True)
        left_scroll.setFrameShape(QtWidgets.QFrame.Shape.NoFrame)
        left_scroll.setHorizontalScrollBarPolicy(QtCore.Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        left_scroll.setVerticalScrollBarPolicy(QtCore.Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        left_scroll.setSizePolicy(
            QtWidgets.QSizePolicy.Policy.Preferred,
            QtWidgets.QSizePolicy.Policy.Expanding,
        )

        left_host = QtWidgets.QWidget()
        left_host_layout = QtWidgets.QVBoxLayout(left_host)
        left_host_layout.setContentsMargins(6, 6, 12, 6)
        left_host_layout.addStretch(0)

        left_card = QtWidgets.QFrame()
        left_card.setObjectName("card")
        left_card.setMinimumWidth(410)
        left_card.setSizePolicy(
            QtWidgets.QSizePolicy.Policy.Preferred,
            QtWidgets.QSizePolicy.Policy.Expanding,
        )
        left = QtWidgets.QVBoxLayout(left_card)
        left.setContentsMargins(18, 18, 18, 18)
        left.setSpacing(10)

        title = QtWidgets.QLabel("Gaussmeter")
        title.setObjectName("title")
        subtitle = QtWidgets.QLabel(
            "Driver-backed control panel for RapidPy gaussmeters. USB auto mode will use the available gm0 or FW Bell backend."
        )
        subtitle.setObjectName("subtitle")
        subtitle.setWordWrap(True)
        left.addWidget(title)
        left.addWidget(subtitle)

        self.driver_status = QtWidgets.QPlainTextEdit()
        self.driver_status.setObjectName("valuePill")
        self.driver_status.setReadOnly(True)
        self.driver_status.setLineWrapMode(QtWidgets.QPlainTextEdit.LineWrapMode.NoWrap)
        self.driver_status.setVerticalScrollBarPolicy(QtCore.Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        self.driver_status.setHorizontalScrollBarPolicy(QtCore.Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        self.driver_status.setSizeAdjustPolicy(
            QtWidgets.QAbstractScrollArea.SizeAdjustPolicy.AdjustIgnored
        )
        self.driver_status.setSizePolicy(
            QtWidgets.QSizePolicy.Policy.Preferred,
            QtWidgets.QSizePolicy.Policy.Fixed,
        )
        self.driver_status.setFixedHeight(58)
        left.addWidget(self.driver_status)

        dll_row = QtWidgets.QHBoxLayout()
        self.dll_path_edit = QtWidgets.QLineEdit()
        self.dll_path_edit.setPlaceholderText("Optional driver DLL path")
        self.browse_dll_btn = QtWidgets.QPushButton("Browse DLL")
        dll_row.addWidget(self.dll_path_edit, stretch=1)
        dll_row.addWidget(self.browse_dll_btn)
        left.addLayout(dll_row)

        mode_group = QtWidgets.QGroupBox("Connection Mode")
        mode_layout = QtWidgets.QHBoxLayout(mode_group)
        self.manual_port_radio = QtWidgets.QRadioButton("RS232 / COM port")
        self.auto_port_radio = QtWidgets.QRadioButton("USB / driver auto")
        self.auto_port_radio.setChecked(True)
        mode_layout.addWidget(self.manual_port_radio)
        mode_layout.addWidget(self.auto_port_radio)
        left.addWidget(mode_group)

        port_row = QtWidgets.QHBoxLayout()
        self.port_combo = QtWidgets.QComboBox()
        self.port_combo.setEditable(True)
        self.port_combo.setInsertPolicy(QtWidgets.QComboBox.NoInsert)
        self.refresh_ports_btn = QtWidgets.QPushButton("Refresh Ports")
        port_row.addWidget(self.port_combo, stretch=1)
        port_row.addWidget(self.refresh_ports_btn)
        left.addLayout(port_row)

        grid = QtWidgets.QGridLayout()
        grid.addWidget(QtWidgets.QLabel("Poll interval (ms)"), 0, 0)
        self.poll_spin = QtWidgets.QSpinBox()
        self.poll_spin.setRange(100, 5000)
        self.poll_spin.setSingleStep(100)
        self.poll_spin.setValue(500)
        grid.addWidget(self.poll_spin, 0, 1)

        grid.addWidget(QtWidgets.QLabel("Measurement mode"), 1, 0)
        self.mode_combo = QtWidgets.QComboBox()
        for index, label in enumerate(MODE_LABELS):
            self.mode_combo.addItem(label, index)
        grid.addWidget(self.mode_combo, 1, 1)

        grid.addWidget(QtWidgets.QLabel("Units"), 2, 0)
        self.units_combo = QtWidgets.QComboBox()
        for index, label in enumerate(BASE_UNIT_LABELS):
            self.units_combo.addItem(label, index)
        grid.addWidget(self.units_combo, 2, 1)

        grid.addWidget(QtWidgets.QLabel("Range"), 3, 0)
        self.range_combo = QtWidgets.QComboBox()
        for label, value in RANGE_OPTIONS:
            self.range_combo.addItem(label, value)
        grid.addWidget(self.range_combo, 3, 1)
        left.addLayout(grid)

        button_row = QtWidgets.QHBoxLayout()
        self.connect_btn = QtWidgets.QPushButton("Connect")
        self.connect_btn.setObjectName("accent")
        self.refresh_btn = QtWidgets.QPushButton("Refresh Now")
        button_row.addWidget(self.connect_btn)
        button_row.addWidget(self.refresh_btn)
        left.addLayout(button_row)

        command_grid = QtWidgets.QGridLayout()
        self.auto_range_btn = QtWidgets.QPushButton("Auto Range")
        self.null_btn = QtWidgets.QPushButton("Auto Null")
        self.auto_zero_btn = QtWidgets.QPushButton("Auto Zero")
        self.reset_peak_btn = QtWidgets.QPushButton("Reset Peak")
        self.get_time_btn = QtWidgets.QPushButton("Get Time")
        self.set_time_btn = QtWidgets.QPushButton("Set System Time")
        buttons = (
            self.auto_range_btn,
            self.null_btn,
            self.auto_zero_btn,
            self.reset_peak_btn,
            self.get_time_btn,
            self.set_time_btn,
        )
        for index, button in enumerate(buttons):
            command_grid.addWidget(button, index // 2, index % 2)
        left.addLayout(command_grid)

        console_title = QtWidgets.QLabel("Console")
        console_title.setObjectName("subtitle")
        left.addWidget(console_title)
        self.console = QtWidgets.QPlainTextEdit()
        self.console.setObjectName("console")
        self.console.setReadOnly(True)
        self.console.setLineWrapMode(QtWidgets.QPlainTextEdit.LineWrapMode.NoWrap)
        self.console.setVerticalScrollBarPolicy(QtCore.Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        self.console.setHorizontalScrollBarPolicy(QtCore.Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        self.console.setSizeAdjustPolicy(
            QtWidgets.QAbstractScrollArea.SizeAdjustPolicy.AdjustIgnored
        )
        self.console.setSizePolicy(
            QtWidgets.QSizePolicy.Policy.Preferred,
            QtWidgets.QSizePolicy.Policy.Expanding,
        )
        self.console.setMinimumHeight(120)
        left.addWidget(self.console, stretch=1)
        left_host_layout.addWidget(left_card)
        left_scroll.setWidget(left_host)

        right_card = QtWidgets.QFrame()
        right_card.setObjectName("card")
        right_card.setSizePolicy(
            QtWidgets.QSizePolicy.Policy.Expanding,
            QtWidgets.QSizePolicy.Policy.Expanding,
        )
        right = QtWidgets.QVBoxLayout(right_card)
        right.setContentsMargins(18, 18, 18, 18)
        right.setSpacing(12)

        top_splitter = QtWidgets.QSplitter(QtCore.Qt.Orientation.Horizontal)
        top_splitter.setChildrenCollapsible(False)
        top_splitter.setHandleWidth(8)

        status_panel = QtWidgets.QFrame()
        status_panel.setObjectName("livePanel")
        status_panel.setSizePolicy(
            QtWidgets.QSizePolicy.Policy.Preferred,
            QtWidgets.QSizePolicy.Policy.Preferred,
        )
        status_layout = QtWidgets.QVBoxLayout(status_panel)
        status_layout.setContentsMargins(16, 16, 16, 16)
        status_layout.setSpacing(8)

        display_title = QtWidgets.QLabel("Live Reading")
        display_title.setObjectName("subtitle")
        status_layout.addWidget(display_title)

        self.reading_label = QtWidgets.QLabel("--")
        self.reading_label.setObjectName("readingDisplay")
        reading_font = QtGui.QFont(self.font())
        reading_font.setPointSize(34)
        reading_font.setBold(True)
        self.reading_label.setFont(reading_font)
        self.reading_label.setAlignment(QtCore.Qt.AlignCenter)
        self.reading_label.setMinimumHeight(60)
        status_layout.addWidget(self.reading_label)

        self.units_label = QtWidgets.QLabel("Units: -")
        self.units_label.setObjectName("valuePill")
        self.units_label.setMinimumHeight(26)
        status_layout.addWidget(self.units_label)

        self.meta_label = QtWidgets.QLabel("Mode: -\nRange: -\nInstrument time: -")
        self.meta_label.setObjectName("valuePill")
        self.meta_label.setWordWrap(True)
        self.meta_label.setAlignment(QtCore.Qt.AlignmentFlag.AlignLeft | QtCore.Qt.AlignmentFlag.AlignTop)
        self.meta_label.setMinimumHeight(56)
        status_layout.addWidget(self.meta_label)
        status_layout.addStretch(1)

        sampling_group = QtWidgets.QGroupBox("Sampling Session")
        sampling_font = QtGui.QFont(self.font())
        sampling_size = sampling_font.pointSizeF()
        if sampling_size > 0:
            sampling_font.setPointSizeF(max(8.6, sampling_size - 0.6))
        sampling_group.setFont(sampling_font)
        sampling_layout = QtWidgets.QGridLayout(sampling_group)
        sampling_layout.setContentsMargins(10, 16, 10, 10)
        sampling_layout.setHorizontalSpacing(12)
        sampling_layout.setVerticalSpacing(6)
        sampling_layout.setColumnStretch(1, 1)
        sampling_layout.setColumnStretch(3, 1)

        rate_label = QtWidgets.QLabel("Rate (ms)")
        rate_label.setFont(sampling_font)
        sampling_layout.addWidget(rate_label, 0, 0)
        self.sample_rate_spin = QtWidgets.QSpinBox()
        self.sample_rate_spin.setFont(sampling_font)
        self.sample_rate_spin.setRange(50, 60000)
        self.sample_rate_spin.setSingleStep(50)
        self.sample_rate_spin.setValue(500)
        self.sample_rate_spin.setMinimumHeight(26)
        self.sample_rate_spin.setMaximumWidth(140)
        sampling_layout.addWidget(self.sample_rate_spin, 0, 1)

        sample_count_label = QtWidgets.QLabel("Sample count")
        sample_count_label.setFont(sampling_font)
        sampling_layout.addWidget(sample_count_label, 0, 2)
        self.sample_count_spin = QtWidgets.QSpinBox()
        self.sample_count_spin.setFont(sampling_font)
        self.sample_count_spin.setRange(1, 100000)
        self.sample_count_spin.setValue(100)
        self.sample_count_spin.setMinimumHeight(26)
        self.sample_count_spin.setMaximumWidth(140)
        sampling_layout.addWidget(self.sample_count_spin, 0, 3)

        alarm_low_label = QtWidgets.QLabel("Alarm low")
        alarm_low_label.setFont(sampling_font)
        sampling_layout.addWidget(alarm_low_label, 1, 0)
        self.alarm_low_spin = QtWidgets.QDoubleSpinBox()
        self.alarm_low_spin.setFont(sampling_font)
        self.alarm_low_spin.setRange(-1_000_000_000.0, 1_000_000_000.0)
        self.alarm_low_spin.setDecimals(3)
        self.alarm_low_spin.setValue(-1.0)
        self.alarm_low_spin.setMinimumHeight(26)
        self.alarm_low_spin.setMaximumWidth(140)
        sampling_layout.addWidget(self.alarm_low_spin, 1, 1)

        alarm_high_label = QtWidgets.QLabel("Alarm high")
        alarm_high_label.setFont(sampling_font)
        sampling_layout.addWidget(alarm_high_label, 1, 2)
        self.alarm_high_spin = QtWidgets.QDoubleSpinBox()
        self.alarm_high_spin.setFont(sampling_font)
        self.alarm_high_spin.setRange(-1_000_000_000.0, 1_000_000_000.0)
        self.alarm_high_spin.setDecimals(3)
        self.alarm_high_spin.setValue(1.0)
        self.alarm_high_spin.setMinimumHeight(26)
        self.alarm_high_spin.setMaximumWidth(140)
        sampling_layout.addWidget(self.alarm_high_spin, 1, 3)

        self.sample_now_btn = QtWidgets.QPushButton("Sample Once Now")
        self.start_sampling_btn = QtWidgets.QPushButton("Start Sampling")
        self.sample_on_alarm_btn = QtWidgets.QPushButton("Sample On Alarm")
        self.clear_session_btn = QtWidgets.QPushButton("Clear Session")
        self.save_session_btn = QtWidgets.QPushButton("Save Session")
        for button in (
            self.sample_now_btn,
            self.start_sampling_btn,
            self.sample_on_alarm_btn,
            self.clear_session_btn,
            self.save_session_btn,
        ):
            button.setFont(sampling_font)
            button.setMinimumHeight(30)
        sampling_layout.addWidget(self.sample_now_btn, 2, 0, 1, 2)
        sampling_layout.addWidget(self.start_sampling_btn, 2, 2, 1, 2)
        sampling_layout.addWidget(self.sample_on_alarm_btn, 3, 0, 1, 2)
        sampling_layout.addWidget(self.clear_session_btn, 3, 2, 1, 2)
        sampling_layout.addWidget(self.save_session_btn, 4, 0, 1, 4)

        self.session_duration_label = QtWidgets.QLabel()
        self.session_duration_label.setObjectName("valuePill")
        self.session_duration_label.setFont(sampling_font)
        self.session_duration_label.setMinimumHeight(26)
        self.session_status = QtWidgets.QPlainTextEdit()
        self.session_status.setObjectName("statusPill")
        self.session_status.setFont(sampling_font)
        self.session_status.setReadOnly(True)
        self.session_status.setLineWrapMode(QtWidgets.QPlainTextEdit.LineWrapMode.WidgetWidth)
        self.session_status.setHorizontalScrollBarPolicy(QtCore.Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        self.session_status.setVerticalScrollBarPolicy(QtCore.Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        self.session_status.setSizeAdjustPolicy(
            QtWidgets.QAbstractScrollArea.SizeAdjustPolicy.AdjustIgnored
        )
        self.session_status.setFrameShape(QtWidgets.QFrame.Shape.NoFrame)
        self.session_status.setFixedHeight(72)
        self.session_status.setSizePolicy(
            QtWidgets.QSizePolicy.Policy.Expanding,
            QtWidgets.QSizePolicy.Policy.Fixed,
        )
        sampling_layout.addWidget(self.session_duration_label, 5, 0, 1, 4)
        sampling_layout.addWidget(self.session_status, 6, 0, 1, 4)

        sampling_scroll = QtWidgets.QScrollArea()
        sampling_scroll.setObjectName("panelScroll")
        sampling_scroll.setWidgetResizable(True)
        sampling_scroll.setFrameShape(QtWidgets.QFrame.Shape.NoFrame)
        sampling_scroll.setHorizontalScrollBarPolicy(QtCore.Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        sampling_scroll.setVerticalScrollBarPolicy(QtCore.Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        sampling_scroll.setSizePolicy(
            QtWidgets.QSizePolicy.Policy.Expanding,
            QtWidgets.QSizePolicy.Policy.Preferred,
        )
        sampling_scroll.setWidget(sampling_group)

        top_splitter.addWidget(status_panel)
        top_splitter.addWidget(sampling_scroll)
        top_splitter.setStretchFactor(0, 0)
        top_splitter.setStretchFactor(1, 1)
        top_splitter.setSizes([240, 600])
        right.addWidget(top_splitter)

        content_splitter = QtWidgets.QSplitter(QtCore.Qt.Orientation.Vertical)
        content_splitter.setChildrenCollapsible(False)
        content_splitter.setHandleWidth(8)

        plot_section = QtWidgets.QWidget()
        plot_layout = QtWidgets.QVBoxLayout(plot_section)
        plot_layout.setContentsMargins(0, 0, 0, 0)
        plot_layout.setSpacing(8)
        plot_title = QtWidgets.QLabel("Session Plot")
        plot_title.setObjectName("subtitle")
        plot_layout.addWidget(plot_title)
        self.session_plot = pg.PlotWidget()
        self.session_plot.setMinimumHeight(200)
        self.session_plot.showGrid(x=True, y=True, alpha=0.25)
        self.session_plot.setMenuEnabled(False)
        self.session_plot.setLabel("bottom", "Elapsed (s)")
        self.session_plot.setLabel("left", "Field")
        self.session_plot.getPlotItem().hideButtons()
        self.session_curve = self.session_plot.plot(
            [],
            [],
            pen=pg.mkPen("#0f766e", width=2),
            symbol="o",
            symbolSize=5,
            symbolBrush="#0f766e",
            symbolPen="#0f766e",
        )
        plot_layout.addWidget(self.session_plot, stretch=1)

        details_section = QtWidgets.QWidget()
        details_layout = QtWidgets.QVBoxLayout(details_section)
        details_layout.setContentsMargins(0, 0, 0, 0)
        details_layout.setSpacing(8)
        details_title = QtWidgets.QLabel("Instrument Details")
        details_title.setObjectName("subtitle")
        details_layout.addWidget(details_title)

        self.details = QtWidgets.QPlainTextEdit()
        self.details.setObjectName("console")
        self.details.setReadOnly(True)
        self.details.setMinimumHeight(80)
        details_layout.addWidget(self.details)

        content_splitter.addWidget(plot_section)
        content_splitter.addWidget(details_section)
        content_splitter.setStretchFactor(0, 3)
        content_splitter.setStretchFactor(1, 1)
        content_splitter.setSizes([360, 120])
        right.addWidget(content_splitter, stretch=1)

        splitter.addWidget(left_scroll)
        splitter.addWidget(right_card)
        splitter.setStretchFactor(0, 0)
        splitter.setStretchFactor(1, 1)
        splitter.setSizes([440, 900])

        apply_card_shadow(left_card)
        apply_card_shadow(right_card)
        apply_card_shadow(status_panel)

    def _wire_events(self) -> None:
        self.browse_dll_btn.clicked.connect(self._browse_for_dll)
        self.dll_path_edit.textChanged.connect(self._refresh_driver_status)
        self.refresh_ports_btn.clicked.connect(self._refresh_ports)
        self.manual_port_radio.toggled.connect(self._update_port_controls)
        self.auto_port_radio.toggled.connect(self._update_port_controls)
        self.manual_port_radio.toggled.connect(self._refresh_driver_status)
        self.auto_port_radio.toggled.connect(self._refresh_driver_status)
        self.connect_btn.clicked.connect(self._toggle_connection)
        self.refresh_btn.clicked.connect(lambda: self._poll_once(wait_for_new_data=False))
        self.mode_combo.currentIndexChanged.connect(self._push_mode)
        self.units_combo.currentIndexChanged.connect(self._push_units)
        self.range_combo.currentIndexChanged.connect(self._push_range)
        self.auto_range_btn.clicked.connect(self._set_auto_range)
        self.null_btn.clicked.connect(self._run_null)
        self.auto_zero_btn.clicked.connect(self._run_auto_zero)
        self.reset_peak_btn.clicked.connect(self._run_reset_peak)
        self.get_time_btn.clicked.connect(self._show_instrument_time)
        self.set_time_btn.clicked.connect(self._set_system_time)
        self.sample_now_btn.clicked.connect(self._sample_now)
        self.start_sampling_btn.clicked.connect(self._toggle_sampling)
        self.sample_on_alarm_btn.clicked.connect(self._toggle_alarm_sampling)
        self.clear_session_btn.clicked.connect(self._clear_session)
        self.save_session_btn.clicked.connect(self._save_session)
        self.sample_rate_spin.valueChanged.connect(self._update_sampling_controls)
        self.sample_count_spin.valueChanged.connect(self._update_sampling_controls)
        self.alarm_low_spin.valueChanged.connect(self._update_sampling_controls)
        self.alarm_high_spin.valueChanged.connect(self._update_sampling_controls)

    def _browse_for_dll(self) -> None:
        file_path, _ = QtWidgets.QFileDialog.getOpenFileName(
            self,
            "Select gaussmeter DLL",
            self.dll_path_edit.text() or str(Path.cwd()),
            "Dynamic Libraries (*.dll)",
        )
        if file_path:
            self.dll_path_edit.setText(file_path)

    def _driver_search_paths(self) -> list[str] | None:
        text = self.dll_path_edit.text().strip()
        if not text:
            return None
        return [text]

    def _refresh_driver_status(self) -> None:
        if self.auto_port_radio.isChecked():
            available, detail = fwbell_driver_status(self._driver_search_paths())
        else:
            available, detail = gaussmeter_driver_status(self._driver_search_paths())
        if available:
            self.driver_status.setPlainText(f"Driver ready: {detail}")
        else:
            self.driver_status.setPlainText(f"Driver unavailable: {detail}")
        self.connect_btn.setEnabled(available)
        if not available and self._client is None:
            self._set_reading(None)

    def _refresh_ports(self) -> None:
        current_text = self.port_combo.currentText().strip()
        self.port_combo.blockSignals(True)
        self.port_combo.clear()
        for info in sorted(list_ports.comports(), key=lambda item: item.device.upper()):
            label = info.device
            if info.description:
                label = f"{info.device} - {info.description}"
            self.port_combo.addItem(label, info.device)
        if current_text:
            self.port_combo.setEditText(current_text)
        elif self.port_combo.count():
            self.port_combo.setCurrentIndex(0)
        else:
            self.port_combo.setEditText("COM1")
        self.port_combo.blockSignals(False)
        self._update_port_controls()

    def _update_port_controls(self) -> None:
        manual = self.manual_port_radio.isChecked()
        self.port_combo.setEnabled(manual)
        self.refresh_ports_btn.setEnabled(manual)

    def _selected_port(self) -> int:
        if self.auto_port_radio.isChecked():
            return AUTO_PORT
        data = self.port_combo.currentData()
        if data is None:
            return serial_port_name_to_number(self.port_combo.currentText())
        return serial_port_name_to_number(str(data))

    def _toggle_connection(self) -> None:
        if self._client is not None and self._client.connected:
            self._disconnect()
            return

        try:
            port = self._selected_port()
            self._client = GaussmeterClient(
                port=port,
                mode=DEFAULT_MODE,
                dll_search_paths=self._driver_search_paths(),
            )
            self._client.connect(timeout_s=8.0)
            self._apply_requested_settings()
            self._set_connected_state(True)
            mode_text = "USB auto" if port == AUTO_PORT else f"COM{port}"
            self._log(f"Connected using {mode_text}. Driver: {self._client.dll_path}")
            self._poll_once(wait_for_new_data=False)
            self._poll_timer.start(self.poll_spin.value())
        except (ValueError, GaussmeterError) as exc:
            self._disconnect(log_message=False)
            QtWidgets.QMessageBox.warning(self, "Gaussmeter", str(exc))

    def _disconnect(self, *, log_message: bool = True) -> None:
        self._stop_sampling_session(log_message=None, restart_poll=False)
        self._poll_timer.stop()
        if self._client is not None:
            self._client.disconnect()
        self._client = None
        self._set_connected_state(False)
        self._set_reading(None)
        if log_message:
            self._log("Disconnected")

    def _set_connected_state(self, connected: bool) -> None:
        self.connect_btn.setText("Disconnect" if connected else "Connect")
        for button in (
            self.refresh_btn,
            self.auto_range_btn,
            self.null_btn,
            self.auto_zero_btn,
            self.reset_peak_btn,
            self.get_time_btn,
            self.set_time_btn,
        ):
            button.setEnabled(connected)
        self._update_sampling_controls()

    def _apply_requested_settings(self) -> None:
        self._push_mode()
        self._push_units()
        self._push_range()

    def _push_mode(self) -> None:
        if self._client is None or not self._client.connected:
            return
        self._client.set_mode(int(self.mode_combo.currentData()))

    def _push_units(self) -> None:
        if self._client is None or not self._client.connected:
            return
        self._client.set_units(int(self.units_combo.currentData()))

    def _push_range(self) -> None:
        if self._client is None or not self._client.connected:
            return
        self._client.set_range(int(self.range_combo.currentData()))

    def _set_auto_range(self) -> None:
        self.range_combo.setCurrentIndex(0)
        if self._client is not None and self._client.connected:
            self._client.set_range(4)
            self._poll_once(wait_for_new_data=False)
            self._log("Requested auto range")

    def _run_null(self) -> None:
        self._run_command("Auto null", lambda client: client.null())

    def _run_auto_zero(self) -> None:
        self._run_command("Auto zero", lambda client: client.auto_zero())

    def _run_reset_peak(self) -> None:
        self._run_command("Reset peak", lambda client: client.reset_peak())

    def _show_instrument_time(self) -> None:
        if self._client is None or not self._client.connected:
            return
        timestamp = self._client.read(wait_for_new_data=False).timestamp
        if timestamp is None:
            self._log("Instrument time is unavailable")
        else:
            self._log(f"Instrument time: {timestamp:%Y-%m-%d %H:%M:%S}")
            self._poll_once(wait_for_new_data=False)

    def _set_system_time(self) -> None:
        self._run_command("Set system time", lambda client: client.set_system_time())

    def _run_command(self, label: str, action) -> None:
        if self._client is None or not self._client.connected:
            return
        try:
            action(self._client)
            self._log(f"{label} complete")
            self._poll_once(wait_for_new_data=False)
        except GaussmeterError as exc:
            QtWidgets.QMessageBox.warning(self, "Gaussmeter", str(exc))

    def _read_current_reading(self, *, wait_for_new_data: bool = True) -> GaussmeterReading | None:
        if self._client is None or not self._client.connected:
            return None
        try:
            timeout_s = max(1.0, (self.poll_spin.value() / 1000.0) + 0.25)
            reading = self._client.read(
                wait_for_new_data=wait_for_new_data,
                sample_timeout_s=timeout_s,
            )
            self._set_reading(reading)
            return reading
        except GaussmeterConnectionError as exc:
            self._disconnect(log_message=False)
            QtWidgets.QMessageBox.warning(self, "Gaussmeter", str(exc))
        except GaussmeterError as exc:
            self._log(f"Read failed: {exc}")
        return None

    def _poll_once(self, *, wait_for_new_data: bool = True) -> None:
        self._read_current_reading(wait_for_new_data=wait_for_new_data)

    def _alarm_window_bounds(self) -> tuple[float, float]:
        low = self.alarm_low_spin.value()
        high = self.alarm_high_spin.value()
        if low <= high:
            return low, high
        return high, low

    def _update_sampling_controls(self) -> None:
        connected = self._client is not None and self._client.connected
        self.sample_now_btn.setEnabled(connected and not self._sampling_active)
        self.start_sampling_btn.setEnabled(connected or self._sampling_active)
        self.start_sampling_btn.setText("Stop Sampling" if self._sampling_active else "Start Sampling")
        self.sample_on_alarm_btn.setEnabled(not self._sampling_active)
        self.sample_on_alarm_btn.setText("Turn Alarm Off" if self._sampling_alarm_enabled else "Sample On Alarm")
        for widget in (
            self.sample_rate_spin,
            self.sample_count_spin,
            self.alarm_low_spin,
            self.alarm_high_spin,
        ):
            widget.setEnabled(not self._sampling_active)
        has_samples = bool(self._session_samples)
        self.clear_session_btn.setEnabled(has_samples and not self._sampling_active)
        self.save_session_btn.setEnabled(has_samples and not self._sampling_active)

        estimated_duration_s = self.sample_rate_spin.value() * self.sample_count_spin.value() / 1000.0
        self.session_duration_label.setText(f"Estimated duration: {estimated_duration_s:.1f} s")

        low, high = self._alarm_window_bounds()
        lines = [
            f"Stored samples: {len(self._session_samples)}",
            f"Run plan: {self.sample_count_spin.value()} samples @ {self.sample_rate_spin.value()} ms",
            (
                f"Alarm window: {low:.6g} to {high:.6g}"
                if self._sampling_alarm_enabled
                else "Alarm window: off"
            ),
        ]
        if self._sampling_active:
            lines.append(f"Run progress: {self._sampling_captured_count}/{self._sampling_target_count}")
        elif self._session_samples:
            lines.append(f"Session span: {self._session_samples[-1].elapsed_s:.2f} s")
        self.session_status.setPlainText("\n".join(lines))

    def _toggle_alarm_sampling(self) -> None:
        self._sampling_alarm_enabled = not self._sampling_alarm_enabled
        low, high = self._alarm_window_bounds()
        if self._sampling_alarm_enabled:
            self._log(f"Alarm-window capture enabled for values between {low:.6g} and {high:.6g}.")
        else:
            self._log("Alarm-window capture disabled.")
        self._update_sampling_controls()

    def _toggle_sampling(self) -> None:
        if self._sampling_active:
            self._stop_sampling_session(log_message="Sampling stopped by user.")
            return
        if self._client is None or not self._client.connected:
            return
        self._sampling_target_count = int(self.sample_count_spin.value())
        self._sampling_captured_count = 0
        self._sampling_active = True
        self._poll_timer.stop()
        self._sample_timer.start(int(self.sample_rate_spin.value()))
        self._log(
            f"Sampling run started for {self._sampling_target_count} samples at {self.sample_rate_spin.value()} ms."
        )
        self._update_sampling_controls()
        self._sample_session_step()

    def _stop_sampling_session(self, *, log_message: str | None, restart_poll: bool = True) -> None:
        was_active = self._sampling_active
        self._sample_timer.stop()
        self._sampling_active = False
        self._sampling_target_count = 0
        self._sampling_captured_count = 0
        self._update_sampling_controls()
        if log_message and was_active:
            self._log(log_message)
        if restart_poll and self._client is not None and self._client.connected and not self._poll_timer.isActive():
            self._poll_timer.start(self.poll_spin.value())

    def _sample_should_be_recorded(self, reading: GaussmeterReading) -> bool:
        if not self._sampling_alarm_enabled:
            return True
        low, high = self._alarm_window_bounds()
        return low <= reading.value <= high

    def _append_session_sample(self, reading: GaussmeterReading) -> None:
        now_s = time.monotonic()
        if self._session_anchor_s is None:
            self._session_anchor_s = now_s
        elapsed_s = now_s - self._session_anchor_s
        sample = SessionSample(
            index=len(self._session_samples) + 1,
            elapsed_s=elapsed_s,
            captured_at=datetime.now(),
            reading=reading,
            driver_path=str(self._client.dll_path) if self._client is not None else "-",
        )
        self._session_samples.append(sample)
        self._update_session_plot()
        self._update_sampling_controls()

    def _sample_session_step(self) -> None:
        if not self._sampling_active:
            return
        if self._client is None or not self._client.connected:
            self._stop_sampling_session(log_message="Sampling stopped: gaussmeter disconnected.", restart_poll=False)
            return
        reading = self._read_current_reading(wait_for_new_data=False)
        if reading is None:
            if self._client is None or not self._client.connected:
                self._stop_sampling_session(log_message="Sampling stopped: gaussmeter disconnected.", restart_poll=False)
            return
        if not self._sample_should_be_recorded(reading):
            self._update_sampling_controls()
            return
        self._append_session_sample(reading)
        self._sampling_captured_count += 1
        if self._sampling_captured_count >= self._sampling_target_count:
            self._stop_sampling_session(
                log_message=f"Sampling run complete with {self._sampling_captured_count} captured samples."
            )
        else:
            self._update_sampling_controls()

    def _sample_now(self) -> None:
        reading = self._read_current_reading(wait_for_new_data=False)
        if reading is None:
            return
        self._append_session_sample(reading)
        self._log(f"Captured sample {len(self._session_samples)} at {reading.value:.6g} {reading.units_label}.")

    def _clear_session(self) -> None:
        if self._sampling_active:
            self._stop_sampling_session(log_message="Sampling stopped before clearing the session.")
        self._session_samples.clear()
        self._session_anchor_s = None
        self._update_session_plot()
        self._update_sampling_controls()
        self._log("Cleared session data.")

    def _save_session(self) -> None:
        if not self._session_samples:
            return
        default_name = f"gaussmeter-session-{datetime.now():%Y%m%d-%H%M%S}.csv"
        file_path, _ = QtWidgets.QFileDialog.getSaveFileName(
            self,
            "Save gaussmeter session",
            str(Path.cwd() / default_name),
            "CSV Files (*.csv)",
        )
        if not file_path:
            return
        with open(file_path, "w", newline="", encoding="utf-8") as handle:
            writer = csv.writer(handle)
            writer.writerow(SAMPLE_EXPORT_HEADERS)
            for sample in self._session_samples:
                writer.writerow(
                    [
                        sample.index,
                        f"{sample.elapsed_s:.6f}",
                        sample.captured_at.isoformat(timespec="seconds"),
                        sample.reading.value,
                        sample.reading.raw_value,
                        sample.reading.units_label,
                        sample.reading.base_units_label,
                        sample.reading.mode_index,
                        sample.reading.mode_label,
                        sample.reading.units_index,
                        sample.reading.range_index,
                        sample.driver_path,
                    ]
                )
        self._log(f"Saved {len(self._session_samples)} session samples to {file_path}.")

    def _update_session_plot(self) -> None:
        if not self._session_samples:
            self.session_curve.setData([], [])
            self.session_plot.setLabel("left", "Field")
            return
        x_values = [sample.elapsed_s for sample in self._session_samples]
        y_values = [sample.reading.value for sample in self._session_samples]
        self.session_curve.setData(x_values, y_values)
        self.session_plot.setLabel("left", f"Field ({self._session_samples[-1].reading.units_label})")

    def _set_reading(self, reading: GaussmeterReading | None) -> None:
        if reading is None:
            self.reading_label.setText("--")
            self.units_label.setText("Units: -")
            self.meta_label.setText("Mode: -\nRange: -\nInstrument time: -")
            self.details.clear()
            return

        self.reading_label.setText(f"{reading.value:.6g}")
        self.units_label.setText(f"Units: {reading.units_label} (base {reading.base_units_label})")
        timestamp_text = "-"
        if reading.timestamp is not None:
            timestamp_text = reading.timestamp.strftime("%Y-%m-%d %H:%M:%S")
        self.meta_label.setText(
            f"Mode: {reading.mode_label}\n"
            f"Range: {reading.range_index}\n"
            f"Instrument time: {timestamp_text}"
        )
        details = [
            f"Displayed value: {reading.value}",
            f"Raw value: {reading.raw_value}",
            f"Mode index: {reading.mode_index}",
            f"Units index: {reading.units_index}",
            f"Range index: {reading.range_index}",
            f"Driver mode: {'USB auto' if self.auto_port_radio.isChecked() else 'Manual COM'}",
            f"Driver path: {self._client.dll_path if self._client is not None else '-'}",
            f"Stored session samples: {len(self._session_samples)}",
            f"Sampling active: {'yes' if self._sampling_active else 'no'}",
        ]
        self.details.setPlainText("\n".join(details))

    def _log(self, message: str) -> None:
        self.console.appendPlainText(message)

    def closeEvent(self, event: QtGui.QCloseEvent) -> None:
        self._disconnect(log_message=False)
        super().closeEvent(event)


def main() -> int:
    pg.setConfigOptions(antialias=True)
    app = QtWidgets.QApplication(sys.argv)
    apply_liquid_glass_theme(app)
    window = MainWindow()
    window.show()
    return app.exec()