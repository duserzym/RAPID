from __future__ import annotations

from pathlib import Path
import sys

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


class MainWindow(QtWidgets.QMainWindow):
    def __init__(self) -> None:
        super().__init__()
        self.setWindowTitle("RapidPy Gaussmeter Control")
        self.resize(1180, 760)
        self._client: GaussmeterClient | None = None
        self._poll_timer = QtCore.QTimer(self)
        self._poll_timer.timeout.connect(self._poll_once)
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

        left_card = QtWidgets.QFrame()
        left_card.setObjectName("card")
        left_card.setMinimumWidth(410)
        left = QtWidgets.QVBoxLayout(left_card)
        left.setContentsMargins(18, 18, 18, 18)
        left.setSpacing(12)

        title = QtWidgets.QLabel("Gaussmeter")
        title.setObjectName("title")
        subtitle = QtWidgets.QLabel(
            "Driver-backed control panel for RapidPy gaussmeters. USB auto mode will use the available gm0 or FW Bell backend."
        )
        subtitle.setObjectName("subtitle")
        subtitle.setWordWrap(True)
        left.addWidget(title)
        left.addWidget(subtitle)

        self.driver_status = QtWidgets.QLabel()
        self.driver_status.setObjectName("valuePill")
        self.driver_status.setWordWrap(True)
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
        left.addWidget(self.console, stretch=1)

        right_card = QtWidgets.QFrame()
        right_card.setObjectName("card")
        right = QtWidgets.QVBoxLayout(right_card)
        right.setContentsMargins(18, 18, 18, 18)
        right.setSpacing(12)

        display_title = QtWidgets.QLabel("Live Reading")
        display_title.setObjectName("subtitle")
        right.addWidget(display_title)

        self.reading_label = QtWidgets.QLabel("--")
        reading_font = QtGui.QFont(self.font())
        reading_font.setPointSize(34)
        reading_font.setBold(True)
        self.reading_label.setFont(reading_font)
        self.reading_label.setAlignment(QtCore.Qt.AlignCenter)
        self.reading_label.setMinimumHeight(88)
        right.addWidget(self.reading_label)

        self.units_label = QtWidgets.QLabel("Units: -")
        self.units_label.setObjectName("valuePill")
        right.addWidget(self.units_label)

        self.meta_label = QtWidgets.QLabel("Mode: -\nRange: -\nInstrument time: -")
        self.meta_label.setObjectName("valuePill")
        self.meta_label.setWordWrap(True)
        right.addWidget(self.meta_label)

        details_title = QtWidgets.QLabel("Session Details")
        details_title.setObjectName("subtitle")
        right.addWidget(details_title)

        self.details = QtWidgets.QPlainTextEdit()
        self.details.setObjectName("console")
        self.details.setReadOnly(True)
        right.addWidget(self.details, stretch=1)

        layout.addWidget(left_card)
        layout.addWidget(right_card, stretch=1)

        apply_card_shadow(left_card)
        apply_card_shadow(right_card)

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
            self.driver_status.setText(f"Driver ready: {detail}")
        else:
            self.driver_status.setText(f"Driver unavailable: {detail}")
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

    def _poll_once(self, *, wait_for_new_data: bool = True) -> None:
        if self._client is None or not self._client.connected:
            return
        try:
            timeout_s = max(1.0, (self.poll_spin.value() / 1000.0) + 0.25)
            reading = self._client.read(
                wait_for_new_data=wait_for_new_data,
                sample_timeout_s=timeout_s,
            )
            self._set_reading(reading)
        except GaussmeterConnectionError as exc:
            self._disconnect(log_message=False)
            QtWidgets.QMessageBox.warning(self, "Gaussmeter", str(exc))
        except GaussmeterError as exc:
            self._log(f"Read failed: {exc}")

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
        ]
        self.details.setPlainText("\n".join(details))

    def _log(self, message: str) -> None:
        self.console.appendPlainText(message)

    def closeEvent(self, event: QtGui.QCloseEvent) -> None:
        self._disconnect(log_message=False)
        super().closeEvent(event)


def main() -> int:
    app = QtWidgets.QApplication(sys.argv)
    apply_liquid_glass_theme(app)
    window = MainWindow()
    window.show()
    return app.exec()