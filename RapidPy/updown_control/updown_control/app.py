from __future__ import annotations

import json
import sys
from dataclasses import asdict, dataclass
from pathlib import Path

from PySide6 import QtCore, QtGui, QtWidgets
from serial.tools import list_ports


def _bootstrap_common_imports() -> None:
    root = Path(__file__).resolve().parents[2]
    if str(root) not in sys.path:
        sys.path.insert(0, str(root))


_bootstrap_common_imports()
from rapidpy_common.hardware import MotorAxisConfig, MotorSerialClient  # noqa: E402
from rapidpy_common.ui import apply_card_shadow, apply_liquid_glass_theme  # noqa: E402


SETTINGS_PATH = Path.home() / ".rapidpy_updown_settings.json"


@dataclass(slots=True)
class UpDownSettings:
    com_port: str = ""
    target_unit: str = "mm"
    min_raw_count: int = -250_000
    max_raw_count: int = 250_000
    sample_pickup_raw: int = 425_000
    sample_dropoff_raw: int = 582_500
    susceptibility_meter_raw: int = 20_000


def load_settings(path: Path = SETTINGS_PATH) -> UpDownSettings:
    if not path.exists():
        return UpDownSettings()
    try:
        payload = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return UpDownSettings()

    defaults = asdict(UpDownSettings())
    for key in defaults:
        if key in payload:
            defaults[key] = payload[key]
    return UpDownSettings(**defaults)


def save_settings(settings: UpDownSettings, path: Path = SETTINGS_PATH) -> None:
    try:
        path.write_text(json.dumps(asdict(settings), indent=2, sort_keys=True), encoding="utf-8")
    except OSError:
        return


class MainWindow(QtWidgets.QMainWindow):
    def __init__(self) -> None:
        super().__init__()
        self.setWindowTitle("RapidPy Up/Down Control")
        self.resize(1040, 700)
        self._assets_dir = Path(__file__).resolve().parent.parent / "assets"
        self._icon_png = self._assets_dir / "updown_icon.png"
        if self._icon_png.exists():
            self.setWindowIcon(QtGui.QIcon(str(self._icon_png)))
        self.motor = MotorSerialClient()
        self.axis = MotorAxisConfig(name="UpDown", motor_id=3, address=3)
        self._settings = load_settings()
        self._build_ui()
        self._apply_local_style()
        self._refresh_ports()
        self._load_settings_into_widgets()

    def _build_ui(self) -> None:
        root = QtWidgets.QWidget(self)
        self.setCentralWidget(root)
        layout = QtWidgets.QHBoxLayout(root)
        layout.setContentsMargins(16, 16, 16, 16)
        layout.setSpacing(14)

        card = QtWidgets.QFrame()
        card.setObjectName("card")
        c = QtWidgets.QVBoxLayout(card)
        c.setContentsMargins(18, 18, 18, 18)

        title = QtWidgets.QLabel("Up/Down Axis")
        title.setObjectName("title")
        subtitle = QtWidgets.QLabel("Height positioning with bounds, units, and saved raw-count presets")
        subtitle.setObjectName("subtitle")
        c.addWidget(title)
        c.addWidget(subtitle)

        self.com_port_combo = QtWidgets.QComboBox()
        self.com_port_combo.setEditable(True)
        self.refresh_ports_btn = QtWidgets.QPushButton("Refresh Ports")
        self.connect_btn = QtWidgets.QPushButton("Connect")
        self.connect_btn.setObjectName("accent")
        self.disconnect_btn = QtWidgets.QPushButton("Disconnect")

        row_conn = QtWidgets.QHBoxLayout()
        row_conn.addWidget(self.com_port_combo, stretch=1)
        row_conn.addWidget(self.refresh_ports_btn)
        row_conn.addWidget(self.connect_btn)
        row_conn.addWidget(self.disconnect_btn)
        c.addLayout(row_conn)

        self.target_edit = QtWidgets.QLineEdit("0")
        self.target_unit_combo = QtWidgets.QComboBox()
        self.target_unit_combo.addItems(["mm", "cm", "raw count"])
        self.target_raw_preview = QtWidgets.QLabel("Target raw count: 0")
        self.target_raw_preview.setObjectName("valuePill")

        self.min_raw_spin = QtWidgets.QSpinBox()
        self.min_raw_spin.setRange(-2_000_000_000, 2_000_000_000)
        self.max_raw_spin = QtWidgets.QSpinBox()
        self.max_raw_spin.setRange(-2_000_000_000, 2_000_000_000)

        self.move_btn = QtWidgets.QPushButton("Change Height")
        self.move_btn.setObjectName("accent")

        self.home_top_btn = QtWidgets.QPushButton("Home To Top")
        self.pickup_btn = QtWidgets.QPushButton("Sample Pickup")
        self.dropoff_btn = QtWidgets.QPushButton("Sample Dropoff")
        self.susceptibility_btn = QtWidgets.QPushButton("Susceptibility Meter")

        self.pickup_raw_edit = QtWidgets.QLineEdit()
        self.dropoff_raw_edit = QtWidgets.QLineEdit()
        self.susceptibility_raw_edit = QtWidgets.QLineEdit()

        self.load_settings_btn = QtWidgets.QPushButton("Load Settings")
        self.save_settings_btn = QtWidgets.QPushButton("Save Settings")
        self.save_as_btn = QtWidgets.QPushButton("Save Settings As")

        form = QtWidgets.QFormLayout()
        form.addRow("Target Height", self.target_edit)
        form.addRow("Target Unit", self.target_unit_combo)
        form.addRow("Min Allowed (raw)", self.min_raw_spin)
        form.addRow("Max Allowed (raw)", self.max_raw_spin)
        form.addRow("", self.target_raw_preview)
        form.addRow("", self.move_btn)
        c.addLayout(form)

        presets_box = QtWidgets.QGroupBox("Preset Positions (raw motor counts)")
        presets = QtWidgets.QGridLayout(presets_box)
        presets.addWidget(self.home_top_btn, 0, 0)
        presets.addWidget(QtWidgets.QLabel("switch-based"), 0, 1)

        presets.addWidget(self.pickup_btn, 1, 0)
        presets.addWidget(self.pickup_raw_edit, 1, 1)

        presets.addWidget(self.dropoff_btn, 2, 0)
        presets.addWidget(self.dropoff_raw_edit, 2, 1)

        presets.addWidget(self.susceptibility_btn, 3, 0)
        presets.addWidget(self.susceptibility_raw_edit, 3, 1)
        c.addWidget(presets_box)

        settings_row = QtWidgets.QHBoxLayout()
        settings_row.addWidget(self.load_settings_btn)
        settings_row.addWidget(self.save_settings_btn)
        settings_row.addWidget(self.save_as_btn)
        c.addLayout(settings_row)

        self.status = QtWidgets.QLabel("Disconnected")
        self.status.setObjectName("valuePill")
        c.addWidget(self.status)

        self.console = QtWidgets.QPlainTextEdit()
        self.console.setReadOnly(True)
        self.console.setObjectName("console")
        c.addWidget(self.console, stretch=1)

        layout.addWidget(card)
        apply_card_shadow(card)

        self.refresh_ports_btn.clicked.connect(self._refresh_ports)
        self.connect_btn.clicked.connect(self._connect)
        self.disconnect_btn.clicked.connect(self._disconnect)
        self.move_btn.clicked.connect(self._move_to_height)
        self.home_top_btn.clicked.connect(self._home_to_top)
        self.pickup_btn.clicked.connect(lambda: self._run_preset("SamplePickup", self.pickup_raw_edit))
        self.dropoff_btn.clicked.connect(lambda: self._run_preset("SampleDropoff", self.dropoff_raw_edit))
        self.susceptibility_btn.clicked.connect(
            lambda: self._run_preset("SusceptibilityMeter", self.susceptibility_raw_edit)
        )
        self.target_edit.textChanged.connect(self._update_target_preview)
        self.target_unit_combo.currentTextChanged.connect(self._update_target_preview)
        self.min_raw_spin.valueChanged.connect(self._update_target_preview)
        self.max_raw_spin.valueChanged.connect(self._update_target_preview)
        self.load_settings_btn.clicked.connect(self._load_settings_from_file)
        self.save_settings_btn.clicked.connect(self._save_settings)
        self.save_as_btn.clicked.connect(self._save_settings_as)

    def _apply_local_style(self) -> None:
        self.setStyleSheet(
            """
            QLineEdit, QSpinBox, QDoubleSpinBox, QComboBox {
                font-size: 16px;
                min-height: 34px;
                selection-background-color: #7A0219;
                selection-color: #ffffff;
                border-radius: 12px;
            }
            QLineEdit:focus, QSpinBox:focus, QDoubleSpinBox:focus, QComboBox:focus {
                border: 1px solid #7A0219;
            }
            """
        )

    def _refresh_ports(self) -> None:
        current = self.com_port_combo.currentText().strip()
        self.com_port_combo.clear()
        ports = sorted(p.device for p in list_ports.comports())
        self.com_port_combo.addItems(ports)

        preferred = self._settings.com_port or current
        if preferred:
            if preferred not in ports:
                self.com_port_combo.addItem(preferred)
            self.com_port_combo.setCurrentText(preferred)

    def _load_settings_into_widgets(self) -> None:
        if self._settings.com_port:
            if self.com_port_combo.findText(self._settings.com_port) < 0:
                self.com_port_combo.addItem(self._settings.com_port)
            self.com_port_combo.setCurrentText(self._settings.com_port)

        if self._settings.target_unit in {"mm", "cm", "raw count"}:
            self.target_unit_combo.setCurrentText(self._settings.target_unit)
        self.min_raw_spin.setValue(self._settings.min_raw_count)
        self.max_raw_spin.setValue(self._settings.max_raw_count)
        self.pickup_raw_edit.setText(str(self._settings.sample_pickup_raw))
        self.dropoff_raw_edit.setText(str(self._settings.sample_dropoff_raw))
        self.susceptibility_raw_edit.setText(str(self._settings.susceptibility_meter_raw))
        self._update_target_preview()

    def _log(self, message: str) -> None:
        self.console.appendPlainText(message)
        self.status.setText(message)

    def _target_value_to_raw(self) -> int:
        unit = self.target_unit_combo.currentText().strip().lower()
        text = self.target_edit.text().strip()
        if not text:
            raise ValueError("Target height is empty.")

        if unit == "raw count":
            return int(round(float(text)))
        if unit == "mm":
            return int(round(float(text) * 1000.0))
        if unit == "cm":
            return int(round(float(text) * 10000.0))
        raise ValueError(f"Unsupported unit {unit!r}")

    def _update_target_preview(self) -> None:
        try:
            raw = self._target_value_to_raw()
            in_bounds = self._is_within_bounds(raw)
            suffix = "(in bounds)" if in_bounds else "(out of bounds)"
            self.target_raw_preview.setText(f"Target raw count: {raw} {suffix}")
        except Exception:
            self.target_raw_preview.setText("Target raw count: invalid")

    def _is_within_bounds(self, raw: int) -> bool:
        low = self.min_raw_spin.value()
        high = self.max_raw_spin.value()
        if low > high:
            return False
        return low <= raw <= high

    def _validate_bounds(self) -> bool:
        if self.min_raw_spin.value() > self.max_raw_spin.value():
            QtWidgets.QMessageBox.warning(self, "Invalid Bounds", "Minimum raw count cannot exceed maximum raw count.")
            return False
        return True

    def _connect(self) -> None:
        port = self.com_port_combo.currentText().strip()
        if not port:
            QtWidgets.QMessageBox.warning(self, "Missing Port", "Select a serial port first.")
            return
        try:
            self.motor.connect(port, baudrate=57600)
        except Exception as exc:
            QtWidgets.QMessageBox.critical(self, "Connection Error", str(exc))
            return
        self._log(f"Connected up/down motor on {port}")

    def _disconnect(self) -> None:
        self.motor.disconnect()
        self._log("Disconnected")

    def _move_to_height(self) -> None:
        if not self._validate_bounds():
            return
        try:
            target = self._target_value_to_raw()
        except ValueError as exc:
            QtWidgets.QMessageBox.warning(self, "Invalid Target", str(exc))
            return

        if not self._is_within_bounds(target):
            QtWidgets.QMessageBox.warning(
                self,
                "Target Out Of Bounds",
                f"Target raw count {target} is outside allowed range [{self.min_raw_spin.value()}, {self.max_raw_spin.value()}].",
            )
            return

        if not self.motor.is_connected:
            QtWidgets.QMessageBox.warning(self, "Not Connected", "Connect motor serial before movement commands.")
            return
        try:
            result = self.motor.updown_move(self.axis, target=target, speed_index=0, wait_for_stop=True)
            final_mm = result.final_position / 1000.0
            self._log(
                f"Move command sent: target_raw={target}, final_raw={result.final_position}, "
                f"final_mm={final_mm:.3f}, success={result.success}"
            )
        except Exception as exc:
            QtWidgets.QMessageBox.warning(self, "Move Error", str(exc))

    def _preset_raw_from_editor(self, editor: QtWidgets.QLineEdit) -> int:
        text = editor.text().strip()
        if not text:
            raise ValueError("Preset raw count is empty.")
        return int(text)

    def _run_preset(self, name: str, editor: QtWidgets.QLineEdit) -> None:
        if not self._validate_bounds():
            return
        try:
            target = self._preset_raw_from_editor(editor)
        except ValueError as exc:
            QtWidgets.QMessageBox.warning(self, "Invalid Preset", str(exc))
            return
        if not self._is_within_bounds(target):
            QtWidgets.QMessageBox.warning(
                self,
                "Preset Out Of Bounds",
                f"Preset raw count {target} is outside allowed range [{self.min_raw_spin.value()}, {self.max_raw_spin.value()}].",
            )
            return
        if not self.motor.is_connected:
            QtWidgets.QMessageBox.warning(self, "Not Connected", "Connect motor serial before preset moves.")
            return
        try:
            result = self.motor.updown_move(self.axis, target=target, speed_index=0, wait_for_stop=True)
            self._log(
                f"Preset {name}: target_raw={target}, final_raw={result.final_position}, success={result.success}"
            )
        except Exception as exc:
            QtWidgets.QMessageBox.warning(self, f"{name} Error", str(exc))

    def _home_to_top(self) -> None:
        if not self.motor.is_connected:
            QtWidgets.QMessageBox.warning(self, "Not Connected", "Connect motor serial before homing commands.")
            return
        try:
            result = self.motor.home_to_top(self.axis)
            self._log(f"HomeToTop complete: final={result.final_position}, success={result.success}")
        except Exception as exc:
            QtWidgets.QMessageBox.warning(self, "HomeToTop Error", str(exc))

    def _settings_from_widgets(self) -> UpDownSettings:
        return UpDownSettings(
            com_port=self.com_port_combo.currentText().strip(),
            target_unit=self.target_unit_combo.currentText().strip(),
            min_raw_count=int(self.min_raw_spin.value()),
            max_raw_count=int(self.max_raw_spin.value()),
            sample_pickup_raw=int(self.pickup_raw_edit.text().strip() or "0"),
            sample_dropoff_raw=int(self.dropoff_raw_edit.text().strip() or "0"),
            susceptibility_meter_raw=int(self.susceptibility_raw_edit.text().strip() or "0"),
        )

    def _save_settings(self) -> None:
        try:
            self._settings = self._settings_from_widgets()
        except ValueError as exc:
            QtWidgets.QMessageBox.warning(self, "Invalid Settings", str(exc))
            return
        if self._settings.min_raw_count > self._settings.max_raw_count:
            QtWidgets.QMessageBox.warning(self, "Invalid Settings", "Minimum raw count cannot exceed maximum raw count.")
            return
        save_settings(self._settings, SETTINGS_PATH)
        self._log(f"Saved settings: {SETTINGS_PATH}")

    def _save_settings_as(self) -> None:
        try:
            candidate = self._settings_from_widgets()
        except ValueError as exc:
            QtWidgets.QMessageBox.warning(self, "Invalid Settings", str(exc))
            return
        path, _ = QtWidgets.QFileDialog.getSaveFileName(
            self,
            "Save Up/Down Settings",
            str(SETTINGS_PATH),
            "JSON Files (*.json)",
        )
        if not path:
            return
        save_settings(candidate, Path(path))
        self._log(f"Saved settings as: {path}")

    def _load_settings_from_file(self) -> None:
        path, _ = QtWidgets.QFileDialog.getOpenFileName(
            self,
            "Load Up/Down Settings",
            str(SETTINGS_PATH),
            "JSON Files (*.json)",
        )
        if not path:
            return
        self._settings = load_settings(Path(path))
        self._load_settings_into_widgets()
        self._log(f"Loaded settings: {path}")

    def closeEvent(self, event: QtGui.QCloseEvent) -> None:  # noqa: N802
        try:
            self._settings = self._settings_from_widgets()
            save_settings(self._settings, SETTINGS_PATH)
        except Exception:
            pass
        super().closeEvent(event)


def main() -> int:
    app = QtWidgets.QApplication(sys.argv)
    apply_liquid_glass_theme(app)
    window = MainWindow()
    window.show()
    return app.exec()
