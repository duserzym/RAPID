"""RapidPy changer XY control application.

This module keeps the VB6 operator-facing raw motor velocity inputs while also
showing physically meaningful speed estimates.

Important velocity clarification:
- VB6 changer move commands use controller raw velocity units, not direct stage
    position counts per second.
- The raw-to-position scale is derived from the loaded INI as
    TurningMotor1rps / abs(TurningMotorFullRotation).
- In Paleomag_v3.INI that is 16,000,000 / 8,000 = 2,000 raw units per
    position-count-per-second.
- A raw command of 1,000,000 therefore means about 500 position counts/s,
    not 1,000,000 position counts/s.
- The UI combines that scale with the loaded XYTable cup spacing and
    UpDownMotor1cm so operators can keep VB6-compatible raw inputs while seeing
    estimated cm/s values that match the physical stage.
"""

from __future__ import annotations

import configparser
from datetime import datetime
import json
import shutil
import sys
import time
from dataclasses import asdict, dataclass, field
from pathlib import Path

import serial
from serial.tools import list_ports
from PySide6 import QtCore, QtGui, QtWidgets


def _bootstrap_common_imports() -> None:
    root = Path(__file__).resolve().parents[2]
    if str(root) not in sys.path:
        sys.path.insert(0, str(root))


_bootstrap_common_imports()
from rapidpy_common.hardware import HardwareError, MotorAxisConfig, MotorControllerConfig, MotorSerialClient, MoveResult  # noqa: E402
from rapidpy_common.ui import apply_card_shadow, apply_liquid_glass_theme, set_app_icon  # noqa: E402


CONFIG_PATH = Path.home() / ".rapidpy_changer_xy_control.json"
DEFAULT_SETTINGS_PATH = Path(__file__).resolve().parents[3] / "VB6" / "settings" / "Paleomag_v3.INI"
TRAY_GRID_COLS = 10
TRAY_GRID_ROWS = 10
TRAY_CENTER_DROP_HOLE = 46
TRAY_ROW_OFFSET = 0.5
TRAY_ROW_STEP = 0.88
TRAY_CUP_DIAMETER_IN = 1.0
TRAY_CUP_PITCH_IN = 1.32
TRAY_CENTER_DROP_DIAMETER_IN = 1.12
MOTION_DEFAULTS = MotorControllerConfig()


def tray_logical_hole_positions(slot_min: int, slot_max: int) -> dict[int, QtCore.QPointF]:
    positions: dict[int, QtCore.QPointF] = {}
    for hole in range(slot_min, slot_max + 1):
        row = (hole - 1) // TRAY_GRID_COLS
        col = (hole - 1) % TRAY_GRID_COLS
        x = col + 0.5 + (TRAY_ROW_OFFSET if row % 2 else 0.0)
        y = row * TRAY_ROW_STEP + 0.5
        positions[hole] = QtCore.QPointF(x, y)
    return positions


def tray_logical_bounds(slot_min: int, slot_max: int) -> tuple[float, float, float, float]:
    positions = tray_logical_hole_positions(slot_min, slot_max)
    xs = [point.x() for point in positions.values()]
    ys = [point.y() for point in positions.values()]
    return min(xs), max(xs), min(ys), max(ys)


@dataclass(slots=True)
class HoleCalibration:
    hole: int
    x: int | None = None
    y: int | None = None


@dataclass(slots=True)
class AppConfig:
    x_port: str = "COM4"
    y_port: str = "COM5"
    updown_port: str = "COM6"
    sensor_port: str = ""
    updown_disabled: bool = False
    settings_path: str = str(DEFAULT_SETTINGS_PATH)
    selected_hole: int = 1
    calibrations: list[HoleCalibration] = field(default_factory=list)


@dataclass(slots=True)
class SettingsProfile:
    path: Path
    slot_min: int
    slot_max: int
    motion_defaults: MotorControllerConfig
    x_axis: MotorAxisConfig
    y_axis: MotorAxisConfig
    updown_axis: MotorAxisConfig
    xy_home_x: int = 0
    xy_home_y: int = 0
    updown_motor_1cm: int = 0
    calibrations: list[HoleCalibration] = field(default_factory=list)


def _parse_ini_number(raw_value: str | None, default: float) -> float:
    if raw_value is None:
        return default
    text = raw_value.strip()
    if not text:
        return default
    try:
        return float(text)
    except ValueError:
        return default


def _parse_ini_int(raw_value: str | None, default: int) -> int:
    return int(round(_parse_ini_number(raw_value, float(default))))


def _new_settings_config() -> configparser.ConfigParser:
    config = configparser.ConfigParser(interpolation=None)
    config.optionxform = str
    return config


def _settings_json_payload_from_config(config: configparser.ConfigParser) -> dict[str, object]:
    return {
        "sections": [
            {
                "name": section,
                "entries": [{"key": key, "value": value} for key, value in config.items(section)],
            }
            for section in config.sections()
        ]
    }


def _settings_config_from_json_payload(payload: object) -> configparser.ConfigParser:
    if not isinstance(payload, dict):
        raise ValueError("JSON root must be an object.")

    if "sections" in payload:
        section_payloads = payload.get("sections")
        if not isinstance(section_payloads, list):
            raise ValueError("The 'sections' field must be a list.")
    else:
        section_payloads = [
            {
                "name": name,
                "entries": [{"key": key, "value": value} for key, value in values.items()],
            }
            for name, values in payload.items()
            if isinstance(values, dict)
        ]

    config = _new_settings_config()
    seen_sections: set[str] = set()
    for section_payload in section_payloads:
        if not isinstance(section_payload, dict):
            raise ValueError("Each section must be an object.")
        name = str(section_payload.get("name", "")).strip()
        if not name:
            raise ValueError("Each section needs a non-empty name.")
        if name in seen_sections:
            raise ValueError(f"Duplicate section name: {name}")
        seen_sections.add(name)
        config.add_section(name)
        entry_payloads = section_payload.get("entries", [])
        if not isinstance(entry_payloads, list):
            raise ValueError(f"Section {name} has an invalid entries list.")
        seen_keys: set[str] = set()
        for entry_payload in entry_payloads:
            if not isinstance(entry_payload, dict):
                raise ValueError(f"Section {name} has a non-object entry.")
            key = str(entry_payload.get("key", "")).strip()
            if not key:
                raise ValueError(f"Section {name} contains an empty key.")
            if key in seen_keys:
                raise ValueError(f"Section {name} contains duplicate key {key}.")
            seen_keys.add(key)
            value = entry_payload.get("value", "")
            config[name][key] = "" if value is None else str(value)
    return config


def _load_settings_config(settings_path: Path) -> configparser.ConfigParser:
    if not settings_path.exists():
        raise FileNotFoundError(f"Settings file not found: {settings_path}")
    if settings_path.suffix.lower() == ".json":
        payload = json.loads(settings_path.read_text(encoding="utf-8"))
        return _settings_config_from_json_payload(payload)
    config = _new_settings_config()
    config.read(settings_path, encoding="utf-8")
    return config


def _settings_profile_from_config(settings_path: Path, config: configparser.ConfigParser) -> SettingsProfile:

    sample_section = config["SampleChanger"] if config.has_section("SampleChanger") else {}
    motor_section = config["SteppingMotor"] if config.has_section("SteppingMotor") else {}
    program_section = config["MotorPrograms"] if config.has_section("MotorPrograms") else {}
    table_section = config["XYTable"] if config.has_section("XYTable") else {}

    slot_min = _parse_ini_int(sample_section.get("SlotMin"), MOTION_DEFAULTS.slot_min)
    slot_max = _parse_ini_int(sample_section.get("SlotMax"), MOTION_DEFAULTS.slot_max - 1)
    motion_defaults = MotorControllerConfig(
        slot_min=slot_min,
        slot_max=slot_max,
        one_step=_parse_ini_int(sample_section.get("OneStep"), MOTION_DEFAULTS.one_step),
        sample_hole_alignment_offset=_parse_ini_int(
            motor_section.get("SampleHoleAlignmentOffset"),
            MOTION_DEFAULTS.sample_hole_alignment_offset,
        ),
        changer_speed=_parse_ini_int(motor_section.get("ChangerSpeed"), MOTION_DEFAULTS.changer_speed),
        turner_speed=_parse_ini_int(motor_section.get("TurnerSpeed"), MOTION_DEFAULTS.turner_speed),
        turning_motor_full_rotation=_parse_ini_int(
            motor_section.get("TurningMotorFullRotation"),
            MOTION_DEFAULTS.turning_motor_full_rotation,
        ),
        turning_motor_1rps=_parse_ini_int(motor_section.get("TurningMotor1rps"), MOTION_DEFAULTS.turning_motor_1rps),
        lift_speed_slow=_parse_ini_int(motor_section.get("LiftSpeedSlow"), MOTION_DEFAULTS.lift_speed_slow),
        lift_speed_normal=_parse_ini_int(motor_section.get("LiftSpeedNormal"), MOTION_DEFAULTS.lift_speed_normal),
        lift_speed_fast=_parse_ini_int(motor_section.get("LiftSpeedFast"), MOTION_DEFAULTS.lift_speed_fast),
        lift_acceleration=_parse_ini_int(motor_section.get("LiftAcceleration"), MOTION_DEFAULTS.lift_acceleration),
        meas_pos=_parse_ini_int(motor_section.get("MeasPos"), MOTION_DEFAULTS.meas_pos),
        sample_bottom=_parse_ini_int(motor_section.get("SampleBottom"), MOTION_DEFAULTS.sample_bottom),
        sample_height=_parse_ini_int(motor_section.get("SampleTop"), MOTION_DEFAULTS.sample_bottom + MOTION_DEFAULTS.sample_height)
        - _parse_ini_int(motor_section.get("SampleBottom"), MOTION_DEFAULTS.sample_bottom),
        updown_torque_factor=_parse_ini_int(
            motor_section.get("UpDownTorqueFactor"),
            MOTION_DEFAULTS.updown_torque_factor,
        ),
        pickup_torque_throttle=float(
            _parse_ini_number(motor_section.get("PickupTorqueThrottle"), MOTION_DEFAULTS.pickup_torque_throttle)
        ),
        xy_neg_homing_distance=MOTION_DEFAULTS.xy_neg_homing_distance,
        xy_pos_homing_distance=MOTION_DEFAULTS.xy_pos_homing_distance,
        xy_corner_distance=MOTION_DEFAULTS.xy_corner_distance,
    )

    changer_address = _parse_ini_int(program_section.get("MotorIDChanger"), 16)
    changer_y_address = _parse_ini_int(program_section.get("MotorIDChangerY"), 16)
    updown_address = _parse_ini_int(program_section.get("MotorIDUpDown"), 16)

    calibrations: list[HoleCalibration] = []
    for hole in range(slot_min, slot_max + 1):
        x_key = f"XY{hole}X"
        y_key = f"XY{hole}Y"
        raw_x = table_section.get(x_key, "").strip()
        raw_y = table_section.get(y_key, "").strip()
        if not raw_x or not raw_y:
            continue
        try:
            calibrations.append(HoleCalibration(hole=hole, x=int(raw_x), y=int(raw_y)))
        except ValueError:
            continue

    return SettingsProfile(
        path=settings_path,
        slot_min=slot_min,
        slot_max=slot_max,
        motion_defaults=motion_defaults,
        x_axis=MotorAxisConfig("ChangerX", 1, changer_address),
        y_axis=MotorAxisConfig("ChangerY", 4, changer_y_address),
        updown_axis=MotorAxisConfig("UpDown", 3, updown_address),
        xy_home_x=_parse_ini_int(table_section.get("XYHomeX"), 0),
        xy_home_y=_parse_ini_int(table_section.get("XYHomeY"), 0),
        updown_motor_1cm=_parse_ini_int(motor_section.get("UpDownMotor1cm"), 0),
        calibrations=calibrations,
    )


def _load_settings_profile(settings_path: Path) -> SettingsProfile:
    return _settings_profile_from_config(settings_path, _load_settings_config(settings_path))


class SerialProbe:
    def __init__(self) -> None:
        self._serial: serial.Serial | None = None

    @property
    def is_connected(self) -> bool:
        return self._serial is not None and self._serial.is_open

    def connect(self, port: str, baudrate: int = 9600, timeout: float = 0.35) -> None:
        self.disconnect()
        self._serial = serial.Serial(
            port=port,
            baudrate=baudrate,
            bytesize=serial.EIGHTBITS,
            parity=serial.PARITY_NONE,
            stopbits=serial.STOPBITS_ONE,
            timeout=timeout,
            write_timeout=timeout,
        )

    def disconnect(self) -> None:
        if self._serial is not None:
            try:
                self._serial.close()
            finally:
                self._serial = None


class ChangerStageController:
    def __init__(self, profile: SettingsProfile) -> None:
        self.motion_defaults = profile.motion_defaults
        self.x_axis = profile.x_axis
        self.y_axis = profile.y_axis
        self.updown_axis = profile.updown_axis
        self._motor_clients: dict[str, MotorSerialClient] = {}
        self._axis_ports: dict[str, str] = {"x": "", "y": "", "updown": ""}
        self._sensor_probe = SerialProbe()
        self._sensor_port = ""

    def apply_settings_profile(self, profile: SettingsProfile) -> None:
        self.motion_defaults = profile.motion_defaults
        self.x_axis = profile.x_axis
        self.y_axis = profile.y_axis
        self.updown_axis = profile.updown_axis
        for client in self._motor_clients.values():
            client.config = profile.motion_defaults

    def shutdown(self) -> None:
        for client in list(self._motor_clients.values()):
            client.disconnect()
        self._motor_clients.clear()
        self._axis_ports = {"x": "", "y": "", "updown": ""}
        self.disconnect_sensor()

    def axis_port(self, key: str) -> str:
        return self._axis_ports.get(key, "")

    def connect_axis(self, key: str, port: str) -> None:
        port = port.strip()
        if not port:
            raise HardwareError("A COM port is required.")
        existing_port = self._axis_ports.get(key, "")
        if existing_port == port and port in self._motor_clients and self._motor_clients[port].is_connected:
            return
        if existing_port and existing_port != port:
            self.disconnect_axis(key)
        client = self._motor_clients.get(port)
        if client is None or not client.is_connected:
            client = MotorSerialClient(config=self.motion_defaults)
            client.connect(port, baudrate=57600)
            self._motor_clients[port] = client
        self._axis_ports[key] = port

    def disconnect_axis(self, key: str) -> None:
        port = self._axis_ports.get(key, "")
        if not port:
            return
        self._axis_ports[key] = ""
        if port and port not in self._axis_ports.values():
            client = self._motor_clients.pop(port, None)
            if client is not None:
                client.disconnect()

    def connect_sensor(self, port: str) -> None:
        port = port.strip()
        if not port:
            raise HardwareError("A COM port is required.")
        self._sensor_probe.connect(port)
        self._sensor_port = port

    def disconnect_sensor(self) -> None:
        self._sensor_probe.disconnect()
        self._sensor_port = ""

    @property
    def sensor_port(self) -> str:
        return self._sensor_port

    @property
    def sensor_connected(self) -> bool:
        return self._sensor_probe.is_connected

    def axis_connected(self, key: str) -> bool:
        port = self._axis_ports.get(key, "")
        return bool(port and port in self._motor_clients and self._motor_clients[port].is_connected)

    def _client_for_axis(self, key: str) -> MotorSerialClient:
        port = self._axis_ports.get(key, "")
        if not port:
            raise HardwareError(f"Axis {key.upper()} is not assigned to a COM port.")
        client = self._motor_clients.get(port)
        if client is None or not client.is_connected:
            raise HardwareError(f"Axis {key.upper()} is not connected.")
        return client

    def _axis(self, key: str) -> MotorAxisConfig:
        if key == "x":
            return self.x_axis
        if key == "y":
            return self.y_axis
        return self.updown_axis

    def read_axis_position(self, key: str) -> int:
        client = self._client_for_axis(key)
        return client.read_position(self._axis(key))

    def read_xy_position(self) -> tuple[int | None, int | None]:
        x_pos = self.read_axis_position("x") if self.axis_connected("x") else None
        y_pos = self.read_axis_position("y") if self.axis_connected("y") else None
        return x_pos, y_pos

    def check_status_bit(self, key: str, bit: int) -> int:
        client = self._client_for_axis(key)
        return client.check_internal_status(self._axis(key), bit)

    def reference_switch_states(self) -> dict[str, bool | None]:
        states: dict[str, bool | None] = {
            "z_top": None,
            "x_load": None,
            "x_center": None,
            "y_load": None,
            "y_center": None,
        }
        if self.axis_connected("updown"):
            states["z_top"] = self.check_status_bit("updown", 4) == 1
        if self.axis_connected("x"):
            states["x_load"] = self.check_status_bit("x", 4) == 0
            states["x_center"] = self.check_status_bit("x", 5) == 0
        if self.axis_connected("y"):
            states["y_load"] = self.check_status_bit("y", 5) == 0
            states["y_center"] = self.check_status_bit("y", 6) == 0
        return states

    def home_to_top(self) -> MoveResult:
        client = self._client_for_axis("updown")
        return client.home_to_top(self.updown_axis)

    def _require_axes(self, *keys: str) -> None:
        missing = [key.upper() for key in keys if not self.axis_connected(key)]
        if missing:
            raise HardwareError(f"Missing motor connections: {', '.join(missing)}")

    def _prepare_xy_motion(self, action: str, *, allow_without_updown: bool = False) -> None:
        self._require_axes("x", "y")
        if not self.axis_connected("updown"):
            if allow_without_updown:
                return
            raise HardwareError(f"Cannot {action}: Z / up-down axis is not connected.")
        if self.check_status_bit("updown", 4) == 0:
            self.home_to_top()
        if self.check_status_bit("updown", 4) == 0:
            raise HardwareError(f"Cannot {action}: up/down axis is not homed to top.")

    def home_xy_to_center(self, *, allow_without_updown: bool = False) -> tuple[MoveResult, MoveResult]:
        self._prepare_xy_motion("home XY to center", allow_without_updown=allow_without_updown)

        start = time.monotonic()
        stop_x = False
        stop_y = False

        while (self.check_status_bit("x", 4) != 0) or (self.check_status_bit("y", 5) != 0):
            if time.monotonic() - start > 60.0:
                raise HardwareError("Timed out during negative-limit XY homing pass.")
            if self.check_status_bit("x", 4) != 0 and not stop_x:
                self._client_for_axis("x").move_motor(
                    self.x_axis,
                    target=self.motion_defaults.xy_neg_homing_distance,
                    velocity=self.motion_defaults.changer_speed,
                    wait_for_stop=False,
                    stop_enable=-1,
                    stop_condition=0,
                    relative_mode=True,
                )
                stop_x = True
            if self.check_status_bit("y", 5) != 0 and not stop_y:
                self._client_for_axis("y").move_motor(
                    self.y_axis,
                    target=self.motion_defaults.xy_neg_homing_distance,
                    velocity=self.motion_defaults.changer_speed,
                    wait_for_stop=False,
                    stop_enable=-2,
                    stop_condition=0,
                    relative_mode=True,
                )
                stop_y = True
            time.sleep(0.05)

        if (self.check_status_bit("x", 4) != 0) or (self.check_status_bit("y", 5) != 0):
            raise HardwareError("Homed XY center pass 1 failed: did not hit negative limit switches.")

        start = time.monotonic()
        stop_x = False
        stop_y = False
        while (self.check_status_bit("x", 5) != 0) or (self.check_status_bit("y", 6) != 0):
            if time.monotonic() - start > 60.0:
                raise HardwareError("Timed out during positive-limit XY homing pass.")
            if self.check_status_bit("x", 5) != 0 and not stop_x:
                self._client_for_axis("x").move_motor(
                    self.x_axis,
                    target=self.motion_defaults.xy_pos_homing_distance,
                    velocity=self.motion_defaults.changer_speed,
                    wait_for_stop=False,
                    stop_enable=-2,
                    stop_condition=0,
                    relative_mode=True,
                )
                stop_x = True
            if self.check_status_bit("y", 6) != 0 and not stop_y:
                self._client_for_axis("y").move_motor(
                    self.y_axis,
                    target=self.motion_defaults.xy_pos_homing_distance,
                    velocity=self.motion_defaults.changer_speed,
                    wait_for_stop=False,
                    stop_enable=-3,
                    stop_condition=0,
                    relative_mode=True,
                )
                stop_y = True
            time.sleep(0.05)

        if (self.check_status_bit("x", 5) != 0) or (self.check_status_bit("y", 6) != 0):
            raise HardwareError("Homed XY center pass 2 failed: did not hit positive limit switches.")

        x_client = self._client_for_axis("x")
        y_client = self._client_for_axis("y")
        x_client.zero_target_pos(self.x_axis)
        y_client.zero_target_pos(self.y_axis)
        x_pos = x_client.read_position(self.x_axis)
        y_pos = y_client.read_position(self.y_axis)
        return (
            MoveResult(target=0, final_position=x_pos, success=True),
            MoveResult(target=0, final_position=y_pos, success=True),
        )

    def move_xy_to_corner(self, *, allow_without_updown: bool = False) -> tuple[MoveResult, MoveResult]:
        self._prepare_xy_motion("move XY to corner", allow_without_updown=allow_without_updown)

        start = time.monotonic()
        stop_x = False
        stop_y = False
        while (self.check_status_bit("x", 4) != 0) or (self.check_status_bit("y", 5) != 0):
            if time.monotonic() - start > 60.0:
                raise HardwareError("Timed out moving XY stage to corner.")
            if self.check_status_bit("x", 4) != 0 and not stop_x:
                self._client_for_axis("x").move_motor(
                    self.x_axis,
                    target=self.motion_defaults.xy_corner_distance,
                    velocity=self.motion_defaults.changer_speed,
                    wait_for_stop=False,
                    stop_enable=-1,
                    stop_condition=0,
                    relative_mode=True,
                )
                stop_x = True
            if self.check_status_bit("y", 5) != 0 and not stop_y:
                self._client_for_axis("y").move_motor(
                    self.y_axis,
                    target=self.motion_defaults.xy_corner_distance,
                    velocity=self.motion_defaults.changer_speed,
                    wait_for_stop=False,
                    stop_enable=-2,
                    stop_condition=0,
                    relative_mode=True,
                )
                stop_y = True
            time.sleep(0.05)

        x_pos = self.read_axis_position("x")
        y_pos = self.read_axis_position("y")
        return (
            MoveResult(target=x_pos, final_position=x_pos, success=True),
            MoveResult(target=y_pos, final_position=y_pos, success=True),
        )

    def move_xy_absolute(self, x_target: int, y_target: int, *, allow_without_updown: bool = False) -> tuple[MoveResult, MoveResult]:
        self._prepare_xy_motion("move XY stage", allow_without_updown=allow_without_updown)
        x_result = self._client_for_axis("x").move_motor(
            self.x_axis,
            target=int(x_target),
            velocity=self.motion_defaults.changer_speed,
            wait_for_stop=False,
            acceleration=483184,
            relative_mode=False,
        )
        y_result = self._client_for_axis("y").move_motor(
            self.y_axis,
            target=int(y_target),
            velocity=self.motion_defaults.changer_speed,
            wait_for_stop=True,
            acceleration=483184,
            relative_mode=False,
        )
        final_x = self.read_axis_position("x")
        final_y = self.read_axis_position("y")
        return (
            MoveResult(target=x_result.target, final_position=final_x, success=True),
            MoveResult(target=y_result.target, final_position=final_y, success=True),
        )

    def move_axis_relative(self, key: str, delta: int, velocity: int | None = None) -> MoveResult:
        self._require_axes(key)
        current = self.read_axis_position(key)
        axis = self._axis(key)
        if velocity is None:
            move_velocity = self.motion_defaults.lift_speed_slow if key == "updown" else self.motion_defaults.changer_speed
        else:
            move_velocity = int(velocity)
        return self._client_for_axis(key).move_motor(
            axis,
            target=int(current + delta),
            velocity=move_velocity,
            wait_for_stop=True,
            acceleration=483184,
            relative_mode=False,
        )


class StageScene(QtWidgets.QWidget):
    holeSelected = QtCore.Signal(int)
    holeActivated = QtCore.Signal(int)
    specialTargetSelected = QtCore.Signal(str)
    specialTargetActivated = QtCore.Signal(str)
    CENTER_DROP_HOLE = TRAY_CENTER_DROP_HOLE
    LOAD_TARGET = "load"
    CENTER_TARGET = "center"

    def __init__(self, slot_min: int, slot_max: int, parent: QtWidgets.QWidget | None = None) -> None:
        super().__init__(parent)
        self._slot_min = slot_min
        self._slot_max = slot_max
        self._target_hole = slot_min
        self._current_hole: int | None = None
        self._current_xy: tuple[int | None, int | None] = (None, None)
        self._calibrations: dict[int, tuple[int, int]] = {}
        self._hole_points: dict[int, QtCore.QPointF] = {}
        self._loading_position_active = False
        self._special_target: str | None = None
        self._logical_positions = tray_logical_hole_positions(slot_min, slot_max)
        self._min_x, self._max_x, self._min_y, self._max_y = tray_logical_bounds(slot_min, slot_max)
        self.setMinimumSize(520, 520)
        self.setMouseTracking(True)
        self.setSizePolicy(QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Expanding)

    def hasHeightForWidth(self) -> bool:
        return True

    def heightForWidth(self, width: int) -> int:
        return width

    def sizeHint(self) -> QtCore.QSize:
        return QtCore.QSize(760, 760)

    def set_target_hole(self, hole: int) -> None:
        self._target_hole = hole
        self._special_target = None
        self.update()

    def set_current_hole(self, hole: int | None) -> None:
        self._current_hole = hole
        self.update()

    def set_calibrations(self, calibrations: dict[int, tuple[int, int]]) -> None:
        self._calibrations = dict(calibrations)
        self.update()

    def set_current_xy(self, x_pos: int | None, y_pos: int | None) -> None:
        self._current_xy = (x_pos, y_pos)
        self.update()

    def set_loading_position_active(self, active: bool) -> None:
        self._loading_position_active = active
        self.update()

    def set_special_target(self, target: str | None) -> None:
        self._special_target = target
        self.update()

    def _stage_rect(self) -> QtCore.QRectF:
        margin = 8.0
        side = max(120.0, min(self.width(), self.height()) - 2 * margin)
        return QtCore.QRectF((self.width() - side) * 0.5, (self.height() - side) * 0.5, side, side)

    def _layout_rect(self) -> tuple[QtCore.QRectF, float]:
        stage = self._stage_rect()
        usable = QtCore.QRectF(
            stage.left() + stage.width() * 0.06,
            stage.top() + stage.height() * 0.10,
            stage.width() * 0.88,
            stage.height() * 0.80,
        )
        width_units = self._max_x - self._min_x
        height_units = self._max_y - self._min_y
        scale = min(usable.width() / width_units, usable.height() / height_units)
        actual = QtCore.QRectF(0, 0, width_units * scale, height_units * scale)
        actual.moveCenter(usable.center())
        return actual, scale

    def _deck_rect(self) -> QtCore.QRectF:
        rect, scale = self._layout_rect()
        return rect.adjusted(-0.92 * scale, -0.92 * scale, 0.92 * scale, 0.92 * scale)

    def _cup_radius(self) -> float:
        _, scale = self._layout_rect()
        return scale * (TRAY_CUP_DIAMETER_IN / TRAY_CUP_PITCH_IN) * 0.5

    def _drop_radius(self) -> float:
        _, scale = self._layout_rect()
        return scale * (TRAY_CENTER_DROP_DIAMETER_IN / TRAY_CUP_PITCH_IN) * 0.5

    def _logical_to_scene(self, point: QtCore.QPointF) -> QtCore.QPointF:
        rect, scale = self._layout_rect()
        x = rect.left() + (point.x() - self._min_x) * scale
        y = rect.top() + (point.y() - self._min_y) * scale
        return QtCore.QPointF(x, y)

    def _hole_centers(self) -> dict[int, QtCore.QPointF]:
        if self._hole_points:
            return self._hole_points
        self._hole_points = {hole: self._logical_to_scene(point) for hole, point in self._logical_positions.items()}
        return self._hole_points

    def _drop_center(self) -> QtCore.QPointF:
        return self._hole_centers()[self.CENTER_DROP_HOLE]

    def _loading_cup_center(self) -> QtCore.QPointF:
        deck = self._deck_rect()
        stage = self._stage_rect()
        cup_radius = self._cup_radius()
        target = QtCore.QPointF(deck.left() - cup_radius * 1.35, deck.bottom() + cup_radius * 1.05)
        padding = cup_radius * 1.7
        clamped_x = min(max(target.x(), stage.left() + padding), stage.right() - padding)
        clamped_y = min(max(target.y(), stage.top() + padding), stage.bottom() - padding)
        return QtCore.QPointF(clamped_x, clamped_y)

    def _crosshair_color(self) -> QtGui.QColor:
        return QtGui.QColor("#f7f3d7")

    def _cup_label_rect(self, point: QtCore.QPointF, radius: float) -> QtCore.QRectF:
        return QtCore.QRectF(point.x() - radius, point.y() - radius * 0.72, radius * 2, radius * 1.44)

    def resizeEvent(self, event: QtGui.QResizeEvent) -> None:
        self._hole_points.clear()
        super().resizeEvent(event)

    def _nearest_hole(self, pos: QtCore.QPointF) -> int:
        centers = self._hole_centers()
        best_hole = self._slot_min
        best_dist = float("inf")
        for hole, point in centers.items():
            dist = (point.x() - pos.x()) ** 2 + (point.y() - pos.y()) ** 2
            if dist < best_dist:
                best_dist = dist
                best_hole = hole
        return best_hole

    def _special_target_at(self, pos: QtCore.QPointF) -> str | None:
        cup_radius = self._cup_radius()
        load_center = self._loading_cup_center()
        if (load_center.x() - pos.x()) ** 2 + (load_center.y() - pos.y()) ** 2 <= (cup_radius * 1.2) ** 2:
            return self.LOAD_TARGET
        drop_center = self._drop_center()
        if (drop_center.x() - pos.x()) ** 2 + (drop_center.y() - pos.y()) ** 2 <= (self._drop_radius() * 1.24) ** 2:
            return self.CENTER_TARGET
        return None

    def _current_marker_point(self) -> QtCore.QPointF | None:
        x_pos, y_pos = self._current_xy
        if x_pos is None or y_pos is None:
            return None
        if self._current_hole is not None and self._current_hole in self._hole_centers():
            return self._hole_centers()[self._current_hole]
        if len(self._calibrations) < 2:
            return None
        points = self._hole_centers()
        calibrated_points = [(points[hole], xy) for hole, xy in self._calibrations.items() if hole in points]
        if len(calibrated_points) < 2:
            return None
        scene_xs = [point.x() for point, _ in calibrated_points]
        scene_ys = [point.y() for point, _ in calibrated_points]
        motor_xs = [xy[0] for _, xy in calibrated_points]
        motor_ys = [xy[1] for _, xy in calibrated_points]
        if max(motor_xs) == min(motor_xs) or max(motor_ys) == min(motor_ys):
            return None
        x_ratio = (x_pos - min(motor_xs)) / float(max(motor_xs) - min(motor_xs))
        y_ratio = (y_pos - min(motor_ys)) / float(max(motor_ys) - min(motor_ys))
        x_ratio = max(0.0, min(1.0, x_ratio))
        y_ratio = max(0.0, min(1.0, y_ratio))
        return QtCore.QPointF(
            min(scene_xs) + x_ratio * (max(scene_xs) - min(scene_xs)),
            min(scene_ys) + y_ratio * (max(scene_ys) - min(scene_ys)),
        )

    def mousePressEvent(self, event: QtGui.QMouseEvent) -> None:
        if event.button() != QtCore.Qt.LeftButton:
            return
        special_target = self._special_target_at(event.position())
        if special_target is not None:
            self.specialTargetSelected.emit(special_target)
            event.accept()
            return
        self.holeSelected.emit(self._nearest_hole(event.position()))

    def mouseDoubleClickEvent(self, event: QtGui.QMouseEvent) -> None:
        if event.button() != QtCore.Qt.LeftButton:
            return
        special_target = self._special_target_at(event.position())
        if special_target is not None:
            self.specialTargetSelected.emit(special_target)
            self.specialTargetActivated.emit(special_target)
            event.accept()
            return
        hole = self._nearest_hole(event.position())
        self.holeSelected.emit(hole)
        self.holeActivated.emit(hole)
        event.accept()

    def paintEvent(self, event: QtGui.QPaintEvent) -> None:
        del event
        painter = QtGui.QPainter(self)
        painter.setRenderHint(QtGui.QPainter.Antialiasing)
        centers = self._hole_centers()
        cup_radius = self._cup_radius()
        drop_center = self._drop_center()
        drop_radius = self._drop_radius()
        font = painter.font()
        font.setPointSize(9)
        painter.setFont(font)
        for hole, point in centers.items():
            tick_rect = QtCore.QRectF(point.x() - cup_radius * 0.10, point.y() - cup_radius * 1.58, cup_radius * 0.20, cup_radius * 0.34)
            painter.setPen(QtCore.Qt.NoPen)
            painter.setBrush(QtGui.QColor("#1f1f1f"))
            painter.drawRoundedRect(tick_rect, 1.5, 1.5)

            if hole == self.CENTER_DROP_HOLE:
                continue
            engraved = QtGui.QRadialGradient(point, cup_radius * 1.17)
            engraved.setColorAt(0.0, QtGui.QColor(244, 236, 229, 72))
            engraved.setColorAt(0.75, QtGui.QColor(210, 195, 181, 118))
            engraved.setColorAt(1.0, QtGui.QColor(154, 138, 126, 126))
            painter.setPen(QtGui.QPen(QtGui.QColor(250, 246, 242, 148), 1.08))
            painter.setBrush(engraved)
            painter.drawEllipse(point, cup_radius, cup_radius)

            inner = QtGui.QRadialGradient(point, cup_radius * 0.80)
            inner.setColorAt(0.0, QtGui.QColor(255, 255, 255, 16))
            inner.setColorAt(1.0, QtGui.QColor(112, 96, 86, 30))
            painter.setPen(QtGui.QPen(QtGui.QColor(255, 251, 247, 82), 0.8))
            painter.setBrush(inner)
            painter.drawEllipse(point, cup_radius * 0.82, cup_radius * 0.82)

            if hole in self._calibrations:
                painter.setPen(QtGui.QPen(QtGui.QColor("#2ca58d"), 2.0))
                painter.setBrush(QtCore.Qt.NoBrush)
                painter.drawEllipse(point, cup_radius + 3, cup_radius + 3)
            if self._special_target is None and hole == self._target_hole:
                painter.setPen(QtGui.QPen(QtGui.QColor("#ff9f43"), 2.8))
                painter.drawEllipse(point, cup_radius + 7, cup_radius + 7)
            if hole == self._current_hole:
                painter.setPen(QtGui.QPen(self._crosshair_color(), 2.4))
                painter.drawEllipse(point, cup_radius + 11, cup_radius + 11)

            painter.setPen(QtGui.QColor(82, 74, 66, 180))
            painter.drawText(self._cup_label_rect(point, cup_radius * 0.82), QtCore.Qt.AlignCenter, str(hole))

        load_center = self._loading_cup_center()
        load_outer = QtGui.QColor("#1b7f6b") if self._loading_position_active else QtGui.QColor(120, 137, 130, 110)
        load_fill = QtGui.QColor("#3bd1ae") if self._loading_position_active else QtGui.QColor(222, 230, 226, 180)
        painter.setPen(QtGui.QPen(load_outer, 2.4))
        painter.setBrush(load_fill)
        painter.drawEllipse(load_center, cup_radius * 0.92, cup_radius * 0.92)
        if self._special_target == self.LOAD_TARGET:
            painter.setPen(QtGui.QPen(QtGui.QColor("#ff9f43"), 2.6))
            painter.setBrush(QtCore.Qt.NoBrush)
            painter.drawEllipse(load_center, cup_radius * 1.18, cup_radius * 1.18)
        elif self._loading_position_active:
            painter.setPen(QtGui.QPen(QtGui.QColor("#f7f3d7"), 2.2))
            painter.setBrush(QtCore.Qt.NoBrush)
            painter.drawEllipse(load_center, cup_radius * 1.2, cup_radius * 1.2)
        painter.setPen(QtGui.QColor(74, 66, 58, 210))
        painter.drawText(self._cup_label_rect(load_center, cup_radius * 1.12), QtCore.Qt.AlignCenter, "LOAD")

        drop_shadow = QtGui.QRadialGradient(drop_center, drop_radius * 1.22)
        drop_shadow.setColorAt(0.0, QtGui.QColor(74, 54, 34, 118))
        drop_shadow.setColorAt(0.65, QtGui.QColor(88, 66, 48, 78))
        drop_shadow.setColorAt(1.0, QtGui.QColor(90, 68, 50, 0))
        painter.setPen(QtCore.Qt.NoPen)
        painter.setBrush(drop_shadow)
        painter.drawEllipse(drop_center, drop_radius * 1.22, drop_radius * 1.22)

        painter.setPen(QtGui.QPen(QtGui.QColor("#fff0d8"), 1.9))
        painter.setBrush(QtGui.QColor("#6b5138"))
        painter.drawEllipse(drop_center, drop_radius, drop_radius)
        painter.setPen(QtGui.QPen(QtGui.QColor("#f4dec2"), 1.05))
        painter.setBrush(QtCore.Qt.NoBrush)
        painter.drawEllipse(drop_center, drop_radius + 4, drop_radius + 4)
        painter.setPen(QtGui.QColor(110, 84, 60, 170))
        painter.drawText(self._cup_label_rect(drop_center, cup_radius * 0.82), QtCore.Qt.AlignCenter, str(self.CENTER_DROP_HOLE))

        if (
            self._target_hole == self.CENTER_DROP_HOLE
            or self._current_hole == self.CENTER_DROP_HOLE
            or self._special_target == self.CENTER_TARGET
        ):
            ring_color = QtGui.QColor("#ffb25c") if self._special_target == self.CENTER_TARGET or self._target_hole == self.CENTER_DROP_HOLE else self._crosshair_color()
            painter.setPen(QtGui.QPen(ring_color, 3))
            ring_padding = 9 if ring_color == QtGui.QColor("#ffb25c") else 13
            painter.drawEllipse(drop_center, drop_radius + ring_padding, drop_radius + ring_padding)

        marker = self._current_marker_point()
        if marker is not None:
            painter.setPen(QtGui.QPen(self._crosshair_color(), 2))
            painter.setBrush(self._crosshair_color())
            painter.drawEllipse(marker, 5, 5)
            painter.drawLine(marker.x() - 16, marker.y(), marker.x() + 16, marker.y())
            painter.drawLine(marker.x(), marker.y() - 16, marker.x(), marker.y() + 16)
        painter.end()


class MainWindow(QtWidgets.QMainWindow):
    SLOT_MIN = 1
    SLOT_MAX = 100
    POSITION_TOLERANCE = 1000
    DEFAULT_MANUAL_SPEED_CM_PER_SEC = 1.0
    HIGH_SPEED_CONFIRM_CM_PER_SEC = 10.0
    XY_UI_SPEED_REFERENCE_CM_PER_SEC = 10.0

    def __init__(self) -> None:
        super().__init__()
        self.setWindowTitle("RapidPy Changer XY Control")
        self.config = self._load_config()
        self.settings_profile = _load_settings_profile(Path(self.config.settings_path))
        self.motion_defaults = self.settings_profile.motion_defaults
        self.SLOT_MIN = self.settings_profile.slot_min
        self.SLOT_MAX = self.settings_profile.slot_max
        self.controller = ChangerStageController(self.settings_profile)
        self.config.calibrations = [HoleCalibration(hole=item.hole, x=item.x, y=item.y) for item in self.settings_profile.calibrations]
        self._port_boxes: dict[str, QtWidgets.QComboBox] = {}
        self._port_status: dict[str, QtWidgets.QLabel] = {}
        self._port_toggle_buttons: dict[str, tuple[QtWidgets.QPushButton, QtWidgets.QPushButton]] = {}
        self._port_row_widgets: dict[str, tuple[QtWidgets.QWidget, ...]] = {}
        self._last_poll_error = ""
        self._switch_labels: dict[str, QtWidgets.QLabel] = {}
        self._current_stage_hole: int | None = None
        self._current_z_pos: int | None = None
        self._position_source = "Unknown"
        self._suppress_speed_confirmations = False
        self._last_confirmed_speed_values: dict[str, int] = {
            "xy": 0,
            "updown": 0,
        }
        self._last_live_positions: dict[str, int | None] = {"x": None, "y": None, "updown": None}
        self._switch_anchor_positions: dict[str, int | None] = {
            "x_load": None,
            "x_center": self.settings_profile.xy_home_x,
            "y_load": None,
            "y_center": self.settings_profile.xy_home_y,
        }
        self._build_ui()
        self._apply_compact_font()
        self._apply_module_style_overrides()
        self.setMinimumSize(1420, 860)
        self.resize(1620, 940)
        self._load_config_into_widgets()
        self._poll_timer = QtCore.QTimer(self)
        self._poll_timer.setInterval(900)
        self._poll_timer.timeout.connect(self._poll_stage_state)
        self._poll_timer.start()
        QtCore.QTimer.singleShot(0, self._poll_stage_state)

    def _load_config(self) -> AppConfig:
        if not CONFIG_PATH.exists():
            return AppConfig()
        try:
            payload = json.loads(CONFIG_PATH.read_text(encoding="utf-8"))
        except Exception:
            return AppConfig()
        calibrations = [HoleCalibration(**item) for item in payload.get("calibrations", []) if isinstance(item, dict)]
        return AppConfig(
            x_port=str(payload.get("x_port", "COM4")),
            y_port=str(payload.get("y_port", "COM5")),
            updown_port=str(payload.get("updown_port", "COM6")),
            sensor_port=str(payload.get("sensor_port", "")),
            updown_disabled=bool(payload.get("updown_disabled", False)),
            settings_path=str(payload.get("settings_path", str(DEFAULT_SETTINGS_PATH))),
            selected_hole=int(payload.get("selected_hole", 1)),
            calibrations=calibrations,
        )

    def _save_config(self) -> None:
        self.config.x_port = self._port_boxes["x"].currentText().strip()
        self.config.y_port = self._port_boxes["y"].currentText().strip()
        self.config.updown_port = self._port_boxes["updown"].currentText().strip()
        self.config.sensor_port = self._port_boxes["sensor"].currentText().strip()
        self.config.updown_disabled = self._xy_only_mode_enabled()
        self.config.settings_path = self.settings_path_edit.text().strip() or str(DEFAULT_SETTINGS_PATH)
        self.config.selected_hole = self.target_hole.value()
        self.config.calibrations = [
            HoleCalibration(hole=hole, x=x_pos, y=y_pos)
            for hole, (x_pos, y_pos) in sorted(self._calibration_map().items())
        ]
        CONFIG_PATH.write_text(json.dumps(asdict(self.config), indent=2), encoding="utf-8")

    def _build_card(self, title_text: str, subtitle_text: str | None = None) -> tuple[QtWidgets.QFrame, QtWidgets.QVBoxLayout]:
        card = QtWidgets.QFrame()
        card.setObjectName("card")
        layout = QtWidgets.QVBoxLayout(card)
        layout.setContentsMargins(12, 12, 12, 12)
        layout.setSpacing(8)
        title = QtWidgets.QLabel(title_text)
        title.setObjectName("title")
        layout.addWidget(title)
        if subtitle_text:
            subtitle = QtWidgets.QLabel(subtitle_text)
            subtitle.setObjectName("subtitle")
            subtitle.setWordWrap(True)
            layout.addWidget(subtitle)
        apply_card_shadow(card)
        return card, layout

    def _apply_compact_font(self) -> None:
        compact_font = QtGui.QFont(self.font())
        compact_size = compact_font.pointSizeF()
        if compact_size > 0:
            compact_font.setPointSizeF(max(9.0, compact_size - 0.4))
            self.setFont(compact_font)

    def _apply_module_style_overrides(self) -> None:
        self.setStyleSheet(
            """
            QScrollArea#panelScroll {
                background: transparent;
                border: none;
            }
            QScrollArea#panelScroll > QWidget > QWidget {
                background: transparent;
            }
            QLabel#mapStatus,
            QLabel#renderStatus {
                background: rgba(255, 255, 255, 0.86);
                border: 1px solid rgba(122, 2, 25, 0.18);
                border-radius: 14px;
                padding: 8px 10px;
                color: #4d3a39;
            }
            QLabel#tableHint {
                color: #6a5b54;
                background: rgba(255, 255, 255, 0.78);
                border: 1px solid rgba(122, 2, 25, 0.10);
                border-radius: 12px;
                padding: 7px 9px;
            }
            QWidget#portToggleGroup {
                background: rgba(255, 255, 255, 0.92);
                border: 1px solid rgba(122, 2, 25, 0.28);
                border-radius: 12px;
            }
            QPushButton#portToggle {
                min-width: 32px;
                padding: 4px 6px;
                border: 1px solid transparent;
                border-radius: 9px;
                background: transparent;
                color: #5d4d49;
            }
            QPushButton#portToggle:checked {
                background: qlineargradient(x1:0, y1:0, x2:1, y2:1, stop:0 #7A0219, stop:1 #5a0013);
                color: #ffffff;
                border: 1px solid rgba(122, 2, 25, 0.82);
            }
            QWidget#renderViewport {
                background: #e8e0d3;
                border-radius: 18px;
            }
            QTableWidget#cupTable {
                font-size: 11px;
            }
            QLabel#profileBadge {
                color: #4f403b;
                background: rgba(255, 255, 255, 0.88);
                border: 1px solid rgba(122, 2, 25, 0.14);
                border-radius: 12px;
                padding: 6px 9px;
            }
            QPlainTextEdit#console {
                background: rgba(255, 255, 255, 0.94);
                color: #2f2827;
                border-radius: 16px;
                border: 1px solid rgba(122, 2, 25, 0.16);
                padding: 8px;
                selection-background-color: rgba(122, 2, 25, 0.18);
                selection-color: #2f2827;
            }
            QCheckBox#updownModeCheckBox {
                color: #4d302c;
                font-size: 12.5px;
                font-weight: 700;
                background: rgba(255, 249, 244, 0.96);
                border: 1px solid rgba(122, 2, 25, 0.26);
                border-radius: 14px;
                padding: 10px 12px;
                spacing: 10px;
            }
            QCheckBox#updownModeCheckBox:hover {
                background: rgba(255, 245, 238, 0.98);
                border: 1px solid rgba(122, 2, 25, 0.38);
            }
            QCheckBox#updownModeCheckBox::indicator {
                width: 22px;
                height: 22px;
                border-radius: 7px;
                border: 2px solid rgba(122, 2, 25, 0.58);
                background: rgba(255, 255, 255, 0.98);
            }
            QCheckBox#updownModeCheckBox::indicator:unchecked {
                background: rgba(255, 255, 255, 0.98);
                border: 2px solid rgba(122, 2, 25, 0.58);
            }
            QCheckBox#updownModeCheckBox::indicator:checked {
                background: qlineargradient(x1:0, y1:0, x2:1, y2:1, stop:0 #7A0219, stop:1 #5a0013);
                border: 2px solid rgba(122, 2, 25, 0.88);
                image: url(none);
            }
            """
        )

    def _stage_limit_instructions(self) -> str:
        return (
            "VB6 limits: XY and Z jogs are entered here as raw counts/s, with live cm/s estimates shown underneath for readability. "
            f"VB6 XY max is {self.motion_defaults.changer_speed} counts/s; XY homing sweeps "
            f"{self.motion_defaults.xy_neg_homing_distance} to {self.motion_defaults.xy_pos_homing_distance} and requires Z top first. "
            "Z top uses bit 4; X load/center use bits 4/5; Y load/center use bits 5/6. "
            f"{self.motion_defaults.lift_speed_fast} counts/s; VB6 references are slow={self.motion_defaults.lift_speed_slow}, "
            f"normal={self.motion_defaults.lift_speed_normal}, fast={self.motion_defaults.lift_speed_fast}."
        )

    def _xy_counts_per_cm(self) -> tuple[float | None, float | None, float | None]:
        x_pitch, y_pitch = self._xy_pitch_counts()
        # VB6 stores XYTable positions column-major. Adjacent X columns are spaced by the
        # staggered-column pitch, while adjacent Y positions inside a column use the full cup pitch.
        x_counts_per_cm = None if x_pitch is None else x_pitch / (TRAY_CUP_PITCH_IN * TRAY_ROW_STEP * 2.54)
        y_counts_per_cm = None if y_pitch is None else y_pitch / (TRAY_CUP_PITCH_IN * 2.54)
        available = [value for value in (x_counts_per_cm, y_counts_per_cm) if value is not None and value > 0]
        average = sum(available) / len(available) if available else None
        return x_counts_per_cm, y_counts_per_cm, average

    def _velocity_raw_scale(self) -> float:
        full_rotation = abs(float(self.motion_defaults.turning_motor_full_rotation))
        one_rps = abs(float(self.motion_defaults.turning_motor_1rps))
        if full_rotation > 0 and one_rps > 0:
            return one_rps / full_rotation
        return 2000.0

    def _raw_velocity_to_position_counts_per_second(self, raw_velocity: int | float) -> float:
        return float(raw_velocity) / self._velocity_raw_scale()

    def _position_counts_per_second_to_raw_velocity(self, position_counts_per_second: float) -> int:
        return max(1, int(round(position_counts_per_second * self._velocity_raw_scale())))

    def _xy_command_speed_for_cm_per_second(self, cm_per_second: float) -> int:
        _x_counts_per_cm, _y_counts_per_cm, average_counts_per_cm = self._xy_counts_per_cm()
        if average_counts_per_cm is None or average_counts_per_cm <= 0:
            scaled = self.motion_defaults.changer_speed * (cm_per_second / self.XY_UI_SPEED_REFERENCE_CM_PER_SEC)
            return max(100_000, min(int(round(scaled)), int(self.motion_defaults.changer_speed)))
        raw_velocity = self._position_counts_per_second_to_raw_velocity(cm_per_second * average_counts_per_cm)
        return max(100_000, min(raw_velocity, int(self.motion_defaults.changer_speed)))

    def _z_counts_for_cm_per_second(self, cm_per_second: float) -> int:
        counts_per_cm = self.settings_profile.updown_motor_1cm
        if counts_per_cm <= 0:
            return max(1, int(round(self.motion_defaults.lift_speed_slow / 5.0)))
        raw_velocity = self._position_counts_per_second_to_raw_velocity(counts_per_cm * cm_per_second)
        return max(1, min(raw_velocity, int(self.motion_defaults.lift_speed_fast)))

    def _max_xy_speed_cm_per_second(self) -> float:
        _x_counts_per_cm, _y_counts_per_cm, average_counts_per_cm = self._xy_counts_per_cm()
        if average_counts_per_cm is None or average_counts_per_cm <= 0:
            return self.XY_UI_SPEED_REFERENCE_CM_PER_SEC
        return max(
            self.HIGH_SPEED_CONFIRM_CM_PER_SEC,
            min(50.0, self._raw_velocity_to_position_counts_per_second(self.motion_defaults.changer_speed) / average_counts_per_cm),
        )

    def _max_z_speed_cm_per_second(self) -> float:
        counts_per_cm = self.settings_profile.updown_motor_1cm
        if counts_per_cm <= 0:
            return 25.0
        return max(
            self.HIGH_SPEED_CONFIRM_CM_PER_SEC,
            min(50.0, self._raw_velocity_to_position_counts_per_second(self.motion_defaults.lift_speed_fast) / counts_per_cm),
        )

    def _default_xy_speed_cm_per_second(self) -> float:
        return min(self.DEFAULT_MANUAL_SPEED_CM_PER_SEC, self._max_xy_speed_cm_per_second())

    def _default_z_speed_cm_per_second(self) -> float:
        return min(self.DEFAULT_MANUAL_SPEED_CM_PER_SEC, self._max_z_speed_cm_per_second())

    def _default_xy_speed_raw(self) -> int:
        return self._xy_command_speed_for_cm_per_second(self._default_xy_speed_cm_per_second())

    def _default_z_speed_raw(self) -> int:
        return self._z_counts_for_cm_per_second(self._default_z_speed_cm_per_second())

    def _safe_jog_step(self) -> int:
        mapping = self._calibration_map()
        candidate_deltas: list[int] = []
        for axis_values in (
            sorted({x_pos for x_pos, _ in mapping.values()}),
            sorted({y_pos for _, y_pos in mapping.values()}),
        ):
            for left, right in zip(axis_values, axis_values[1:]):
                delta = abs(right - left)
                if delta >= 500:
                    candidate_deltas.append(delta)
        if not candidate_deltas:
            return 500
        return max(250, min(5_000, min(candidate_deltas) // 4))

    def _settings_history_dir(self, target: Path) -> Path:
        return target.parent / ".rapidpy_history" / target.stem

    def _create_settings_snapshot(self, target: Path) -> Path | None:
        if not target.exists():
            return None
        history_dir = self._settings_history_dir(target)
        history_dir.mkdir(parents=True, exist_ok=True)
        timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
        snapshot = history_dir / f"{timestamp}{target.suffix}"
        counter = 1
        while snapshot.exists():
            snapshot = history_dir / f"{timestamp}_{counter}{target.suffix}"
            counter += 1
        shutil.copy2(target, snapshot)
        return snapshot

    def _format_profile_header(self, profile: SettingsProfile) -> str:
        return f"{profile.path.name} | Cups {len(profile.calibrations)} | Slots {profile.slot_min}-{profile.slot_max}"

    def _format_profile_ids(self, profile: SettingsProfile) -> str:
        return (
            "Motor IDs  X/Y/Z = "
            f"{profile.x_axis.address}/{profile.y_axis.address}/{profile.updown_axis.address}"
        )

    def _format_profile_motion(self, profile: SettingsProfile) -> str:
        motion = profile.motion_defaults
        return (
            f"VB6 limits  XY max {motion.changer_speed:,} | "
            f"Z slow/normal/fast {motion.lift_speed_slow:,}/{motion.lift_speed_normal:,}/{motion.lift_speed_fast:,}"
        )

    def _xy_pitch_counts(self) -> tuple[float | None, float | None]:
        mapping = self._calibration_map()
        if len(mapping) < 4:
            return None, None
        x_deltas: list[int] = []
        y_deltas: list[int] = []

        for hole, (x_pos, y_pos) in mapping.items():
            next_row = hole + 1
            if (hole - self.SLOT_MIN) % TRAY_GRID_ROWS != TRAY_GRID_ROWS - 1 and next_row in mapping:
                _, next_y = mapping[next_row]
                delta_y = abs(next_y - y_pos)
                if delta_y >= 100:
                    y_deltas.append(delta_y)

            next_column = hole + TRAY_GRID_ROWS
            if next_column in mapping:
                next_x, _ = mapping[next_column]
                delta_x = abs(next_x - x_pos)
                if delta_x >= 100:
                    x_deltas.append(delta_x)

        x_pitch = sum(x_deltas) / len(x_deltas) if x_deltas else None
        y_pitch = sum(y_deltas) / len(y_deltas) if y_deltas else None
        return x_pitch, y_pitch

    def _format_xy_velocity_estimate(self, counts_per_second: int) -> str:
        x_counts_per_cm, y_counts_per_cm, _average = self._xy_counts_per_cm()
        command_text = f"{counts_per_second:,} counts/s"
        position_counts_per_second = self._raw_velocity_to_position_counts_per_second(counts_per_second)
        if x_counts_per_cm is None and y_counts_per_cm is None:
            return f"Raw XY command: {command_text}. Estimated cm/s becomes available once multiple cup positions are loaded."
        parts: list[str] = []
        if x_counts_per_cm is not None and x_counts_per_cm > 0:
            parts.append(f"X ~{position_counts_per_second / x_counts_per_cm:,.2f} cm/s")
        if y_counts_per_cm is not None and y_counts_per_cm > 0:
            parts.append(f"Y ~{position_counts_per_second / y_counts_per_cm:,.2f} cm/s")
        return f"Estimated XY actual speed from raw command {command_text}: {', '.join(parts)}"

    def _format_z_velocity_estimate(self, counts_per_second: int) -> str:
        counts_per_cm = self.settings_profile.updown_motor_1cm
        if counts_per_cm <= 0:
            return "Estimated Z actual speed unavailable because UpDownMotor1cm is not set in the loaded INI."
        cm_per_second = self._raw_velocity_to_position_counts_per_second(counts_per_second) / counts_per_cm
        return f"Estimated Z actual speed: ~{cm_per_second:,.2f} cm/s ({counts_per_second:,} counts/s command)"

    def _update_velocity_estimates(self) -> None:
        if hasattr(self, "xy_velocity_estimate"):
            self.xy_velocity_estimate.setText(
                self._format_xy_velocity_estimate(int(self.model_xy_speed.value()))
            )
        if hasattr(self, "z_velocity_estimate"):
            self.z_velocity_estimate.setText(
                self._format_z_velocity_estimate(int(self.model_z_speed.value()))
            )

    def _estimated_speed_cm_per_second(self, key: str, counts_per_second: int) -> float | None:
        position_counts_per_second = self._raw_velocity_to_position_counts_per_second(counts_per_second)
        if key == "xy":
            x_counts_per_cm, y_counts_per_cm, _average = self._xy_counts_per_cm()
            estimates = [
                position_counts_per_second / axis_counts
                for axis_counts in (x_counts_per_cm, y_counts_per_cm)
                if axis_counts is not None and axis_counts > 0
            ]
            if estimates:
                return max(estimates)
            if self.motion_defaults.changer_speed <= 0:
                return None
            return counts_per_second / self.motion_defaults.changer_speed * self._max_xy_speed_cm_per_second()
        counts_per_cm = self.settings_profile.updown_motor_1cm
        if counts_per_cm <= 0:
            return None
        return position_counts_per_second / counts_per_cm

    def _confirm_speed_if_needed(self, key: str, value: int) -> None:
        if self._suppress_speed_confirmations:
            self._last_confirmed_speed_values[key] = value
            return
        estimated_cm_per_second = self._estimated_speed_cm_per_second(key, value)
        if estimated_cm_per_second is None or estimated_cm_per_second < self.HIGH_SPEED_CONFIRM_CM_PER_SEC:
            self._last_confirmed_speed_values[key] = value
            self._update_velocity_estimates()
            return
        previous = self._last_confirmed_speed_values.get(
            key,
            self._default_xy_speed_raw() if key == "xy" else self._default_z_speed_raw(),
        )
        previous_cm_per_second = self._estimated_speed_cm_per_second(key, previous)
        if previous_cm_per_second is not None and previous_cm_per_second >= self.HIGH_SPEED_CONFIRM_CM_PER_SEC:
            self._last_confirmed_speed_values[key] = value
            self._update_velocity_estimates()
            return
        response = QtWidgets.QMessageBox.question(
            self,
            "High Speed Confirmation",
            f"{value:,} counts/s is about {estimated_cm_per_second:.2f} cm/s, which is unusually fast for manual changer motion. Continue?",
            QtWidgets.QMessageBox.StandardButton.Yes | QtWidgets.QMessageBox.StandardButton.No,
            QtWidgets.QMessageBox.StandardButton.No,
        )
        if response == QtWidgets.QMessageBox.StandardButton.Yes:
            self._last_confirmed_speed_values[key] = value
            self._update_velocity_estimates()
            return
        spinbox = self.model_xy_speed if key == "xy" else self.model_z_speed
        self._suppress_speed_confirmations = True
        spinbox.setValue(previous)
        self._suppress_speed_confirmations = False
        self._update_velocity_estimates()

    def _record_switch_anchor_positions(
        self,
        x_pos: int | None,
        y_pos: int | None,
        switches: dict[str, bool | None],
    ) -> None:
        if x_pos is not None:
            if switches.get("x_load"):
                self._switch_anchor_positions["x_load"] = x_pos
            if switches.get("x_center"):
                self._switch_anchor_positions["x_center"] = x_pos
        if y_pos is not None:
            if switches.get("y_load"):
                self._switch_anchor_positions["y_load"] = y_pos
            if switches.get("y_center"):
                self._switch_anchor_positions["y_center"] = y_pos

    def _infer_xy_from_switches(self, switches: dict[str, bool | None]) -> tuple[int | None, int | None, str]:
        inferred_x: int | None = None
        inferred_y: int | None = None
        sources: list[str] = []
        if switches.get("x_load") and self._switch_anchor_positions["x_load"] is not None:
            inferred_x = self._switch_anchor_positions["x_load"]
            sources.append("X load")
        elif switches.get("x_center") and self._switch_anchor_positions["x_center"] is not None:
            inferred_x = self._switch_anchor_positions["x_center"]
            sources.append("X center")
        if switches.get("y_load") and self._switch_anchor_positions["y_load"] is not None:
            inferred_y = self._switch_anchor_positions["y_load"]
            sources.append("Y load")
        elif switches.get("y_center") and self._switch_anchor_positions["y_center"] is not None:
            inferred_y = self._switch_anchor_positions["y_center"]
            sources.append("Y center")
        if inferred_x is None:
            inferred_x = self._last_live_positions["x"]
        if inferred_y is None:
            inferred_y = self._last_live_positions["y"]
        source_text = "switch anchors" if sources else "last known"
        return inferred_x, inferred_y, source_text

    def _loading_position_active(self, x_pos: int | None, y_pos: int | None, switches: dict[str, bool | None]) -> bool:
        if switches.get("x_load") and switches.get("y_load"):
            return True
        x_anchor = self._switch_anchor_positions["x_load"]
        y_anchor = self._switch_anchor_positions["y_load"]
        if x_pos is None or y_pos is None or x_anchor is None or y_anchor is None:
            return False
        return abs(x_pos - x_anchor) < self.POSITION_TOLERANCE and abs(y_pos - y_anchor) < self.POSITION_TOLERANCE

    def _read_current_position(self) -> None:
        self._poll_stage_state()
        x_text = self.x_pos_label.text()
        y_text = self.y_pos_label.text()
        z_text = self.z_pos_label.text()
        self.stage_status.setText(
            f"Current position ({self._position_source}): X {x_text}, Y {y_text}, Z {z_text}."
        )
        self._append(f"Read current position via {self._position_source}: x={x_text}, y={y_text}, z={z_text}.")

    def _build_current_settings_config(self, base_path: Path | None = None) -> configparser.ConfigParser:
        source_path = base_path or Path(self.settings_path_edit.text().strip() or str(self.settings_profile.path))
        if source_path.exists():
            config = _load_settings_config(source_path)
        else:
            config = _new_settings_config()

        for section_name in ("SampleChanger", "SteppingMotor", "MotorPrograms", "XYTable"):
            if not config.has_section(section_name):
                config.add_section(section_name)

        sample_section = config["SampleChanger"]
        motor_section = config["SteppingMotor"]
        program_section = config["MotorPrograms"]
        table_section = config["XYTable"]

        sample_section["SlotMin"] = str(self.SLOT_MIN)
        sample_section["SlotMax"] = str(self.SLOT_MAX)
        sample_section["OneStep"] = str(self.motion_defaults.one_step)

        motor_section["SampleHoleAlignmentOffset"] = str(self.motion_defaults.sample_hole_alignment_offset)
        motor_section["ChangerSpeed"] = str(self.motion_defaults.changer_speed)
        motor_section["TurnerSpeed"] = str(self.motion_defaults.turner_speed)
        motor_section["TurningMotorFullRotation"] = str(self.motion_defaults.turning_motor_full_rotation)
        motor_section["TurningMotor1rps"] = str(self.motion_defaults.turning_motor_1rps)
        motor_section["LiftSpeedSlow"] = str(self.motion_defaults.lift_speed_slow)
        motor_section["LiftSpeedNormal"] = str(self.motion_defaults.lift_speed_normal)
        motor_section["LiftSpeedFast"] = str(self.motion_defaults.lift_speed_fast)
        motor_section["LiftAcceleration"] = str(self.motion_defaults.lift_acceleration)
        motor_section["MeasPos"] = str(self.motion_defaults.meas_pos)
        motor_section["SampleBottom"] = str(self.motion_defaults.sample_bottom)
        motor_section["SampleTop"] = str(self.motion_defaults.sample_bottom + self.motion_defaults.sample_height)
        motor_section["UpDownTorqueFactor"] = str(self.motion_defaults.updown_torque_factor)
        motor_section["PickupTorqueThrottle"] = str(self.motion_defaults.pickup_torque_throttle)
        motor_section["UpDownMotor1cm"] = str(self.settings_profile.updown_motor_1cm)

        program_section["MotorIDChanger"] = str(self.settings_profile.x_axis.address)
        program_section["MotorIDChangerY"] = str(self.settings_profile.y_axis.address)
        program_section["MotorIDUpDown"] = str(self.settings_profile.updown_axis.address)

        table_section["XYHomeX"] = str(self.settings_profile.xy_home_x)
        table_section["XYHomeY"] = str(self.settings_profile.xy_home_y)
        mapping = self._calibration_map()
        for hole in range(self.SLOT_MIN, self.SLOT_MAX + 1):
            x_key = f"XY{hole}X"
            y_key = f"XY{hole}Y"
            if hole in mapping:
                x_pos, y_pos = mapping[hole]
                table_section[x_key] = str(x_pos)
                table_section[y_key] = str(y_pos)
            else:
                table_section[x_key] = ""
                table_section[y_key] = ""
        return config

    def _write_settings_ini(self, target: Path) -> Path | None:
        config = self._build_current_settings_config(target if target.exists() else self.settings_profile.path)

        snapshot = self._create_settings_snapshot(target)
        with target.open("w", encoding="utf-8", newline="\n") as handle:
            config.write(handle)
        return snapshot

    def _write_settings_json(self, target: Path) -> None:
        config = self._build_current_settings_config(self.settings_profile.path)
        target.write_text(json.dumps(_settings_json_payload_from_config(config), indent=2), encoding="utf-8")

    def _apply_settings_profile(self, profile: SettingsProfile) -> None:
        if profile.slot_min != self.SLOT_MIN or profile.slot_max != self.SLOT_MAX:
            raise ValueError("Loading settings files with a different slot range is not supported by this UI yet.")
        self.settings_profile = profile
        self.motion_defaults = profile.motion_defaults
        self.controller.apply_settings_profile(profile)
        self.config.calibrations = [HoleCalibration(hole=item.hole, x=item.x, y=item.y) for item in profile.calibrations]
        self._switch_anchor_positions["x_center"] = profile.xy_home_x
        self._switch_anchor_positions["y_center"] = profile.xy_home_y
        if hasattr(self, "settings_path_edit"):
            self.settings_path_edit.setText(str(profile.path))
        if hasattr(self, "settings_status"):
            self.settings_status.setText(self._format_profile_header(profile))
        if hasattr(self, "settings_ids_status"):
            self.settings_ids_status.setText(self._format_profile_ids(profile))
        if hasattr(self, "settings_motion_status"):
            self.settings_motion_status.setText(self._format_profile_motion(profile))
        if hasattr(self, "model_xy_speed"):
            self._suppress_speed_confirmations = True
            self.model_xy_speed.setRange(100_000, self.motion_defaults.changer_speed)
            self.model_xy_speed.setValue(self._default_xy_speed_raw())
            self.model_xy_speed.setToolTip(
                "XY manual velocity in raw counts/s. The live text below converts the current raw value into estimated cm/s from the loaded cup spacing."
            )
            self._last_confirmed_speed_values["xy"] = int(self.model_xy_speed.value())
        if hasattr(self, "model_z_speed"):
            self.model_z_speed.setRange(1, self.motion_defaults.lift_speed_fast)
            self.model_z_speed.setValue(self._default_z_speed_raw())
            self.model_z_speed.setToolTip(
                "Z manual velocity in raw counts/s. The live text below converts the current raw value into estimated cm/s using UpDownMotor1cm."
            )
            self._last_confirmed_speed_values["updown"] = int(self.model_z_speed.value())
            self._suppress_speed_confirmations = False
        if hasattr(self, "model_step"):
            self.model_step.setValue(self._safe_jog_step())
        self._update_velocity_estimates()
        if hasattr(self, "target_hole"):
            self.target_hole.setRange(self.SLOT_MIN, self.SLOT_MAX)
        if hasattr(self, "stage_scene"):
            self._update_stage_calibrations()
            self.stage_scene.set_target_hole(self.target_hole.value())

    def _browse_settings_file(self) -> None:
        file_name, _ = QtWidgets.QFileDialog.getOpenFileName(
            self,
            "Load Changer Settings",
            str(Path(self.settings_path_edit.text()).parent if self.settings_path_edit.text() else DEFAULT_SETTINGS_PATH.parent),
            "Settings Files (*.ini *.json);;INI Files (*.ini);;JSON Files (*.json);;All Files (*)",
        )
        if not file_name:
            return
        self.settings_path_edit.setText(file_name)
        self._load_settings_file(Path(file_name))

    def _save_settings_file(self) -> None:
        target = Path(self.settings_path_edit.text().strip() or str(self.settings_profile.path))
        if target.suffix.lower() != ".ini":
            target = target.with_suffix(".ini")
            self.settings_path_edit.setText(str(target))
        try:
            snapshot = self._write_settings_ini(target)
        except Exception as exc:
            QtWidgets.QMessageBox.warning(self, "Save Settings Error", str(exc))
            return
        self.settings_profile = _load_settings_profile(target)
        self._apply_settings_profile(self.settings_profile)
        self._save_config()
        if snapshot is None:
            self._append(f"Saved changer settings to {target}.")
        else:
            self._append(f"Saved changer settings to {target} after snapshotting {snapshot.name}.")

    def _save_settings_file_as(self) -> None:
        selected, _ = QtWidgets.QFileDialog.getSaveFileName(
            self,
            "Save Changer Settings",
            str(self.settings_profile.path.with_suffix(".ini")),
            "INI Files (*.ini);;All Files (*)",
        )
        if not selected:
            return
        target = Path(selected)
        self.settings_path_edit.setText(str(target))
        try:
            snapshot = self._write_settings_ini(target)
        except Exception as exc:
            QtWidgets.QMessageBox.warning(self, "Save Settings Error", str(exc))
            return
        self.settings_profile = _load_settings_profile(target)
        self._apply_settings_profile(self.settings_profile)
        self._save_config()
        if snapshot is None:
            self._append(f"Saved changer settings copy to {target}.")
        else:
            self._append(f"Saved changer settings copy to {target} after snapshotting {snapshot.name}.")

    def _load_settings_file(self, path: Path | None = None) -> None:
        target_path = Path(path or self.settings_path_edit.text().strip() or DEFAULT_SETTINGS_PATH)
        try:
            profile = _load_settings_profile(target_path)
            self._apply_settings_profile(profile)
        except Exception as exc:
            QtWidgets.QMessageBox.warning(self, "Load Settings Error", str(exc))
            return
        self._append(f"Loaded changer settings from {target_path}.")
        self._save_config()
        self._poll_stage_state()

    def _export_settings_json(self) -> None:
        selected, _ = QtWidgets.QFileDialog.getSaveFileName(
            self,
            "Export Changer Settings As JSON",
            str(self.settings_profile.path.with_suffix(".json")),
            "JSON Files (*.json)",
        )
        if not selected:
            return
        target = Path(selected)
        try:
            self._write_settings_json(target)
        except Exception as exc:
            QtWidgets.QMessageBox.warning(self, "Export JSON Error", str(exc))
            return
        self._append(f"Exported changer settings JSON to {target}.")

    def _import_settings_json(self) -> None:
        selected, _ = QtWidgets.QFileDialog.getOpenFileName(
            self,
            "Import Changer Settings JSON",
            str(self.settings_profile.path.with_suffix(".json")),
            "JSON Files (*.json)",
        )
        if not selected:
            return
        self.settings_path_edit.setText(selected)
        self._load_settings_file(Path(selected))

    def _set_port_toggle_state(self, key: str, connected: bool) -> None:
        buttons = self._port_toggle_buttons.get(key)
        if buttons is None:
            return
        on_button, off_button = buttons
        on_button.setChecked(connected)
        off_button.setChecked(not connected)

    def _build_port_row(self, layout: QtWidgets.QGridLayout, row: int, key: str, label_text: str) -> None:
        label = QtWidgets.QLabel(label_text)
        combo = QtWidgets.QComboBox()
        combo.setEditable(True)
        combo.setInsertPolicy(QtWidgets.QComboBox.NoInsert)
        combo.setMinimumWidth(76)
        combo.setSizePolicy(QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Fixed)
        status = QtWidgets.QLabel("Off")
        status.setAlignment(QtCore.Qt.AlignCenter)
        status.setMinimumWidth(50)

        toggle_group = QtWidgets.QWidget()
        toggle_group.setObjectName("portToggleGroup")
        toggle_layout = QtWidgets.QHBoxLayout(toggle_group)
        toggle_layout.setContentsMargins(3, 3, 3, 3)
        toggle_layout.setSpacing(2)

        on_button = QtWidgets.QPushButton("On")
        on_button.setObjectName("portToggle")
        on_button.setCheckable(True)
        on_button.setAutoExclusive(True)
        off_button = QtWidgets.QPushButton("Off")
        off_button.setObjectName("portToggle")
        off_button.setCheckable(True)
        off_button.setAutoExclusive(True)
        off_button.setChecked(True)
        on_button.clicked.connect(lambda: self._connect_port(key))
        off_button.clicked.connect(lambda: self._disconnect_port(key))
        toggle_layout.addWidget(on_button)
        toggle_layout.addWidget(off_button)

        layout.addWidget(label, row, 0)
        layout.addWidget(combo, row, 1)
        layout.addWidget(toggle_group, row, 2)
        layout.addWidget(status, row, 3)
        self._port_boxes[key] = combo
        self._port_status[key] = status
        self._port_toggle_buttons[key] = (on_button, off_button)
        self._port_row_widgets[key] = (label, combo, toggle_group)

    def _xy_only_mode_enabled(self) -> bool:
        return bool(hasattr(self, "disable_updown_check") and self.disable_updown_check.isChecked())

    def _apply_updown_mode(self) -> None:
        disabled = self._xy_only_mode_enabled()
        if disabled and self.controller.axis_connected("updown"):
            self.controller.disconnect_axis("updown")
            self._append("Up/down axis disconnected because XY-only mode is active.")

        for widget in self._port_row_widgets.get("updown", ()):  # Z port row
            widget.setEnabled(not disabled)

        for widget in (
            getattr(self, "home_top_btn", None),
            getattr(self, "model_z_speed", None),
            getattr(self, "model_z_minus_btn", None),
            getattr(self, "model_z_plus_btn", None),
            getattr(self, "model_home_z_btn", None),
        ):
            if widget is not None:
                widget.setEnabled(not disabled)

        if hasattr(self, "z_velocity_estimate"):
            if disabled:
                self.z_velocity_estimate.setText("Up/down disabled. XY-only mode leaves Z jog and Z homing offline.")
            else:
                self._update_velocity_estimates()

        if hasattr(self, "updown_mode_status"):
            if disabled:
                self.updown_mode_status.setText("XY-only mode active: up/down controls are disabled and XY motion will not require Z.")
            else:
                self.updown_mode_status.setText("Up/down enabled: Z homing and Z jog controls are available when connected.")

        if disabled:
            if "updown" in self._port_status:
                self._port_status["updown"].setText("Disabled")
            if "z_top" in self._switch_labels:
                self._switch_labels["z_top"].setText("Disabled")
            if hasattr(self, "z_pos_label"):
                self.z_pos_label.setText("--")

    def _on_disable_updown_toggled(self, checked: bool) -> None:
        if checked:
            response = QtWidgets.QMessageBox.question(
                self,
                "Enable XY-Only Mode?",
                "Disable Up/Down will isolate XY control from the Z axis.\n\n"
                "Confirm that this test is intentionally running without up/down control and that the sample holder rod is raised clear above the XY stage.\n\n"
                "Enable XY-only mode?",
                QtWidgets.QMessageBox.StandardButton.Yes | QtWidgets.QMessageBox.StandardButton.No,
                QtWidgets.QMessageBox.StandardButton.No,
            )
            if response != QtWidgets.QMessageBox.StandardButton.Yes:
                self.disable_updown_check.blockSignals(True)
                self.disable_updown_check.setChecked(False)
                self.disable_updown_check.blockSignals(False)
                self._apply_updown_mode()
                return
            self._append("XY-only mode enabled after operator confirmation.")
        else:
            self._append("Up/down control re-enabled.")

        self._apply_updown_mode()
        self._save_config()
        self._poll_stage_state()

    def _build_ui(self) -> None:
        root = QtWidgets.QWidget(self)
        self.setCentralWidget(root)
        shell = QtWidgets.QHBoxLayout(root)
        shell.setContentsMargins(12, 12, 12, 12)
        shell.setSpacing(10)

        left_scroll = QtWidgets.QScrollArea()
        left_scroll.setObjectName("panelScroll")
        left_scroll.setWidgetResizable(True)
        left_scroll.setFrameShape(QtWidgets.QFrame.Shape.NoFrame)
        left_scroll.setHorizontalScrollBarPolicy(QtCore.Qt.ScrollBarAlwaysOff)
        left_scroll.setVerticalScrollBarPolicy(QtCore.Qt.ScrollBarAsNeeded)
        left_scroll.setMinimumWidth(320)
        left_scroll.setMaximumWidth(370)
        left_host = QtWidgets.QWidget()
        left_scroll.setWidget(left_host)
        left = QtWidgets.QVBoxLayout(left_host)
        left.setContentsMargins(0, 0, 0, 0)
        left.setSpacing(10)

        center_host = QtWidgets.QWidget()
        center_host.setMinimumWidth(650)
        center_layout = QtWidgets.QVBoxLayout(center_host)
        center_layout.setContentsMargins(0, 0, 0, 0)
        center_layout.setSpacing(12)

        right_scroll = QtWidgets.QScrollArea()
        right_scroll.setObjectName("panelScroll")
        right_scroll.setWidgetResizable(True)
        right_scroll.setFrameShape(QtWidgets.QFrame.Shape.NoFrame)
        right_scroll.setHorizontalScrollBarPolicy(QtCore.Qt.ScrollBarAlwaysOff)
        right_scroll.setVerticalScrollBarPolicy(QtCore.Qt.ScrollBarAsNeeded)
        right_scroll.setMinimumWidth(360)
        right_scroll.setMaximumWidth(420)
        right_host = QtWidgets.QWidget()
        right_scroll.setWidget(right_host)
        right = QtWidgets.QVBoxLayout(right_host)
        right.setContentsMargins(0, 0, 0, 0)
        right.setSpacing(10)

        connections_card, connections_layout = self._build_card(
            "Connections",
            "Expose X, Y, Z, and reference-monitor serial links separately. Motion defaults and cup positions come from the VB6 INI; COM selections remain local to this machine.",
        )
        settings_row = QtWidgets.QHBoxLayout()
        self.settings_path_edit = QtWidgets.QLineEdit(str(self.settings_profile.path))
        self.settings_path_edit.setReadOnly(True)
        self.browse_settings_btn = QtWidgets.QPushButton("Browse")
        settings_row.addWidget(self.settings_path_edit, stretch=1)
        settings_row.addWidget(self.browse_settings_btn)
        connections_layout.addLayout(settings_row)
        settings_actions = QtWidgets.QHBoxLayout()
        self.load_settings_btn = QtWidgets.QPushButton("Reload Settings")
        self.save_settings_btn = QtWidgets.QPushButton("Save INI")
        self.save_settings_as_btn = QtWidgets.QPushButton("Save INI As")
        settings_actions.addWidget(self.load_settings_btn)
        settings_actions.addWidget(self.save_settings_btn)
        settings_actions.addWidget(self.save_settings_as_btn)
        connections_layout.addLayout(settings_actions)
        json_actions = QtWidgets.QHBoxLayout()
        self.export_settings_json_btn = QtWidgets.QPushButton("Export JSON")
        self.import_settings_json_btn = QtWidgets.QPushButton("Import JSON")
        json_actions.addWidget(self.export_settings_json_btn)
        json_actions.addWidget(self.import_settings_json_btn)
        json_actions.addStretch(1)
        connections_layout.addLayout(json_actions)
        self.settings_status = QtWidgets.QLabel()
        self.settings_status.setObjectName("profileBadge")
        self.settings_status.setWordWrap(True)
        self.settings_ids_status = QtWidgets.QLabel()
        self.settings_ids_status.setObjectName("profileBadge")
        self.settings_ids_status.setWordWrap(True)
        self.settings_motion_status = QtWidgets.QLabel()
        self.settings_motion_status.setObjectName("profileBadge")
        self.settings_motion_status.setWordWrap(True)
        connections_layout.addWidget(self.settings_status)
        connections_layout.addWidget(self.settings_ids_status)
        connections_layout.addWidget(self.settings_motion_status)
        toolbar = QtWidgets.QHBoxLayout()
        self.refresh_ports_btn = QtWidgets.QPushButton("Refresh Ports")
        self.read_status_btn = QtWidgets.QPushButton("Read Switches")
        self.read_position_btn = QtWidgets.QPushButton("Read Position")
        toolbar.addWidget(self.refresh_ports_btn)
        toolbar.addWidget(self.read_status_btn)
        toolbar.addWidget(self.read_position_btn)
        connections_layout.addLayout(toolbar)
        grid = QtWidgets.QGridLayout()
        grid.setHorizontalSpacing(8)
        grid.setVerticalSpacing(4)
        grid.setColumnStretch(1, 1)
        grid.setColumnStretch(3, 1)
        self._build_port_row(grid, 0, "x", "X Motor")
        self._build_port_row(grid, 1, "y", "Y Motor")
        self._build_port_row(grid, 2, "updown", "Z Axis")
        self._build_port_row(grid, 3, "sensor", "Reference")
        connections_layout.addLayout(grid)
        self.disable_updown_check = QtWidgets.QCheckBox("Disable Up/Down (XY-only mode)")
        self.disable_updown_check.setObjectName("updownModeCheckBox")
        self.disable_updown_check.setToolTip(
            "Isolate XY control from the Z axis for testing. This disables Z connection and Z jog/home controls and lets XY motion proceed without Z."
        )
        self.updown_mode_status = QtWidgets.QLabel()
        self.updown_mode_status.setObjectName("tableHint")
        self.updown_mode_status.setWordWrap(True)
        connections_layout.addWidget(self.disable_updown_check)
        connections_layout.addWidget(self.updown_mode_status)
        switches = QtWidgets.QGridLayout()
        switch_rows = (
            ("z_top", "Z TOP"),
            ("x_load", "X LOAD"),
            ("x_center", "X CENTER"),
            ("y_load", "Y LOAD"),
            ("y_center", "Y CENTER"),
        )
        for row, (key, title) in enumerate(switch_rows):
            label = QtWidgets.QLabel("Unknown")
            label.setMinimumWidth(72)
            self._switch_labels[key] = label
            switches.addWidget(QtWidgets.QLabel(title), row, 0)
            switches.addWidget(label, row, 1)
        connections_layout.addLayout(switches)

        motion_card, motion_layout = self._build_card(
            "Stage Controls",
            "Home Z first if the quartz rod is not clear, then move XY.",
        )
        metrics = QtWidgets.QGridLayout()
        self.x_pos_label = QtWidgets.QLabel("--")
        self.y_pos_label = QtWidgets.QLabel("--")
        self.z_pos_label = QtWidgets.QLabel("--")
        self.current_hole_label = QtWidgets.QLabel("Unknown")
        self.position_source_label = QtWidgets.QLabel("Unknown")
        self.position_source_label.setObjectName("tableHint")
        metrics.addWidget(QtWidgets.QLabel("X Position"), 0, 0)
        metrics.addWidget(self.x_pos_label, 0, 1)
        metrics.addWidget(QtWidgets.QLabel("Y Position"), 1, 0)
        metrics.addWidget(self.y_pos_label, 1, 1)
        metrics.addWidget(QtWidgets.QLabel("Z Position"), 2, 0)
        metrics.addWidget(self.z_pos_label, 2, 1)
        metrics.addWidget(QtWidgets.QLabel("Nearest Cup"), 3, 0)
        metrics.addWidget(self.current_hole_label, 3, 1)
        metrics.addWidget(QtWidgets.QLabel("Position Source"), 4, 0)
        metrics.addWidget(self.position_source_label, 4, 1)
        motion_layout.addLayout(metrics)

        selection = QtWidgets.QFormLayout()
        self.target_hole = QtWidgets.QSpinBox()
        self.target_hole.setRange(self.SLOT_MIN, self.SLOT_MAX)
        selection.addRow("Active Cup", self.target_hole)
        motion_layout.addLayout(selection)

        commands = QtWidgets.QGridLayout()
        self.goto_btn = QtWidgets.QPushButton("Move To Selected Cup")
        self.goto_btn.setObjectName("accent")
        self.goto_btn.setText("Move To Cup")
        self.home_top_btn = QtWidgets.QPushButton("Home Z")
        self.home_center_btn = QtWidgets.QPushButton("Home XY")
        self.corner_btn = QtWidgets.QPushButton("Load")
        self.refresh_live_btn = QtWidgets.QPushButton("Refresh")
        commands.addWidget(self.goto_btn, 0, 0, 1, 2)
        commands.addWidget(self.home_top_btn, 1, 0)
        commands.addWidget(self.home_center_btn, 1, 1)
        commands.addWidget(self.corner_btn, 2, 0)
        commands.addWidget(self.refresh_live_btn, 2, 1)
        motion_layout.addLayout(commands)

        left.addWidget(motion_card)

        cup_card, cup_layout = self._build_card(
            "Cup Position Sheet",
            "Compact editable X/Y count table inspired by the VB6 settings grid. Edit counts directly; blank both X and Y cells to clear a cup.",
        )
        self.cup_table_summary = QtWidgets.QLabel("0 cups stored. Capture XY or edit counts directly.")
        self.cup_table_summary.setObjectName("tableHint")
        self.cup_table_summary.setWordWrap(True)
        cup_layout.addWidget(self.cup_table_summary)
        self.cup_table = QtWidgets.QTableWidget(self.SLOT_MAX - self.SLOT_MIN + 1, 3)
        self.cup_table.setObjectName("cupTable")
        self.cup_table.setHorizontalHeaderLabels(["Cup", "X counts", "Y counts"])
        self.cup_table.verticalHeader().setVisible(False)
        self.cup_table.setAlternatingRowColors(True)
        self.cup_table.setSelectionBehavior(QtWidgets.QAbstractItemView.SelectRows)
        self.cup_table.setSelectionMode(QtWidgets.QAbstractItemView.SingleSelection)
        self.cup_table.setEditTriggers(
            QtWidgets.QAbstractItemView.DoubleClicked
            | QtWidgets.QAbstractItemView.EditKeyPressed
            | QtWidgets.QAbstractItemView.SelectedClicked
        )
        self.cup_table.setMinimumHeight(248)
        self.cup_table.setMaximumHeight(312)
        header = self.cup_table.horizontalHeader()
        header.setSectionResizeMode(0, QtWidgets.QHeaderView.ResizeToContents)
        header.setSectionResizeMode(1, QtWidgets.QHeaderView.Stretch)
        header.setSectionResizeMode(2, QtWidgets.QHeaderView.Stretch)
        cup_layout.addWidget(self.cup_table)
        left.addWidget(cup_card)
        left.addStretch(1)

        stage_card, stage_layout = self._build_card(
            "Stage Model",
            "Stage-first tray view matched to the photographed acrylic plate. Inset grip slots and center drop-off hole are drawn to match the physical holder.",
        )
        stage_help_text = (
            "Click a cup to choose the next target. Double-click a cup to run the same move as the Move To Cup button. "
            "Click the LOAD marker to select the loading corner and double-click it to move there. Click the center drop hole to select center and double-click it to home XY to center. "
            "Use clicks and double-clicks to work between LOAD, center, and the numbered cups. Use the side panels for connections, motion tuning, jogs, and calibration. "
            + self._stage_limit_instructions()
        )

        self.stage_scene = StageScene(self.SLOT_MIN, self.SLOT_MAX)
        self.stage_scene.setToolTip(stage_help_text)
        self.stage_status = QtWidgets.QLabel(
            "Click a cup to select it. Double-click a cup to move there. LOAD and center markers also respond to click and double-click."
        )
        self.stage_status.setObjectName("mapStatus")
        self.stage_status.setWordWrap(True)
        self.stage_status.setMinimumHeight(50)
        self.stage_status.setToolTip(stage_help_text)
        stage_layout.addWidget(self.stage_scene, stretch=1)
        stage_layout.addWidget(self.stage_status)
        center_layout.addWidget(stage_card, stretch=1)

        console_card, console_layout = self._build_card("Console")
        self.console = QtWidgets.QPlainTextEdit()
        self.console.setReadOnly(True)
        self.console.setObjectName("console")
        self.console.setMinimumHeight(96)
        self.console.setMaximumHeight(132)
        console_layout.addWidget(self.console)
        center_layout.addWidget(console_card)

        tuning_card, tuning_layout = self._build_card(
            "Velocity And Jog",
            "Keep motion tuning near the stage while leaving the tray itself visually dominant.",
        )
        self.model_help_btn = QtWidgets.QToolButton()
        self.model_help_btn.setText("?")
        self.model_help_btn.setAutoRaise(True)
        self.model_help_btn.setToolTip(stage_help_text)
        self.model_xy_speed = QtWidgets.QSpinBox()
        self.model_xy_speed.setRange(100_000, self.motion_defaults.changer_speed)
        self.model_xy_speed.setSingleStep(100_000)
        self.model_xy_speed.setGroupSeparatorShown(True)
        self.model_xy_speed.setValue(self._default_xy_speed_raw())
        self.model_xy_speed.setToolTip(
            "XY manual velocity in raw counts/s. The live text below converts the current raw value into estimated cm/s from the loaded cup spacing."
        )
        self.model_xy_speed.setWhatsThis(self.model_xy_speed.toolTip())
        self.xy_velocity_estimate = QtWidgets.QLabel()
        self.xy_velocity_estimate.setObjectName("tableHint")
        self.xy_velocity_estimate.setWordWrap(True)
        self.model_z_speed = QtWidgets.QSpinBox()
        self.model_z_speed.setRange(1, self.motion_defaults.lift_speed_fast)
        self.model_z_speed.setSingleStep(max(1, self.settings_profile.updown_motor_1cm // 10, 10))
        self.model_z_speed.setGroupSeparatorShown(True)
        self.model_z_speed.setValue(self._default_z_speed_raw())
        self.model_z_speed.setToolTip(
            "Z manual velocity in raw counts/s. The live text below converts the current raw value into estimated cm/s using UpDownMotor1cm."
        )
        self.model_z_speed.setWhatsThis(self.model_z_speed.toolTip())
        self.z_velocity_estimate = QtWidgets.QLabel()
        self.z_velocity_estimate.setObjectName("tableHint")
        self.z_velocity_estimate.setWordWrap(True)
        self.model_step = QtWidgets.QSpinBox()
        self.model_step.setRange(100, 2_000_000)
        self.model_step.setSingleStep(100)
        self.model_step.setValue(self._safe_jog_step())
        self.model_step.setToolTip(
            "Relative jog distance in counts. This starts at roughly a quarter of one stored cup pitch for safer manual motion."
        )
        self.model_x_minus_btn = QtWidgets.QPushButton("X-")
        self.model_x_plus_btn = QtWidgets.QPushButton("X+")
        self.model_y_minus_btn = QtWidgets.QPushButton("Y-")
        self.model_y_plus_btn = QtWidgets.QPushButton("Y+")
        self.model_z_minus_btn = QtWidgets.QPushButton("Z-")
        self.model_z_plus_btn = QtWidgets.QPushButton("Z+")
        self.model_home_xy_btn = QtWidgets.QPushButton("Home XY")
        self.model_home_z_btn = QtWidgets.QPushButton("Home Z")
        tuning_grid = QtWidgets.QGridLayout()
        tuning_grid.addWidget(QtWidgets.QLabel("XY velocity"), 0, 0)
        tuning_grid.addWidget(self.model_xy_speed, 0, 1)
        tuning_grid.addWidget(self.model_help_btn, 0, 2)
        tuning_grid.addWidget(self.xy_velocity_estimate, 1, 0, 1, 3)
        tuning_grid.addWidget(QtWidgets.QLabel("Z velocity"), 2, 0)
        tuning_grid.addWidget(self.model_z_speed, 2, 1)
        tuning_grid.addWidget(self.z_velocity_estimate, 3, 0, 1, 3)
        tuning_grid.addWidget(QtWidgets.QLabel("Jog step"), 4, 0)
        tuning_grid.addWidget(self.model_step, 4, 1)
        tuning_grid.addWidget(self.model_x_minus_btn, 5, 0)
        tuning_grid.addWidget(self.model_x_plus_btn, 5, 1)
        tuning_grid.addWidget(self.model_y_minus_btn, 6, 0)
        tuning_grid.addWidget(self.model_y_plus_btn, 6, 1)
        tuning_grid.addWidget(self.model_z_minus_btn, 7, 0)
        tuning_grid.addWidget(self.model_z_plus_btn, 7, 1)
        tuning_grid.addWidget(self.model_home_xy_btn, 8, 0)
        tuning_grid.addWidget(self.model_home_z_btn, 8, 1)
        tuning_grid.setColumnStretch(0, 1)
        tuning_grid.setColumnStretch(1, 1)
        tuning_layout.addLayout(tuning_grid)
        right.addWidget(tuning_card)

        calibration_card, calibration_layout = self._build_card(
            "Calibration",
            "Capture the current X/Y counts for each cup to build the tray lookup table.",
        )
        self.capture_btn = QtWidgets.QPushButton("Capture XY")
        self.capture_btn.setObjectName("accent")
        self.clear_calibration_btn = QtWidgets.QPushButton("Clear Cup")
        self.calibration_summary = QtWidgets.QLabel("0 calibrated cups")
        self.calibration_summary.setWordWrap(True)
        calibration_layout.addWidget(self.capture_btn)
        calibration_layout.addWidget(self.clear_calibration_btn)
        calibration_layout.addWidget(self.calibration_summary)
        right.addWidget(calibration_card)

        right.addWidget(connections_card)
        right.addStretch(1)

        shell.addWidget(left_scroll)
        shell.addWidget(center_host, stretch=1)
        shell.addWidget(right_scroll)

        self.refresh_ports_btn.clicked.connect(self._refresh_port_boxes)
        self.load_settings_btn.clicked.connect(self._load_settings_file)
        self.browse_settings_btn.clicked.connect(self._browse_settings_file)
        self.save_settings_btn.clicked.connect(self._save_settings_file)
        self.save_settings_as_btn.clicked.connect(self._save_settings_file_as)
        self.export_settings_json_btn.clicked.connect(self._export_settings_json)
        self.import_settings_json_btn.clicked.connect(self._import_settings_json)
        self.read_status_btn.clicked.connect(self._poll_stage_state)
        self.read_position_btn.clicked.connect(self._read_current_position)
        self.goto_btn.clicked.connect(self._move_to_selected_hole)
        self.home_top_btn.clicked.connect(self._home_top)
        self.home_center_btn.clicked.connect(self._home_xy_center)
        self.corner_btn.clicked.connect(self._move_xy_corner)
        self.refresh_live_btn.clicked.connect(self._poll_stage_state)
        self.capture_btn.clicked.connect(self._capture_calibration)
        self.clear_calibration_btn.clicked.connect(self._clear_calibration)
        self.target_hole.valueChanged.connect(self._sync_selected_hole)
        self.cup_table.itemChanged.connect(self._on_cup_table_item_changed)
        self.cup_table.currentCellChanged.connect(self._on_cup_table_current_cell_changed)
        self.stage_scene.holeSelected.connect(self._on_scene_hole_selected)
        self.stage_scene.holeActivated.connect(self._on_scene_hole_activated)
        self.stage_scene.specialTargetSelected.connect(self._on_scene_special_selected)
        self.stage_scene.specialTargetActivated.connect(self._on_scene_special_activated)
        self.model_x_minus_btn.clicked.connect(lambda: self._jog_axis("x", -self.model_step.value(), self._model_speed("x")))
        self.model_x_plus_btn.clicked.connect(lambda: self._jog_axis("x", self.model_step.value(), self._model_speed("x")))
        self.model_y_minus_btn.clicked.connect(lambda: self._jog_axis("y", -self.model_step.value(), self._model_speed("y")))
        self.model_y_plus_btn.clicked.connect(lambda: self._jog_axis("y", self.model_step.value(), self._model_speed("y")))
        self.model_z_minus_btn.clicked.connect(lambda: self._jog_axis("updown", -self.model_step.value(), self._model_speed("updown")))
        self.model_z_plus_btn.clicked.connect(lambda: self._jog_axis("updown", self.model_step.value(), self._model_speed("updown")))
        self.model_home_xy_btn.clicked.connect(self._home_xy_center)
        self.model_home_z_btn.clicked.connect(self._home_top)
        self.model_xy_speed.valueChanged.connect(lambda value: self._confirm_speed_if_needed("xy", int(value)))
        self.model_z_speed.valueChanged.connect(lambda value: self._confirm_speed_if_needed("updown", int(value)))
        self.disable_updown_check.toggled.connect(self._on_disable_updown_toggled)

    def _available_ports(self) -> list[str]:
        ports = sorted(port.device for port in list_ports.comports())
        return ports or [""]

    def _set_combo_value(self, combo: QtWidgets.QComboBox, value: str) -> None:
        current = combo.currentText()
        items = [combo.itemText(index) for index in range(combo.count())]
        if value and value not in items:
            combo.addItem(value)
        combo.setCurrentText(value or current)

    def _refresh_port_boxes(self) -> None:
        ports = self._available_ports()
        for combo in self._port_boxes.values():
            current = combo.currentText()
            combo.blockSignals(True)
            combo.clear()
            combo.addItems(ports)
            combo.setCurrentText(current)
            combo.blockSignals(False)

    def _load_config_into_widgets(self) -> None:
        self._refresh_port_boxes()
        self._set_combo_value(self._port_boxes["x"], self.config.x_port)
        self._set_combo_value(self._port_boxes["y"], self.config.y_port)
        self._set_combo_value(self._port_boxes["updown"], self.config.updown_port)
        self._set_combo_value(self._port_boxes["sensor"], self.config.sensor_port)
        self.disable_updown_check.blockSignals(True)
        self.disable_updown_check.setChecked(self.config.updown_disabled)
        self.disable_updown_check.blockSignals(False)
        self.settings_path_edit.setText(str(self.settings_profile.path))
        self.target_hole.setValue(self.config.selected_hole)
        self._apply_settings_profile(self.settings_profile)
        self._update_stage_calibrations()
        self._apply_updown_mode()
        self.stage_scene.set_target_hole(self.config.selected_hole)
        self._sync_cup_table_selection(self.config.selected_hole)

    def _model_speed(self, axis: str) -> int:
        if axis == "updown":
            return int(self.model_z_speed.value())
        return int(self.model_xy_speed.value())

    def _append(self, message: str) -> None:
        self.console.appendPlainText(message)

    def _connect_port(self, key: str) -> None:
        if key == "updown" and self._xy_only_mode_enabled():
            QtWidgets.QMessageBox.information(
                self,
                "Up/Down Disabled",
                "XY-only mode is active. Re-enable up/down before connecting or using the Z axis.",
            )
            self._set_port_toggle_state(key, False)
            return
        port = self._port_boxes[key].currentText().strip()
        try:
            if key == "sensor":
                self.controller.connect_sensor(port)
            else:
                self.controller.connect_axis(key, port)
        except Exception as exc:
            self._set_port_toggle_state(key, False)
            QtWidgets.QMessageBox.warning(self, "Connection Error", str(exc))
            return
        self._port_status[key].setText(f"On {port}")
        self._set_port_toggle_state(key, True)
        self._append(f"Connected {key.upper()} on {port}.")
        self._save_config()
        self._poll_stage_state()

    def _disconnect_port(self, key: str) -> None:
        if key == "sensor":
            self.controller.disconnect_sensor()
        else:
            self.controller.disconnect_axis(key)
        self._port_status[key].setText("Off")
        self._set_port_toggle_state(key, False)
        self._append(f"Disconnected {key.upper()}.")
        self._save_config()
        self._poll_stage_state()

    def _sync_selected_hole(self, value: int) -> None:
        self.stage_scene.set_target_hole(value)
        self._sync_cup_table_selection(value)
        self._save_config()

    def _calibration_map(self) -> dict[int, tuple[int, int]]:
        mapping: dict[int, tuple[int, int]] = {}
        for item in self.config.calibrations:
            if item.x is not None and item.y is not None:
                mapping[item.hole] = (int(item.x), int(item.y))
        return mapping

    def _set_calibration(self, hole: int, x_pos: int, y_pos: int) -> None:
        updated = False
        for item in self.config.calibrations:
            if item.hole == hole:
                item.x = x_pos
                item.y = y_pos
                updated = True
                break
        if not updated:
            self.config.calibrations.append(HoleCalibration(hole=hole, x=x_pos, y=y_pos))
        self._update_stage_calibrations()
        self._save_config()

    def _clear_calibration(self) -> None:
        hole = self.target_hole.value()
        self.config.calibrations = [item for item in self.config.calibrations if item.hole != hole]
        self._append(f"Cleared cup {hole} calibration.")
        self._update_stage_calibrations()
        self._save_config()

    def _update_stage_calibrations(self) -> None:
        mapping = self._calibration_map()
        self.stage_scene.set_calibrations(mapping)
        self.calibration_summary.setText(f"{len(mapping)} calibrated cups")
        self.cup_table_summary.setText(
            f"{len(mapping)} cups stored. Double-click X or Y to edit counts; blank both cells to clear a cup."
        )
        self._refresh_cup_table()

    def _refresh_cup_table(self) -> None:
        blocker = QtCore.QSignalBlocker(self.cup_table)
        del blocker
        mapping = self._calibration_map()
        self.cup_table.blockSignals(True)
        try:
            for row, hole in enumerate(range(self.SLOT_MIN, self.SLOT_MAX + 1)):
                cup_item = self.cup_table.item(row, 0)
                if cup_item is None:
                    cup_item = QtWidgets.QTableWidgetItem()
                    cup_item.setFlags(QtCore.Qt.ItemIsSelectable | QtCore.Qt.ItemIsEnabled)
                    cup_item.setTextAlignment(QtCore.Qt.AlignCenter)
                    self.cup_table.setItem(row, 0, cup_item)
                cup_item.setText(str(hole))

                x_item = self.cup_table.item(row, 1)
                if x_item is None:
                    x_item = QtWidgets.QTableWidgetItem()
                    x_item.setTextAlignment(QtCore.Qt.AlignRight | QtCore.Qt.AlignVCenter)
                    self.cup_table.setItem(row, 1, x_item)

                y_item = self.cup_table.item(row, 2)
                if y_item is None:
                    y_item = QtWidgets.QTableWidgetItem()
                    y_item.setTextAlignment(QtCore.Qt.AlignRight | QtCore.Qt.AlignVCenter)
                    self.cup_table.setItem(row, 2, y_item)

                if hole in mapping:
                    x_val, y_val = mapping[hole]
                    x_item.setText(str(x_val))
                    y_item.setText(str(y_val))
                else:
                    x_item.setText("")
                    y_item.setText("")
            self._sync_cup_table_selection(self.target_hole.value())
        finally:
            self.cup_table.blockSignals(False)

    def _sync_cup_table_selection(self, hole: int) -> None:
        if not hasattr(self, "cup_table"):
            return
        row = max(0, min(self.SLOT_MAX - self.SLOT_MIN, hole - self.SLOT_MIN))
        self.cup_table.blockSignals(True)
        try:
            self.cup_table.setCurrentCell(row, 0)
            self.cup_table.selectRow(row)
        finally:
            self.cup_table.blockSignals(False)

    def _on_cup_table_current_cell_changed(self, current_row: int, current_column: int, previous_row: int, previous_column: int) -> None:
        del current_column, previous_row, previous_column
        if current_row < 0:
            return
        hole = self.SLOT_MIN + current_row
        if self.target_hole.value() != hole:
            self.target_hole.setValue(hole)

    def _on_cup_table_item_changed(self, item: QtWidgets.QTableWidgetItem) -> None:
        if item.column() not in (1, 2):
            return
        row = item.row()
        hole = self.SLOT_MIN + row
        x_item = self.cup_table.item(row, 1)
        y_item = self.cup_table.item(row, 2)
        x_text = "" if x_item is None else x_item.text().strip()
        y_text = "" if y_item is None else y_item.text().strip()

        if not x_text and not y_text:
            self.config.calibrations = [entry for entry in self.config.calibrations if entry.hole != hole]
            self.stage_status.setText(f"Cleared cup {hole} from the cup-position sheet.")
            self._update_stage_calibrations()
            self._save_config()
            return

        if not x_text or not y_text:
            self.stage_status.setText(f"Cup {hole} needs both X and Y counts before it can be saved.")
            return

        try:
            x_pos = int(x_text)
            y_pos = int(y_text)
        except ValueError:
            self.stage_status.setText(f"Cup {hole} counts must be whole integers.")
            self._refresh_cup_table()
            return

        self._set_calibration(hole, x_pos, y_pos)
        self.stage_status.setText(f"Saved cup {hole} counts from the cup-position sheet.")

    def _nearest_hole_from_xy(self, x_pos: int | None, y_pos: int | None) -> int | None:
        if x_pos is None or y_pos is None:
            return None
        for hole, (hole_x, hole_y) in self._calibration_map().items():
            if abs(x_pos - hole_x) < self.POSITION_TOLERANCE and abs(y_pos - hole_y) < self.POSITION_TOLERANCE:
                return hole
        return None

    def _poll_stage_state(self) -> None:
        for key in ("x", "y", "updown"):
            if key == "updown" and self._xy_only_mode_enabled():
                self._set_port_toggle_state(key, False)
                self._port_status[key].setText("Disabled")
                continue
            connected = self.controller.axis_connected(key)
            self._set_port_toggle_state(key, connected)
            self._port_status[key].setText(f"On {self.controller.axis_port(key)}" if connected else "Off")
        self._set_port_toggle_state("sensor", self.controller.sensor_connected)
        self._port_status["sensor"].setText(f"On {self.controller.sensor_port}" if self.controller.sensor_connected else "Off")
        try:
            x_pos, y_pos = self.controller.read_xy_position()
            switches = self.controller.reference_switch_states()
            z_pos = self.controller.read_axis_position("updown") if self.controller.axis_connected("updown") and not self._xy_only_mode_enabled() else None
            if self._xy_only_mode_enabled():
                switches["z_top"] = None
        except Exception as exc:
            x_pos, y_pos = None, None
            z_pos = None
            switches = {"z_top": None, "x_load": None, "x_center": None, "y_load": None, "y_center": None}
            message = str(exc)
            if message and message != self._last_poll_error:
                self._last_poll_error = message
                self._append(f"Polling warning: {message}")
        else:
            self._last_poll_error = ""

        if x_pos is not None or y_pos is not None:
            self._record_switch_anchor_positions(x_pos, y_pos, switches)
            if x_pos is not None:
                self._last_live_positions["x"] = x_pos
            if y_pos is not None:
                self._last_live_positions["y"] = y_pos
            if z_pos is not None:
                self._last_live_positions["updown"] = z_pos
            self._position_source = "live read"
        else:
            x_pos, y_pos, source = self._infer_xy_from_switches(switches)
            if z_pos is None and not self._xy_only_mode_enabled():
                z_pos = self._last_live_positions["updown"]
            self._position_source = source if x_pos is not None or y_pos is not None else "unknown"

        self.x_pos_label.setText("--" if x_pos is None else str(x_pos))
        self.y_pos_label.setText("--" if y_pos is None else str(y_pos))
        self.z_pos_label.setText("--" if z_pos is None else str(z_pos))
        self.position_source_label.setText(self._position_source.title())
        current_hole = self._nearest_hole_from_xy(x_pos, y_pos)
        self.current_hole_label.setText("Unknown" if current_hole is None else str(current_hole))
        self.stage_scene.set_current_hole(current_hole)
        self.stage_scene.set_current_xy(x_pos, y_pos)
        self.stage_scene.set_loading_position_active(self._loading_position_active(x_pos, y_pos, switches))
        self._current_stage_hole = current_hole
        self._current_z_pos = z_pos

        for key, label in self._switch_labels.items():
            if key == "z_top" and self._xy_only_mode_enabled():
                label.setText("Disabled")
                continue
            state = switches.get(key)
            if state is None:
                label.setText("Unknown")
            else:
                label.setText("TRIPPED" if state else "Released")

    def _home_top(self) -> None:
        if self._xy_only_mode_enabled():
            QtWidgets.QMessageBox.information(self, "Up/Down Disabled", "XY-only mode is active. Re-enable up/down before using Z homing.")
            return
        try:
            result = self.controller.home_to_top()
        except Exception as exc:
            QtWidgets.QMessageBox.warning(self, "Home Z Error", str(exc))
            return
        self._last_live_positions["updown"] = result.final_position
        self._append(f"HomeToTop complete: z={result.final_position}, success={result.success}.")
        self._poll_stage_state()

    def _confirm_xy_motion_without_z(self, action: str) -> bool:
        if self._xy_only_mode_enabled():
            return True
        if self.controller.axis_connected("updown"):
            return True
        response = QtWidgets.QMessageBox.question(
            self,
            "Proceed Without Z Axis?",
            "Z axis motor is not connected.\n\n"
            f"Before continuing with {action}, confirm that this test is intentionally running without a Z motor connected and that the sample holder rod is raised clear above the XY stage.\n\n"
            "Continue with XY motion?",
            QtWidgets.QMessageBox.StandardButton.Yes | QtWidgets.QMessageBox.StandardButton.No,
            QtWidgets.QMessageBox.StandardButton.No,
        )
        if response != QtWidgets.QMessageBox.StandardButton.Yes:
            return False
        self._append(f"Proceeding with {action} without Z axis connection after operator confirmation.")
        return True

    def _home_xy_center(self) -> None:
        allow_without_updown = not self.controller.axis_connected("updown")
        if allow_without_updown and not self._confirm_xy_motion_without_z("Home XY"):
            return
        try:
            x_res, y_res = self.controller.home_xy_to_center(allow_without_updown=allow_without_updown)
        except Exception as exc:
            QtWidgets.QMessageBox.warning(self, "Home XY Error", str(exc))
            return
        self._append(
            "HomeToCenter complete: "
            f"x={x_res.final_position}, y={y_res.final_position}, success={x_res.success and y_res.success}."
        )
        self._last_live_positions["x"] = x_res.final_position
        self._last_live_positions["y"] = y_res.final_position
        self._poll_stage_state()

    def _move_xy_corner(self) -> None:
        allow_without_updown = not self.controller.axis_connected("updown")
        if allow_without_updown and not self._confirm_xy_motion_without_z("Move To Corner"):
            return
        try:
            x_res, y_res = self.controller.move_xy_to_corner(allow_without_updown=allow_without_updown)
        except Exception as exc:
            QtWidgets.QMessageBox.warning(self, "Corner Move Error", str(exc))
            return
        self._last_live_positions["x"] = x_res.final_position
        self._last_live_positions["y"] = y_res.final_position
        self._append(f"MoveToCorner complete: x={x_res.final_position}, y={y_res.final_position}.")
        self._poll_stage_state()

    def _move_to_selected_hole(self) -> None:
        hole = self.target_hole.value()
        calibration = self._calibration_map().get(hole)
        if calibration is None:
            QtWidgets.QMessageBox.warning(
                self,
                "Uncalibrated Cup",
                f"Cup {hole} does not have a stored X/Y calibration yet. Capture the current position first.",
            )
            return
        allow_without_updown = not self.controller.axis_connected("updown")
        if allow_without_updown and not self._confirm_xy_motion_without_z(f"Move To Cup {hole}"):
            return
        try:
            x_res, y_res = self.controller.move_xy_absolute(
                calibration[0],
                calibration[1],
                allow_without_updown=allow_without_updown,
            )
        except Exception as exc:
            QtWidgets.QMessageBox.warning(self, "Move Error", str(exc))
            return
        self._append(
            f"Moved to cup {hole}: x_target={x_res.target}, x_final={x_res.final_position}, "
            f"y_target={y_res.target}, y_final={y_res.final_position}."
        )
        self._last_live_positions["x"] = x_res.final_position
        self._last_live_positions["y"] = y_res.final_position
        self.stage_status.setText(f"Cup {hole} is the active target.")
        self._poll_stage_state()

    def _capture_calibration(self) -> None:
        try:
            x_pos, y_pos = self.controller.read_xy_position()
        except Exception as exc:
            QtWidgets.QMessageBox.warning(self, "Read Position Error", str(exc))
            return
        if x_pos is None or y_pos is None:
            QtWidgets.QMessageBox.warning(self, "Read Position Error", "Connect X and Y before capturing calibration.")
            return
        self._last_live_positions["x"] = x_pos
        self._last_live_positions["y"] = y_pos
        hole = self.target_hole.value()
        self._set_calibration(hole, x_pos, y_pos)
        self._append(f"Captured cup {hole} calibration: x={x_pos}, y={y_pos}.")
        self.stage_status.setText(f"Cup {hole} calibration updated from the live stage position.")
        self._poll_stage_state()

    def _select_scene_hole(self, hole: int) -> str:
        self.target_hole.setValue(hole)
        return "measurement drop-off" if hole == StageScene.CENTER_DROP_HOLE else f"cup {hole}"

    def _on_scene_hole_selected(self, hole: int) -> None:
        label = self._select_scene_hole(hole)
        self.stage_status.setText(f"Selected {label}. Double-click to move there.")

    def _on_scene_hole_activated(self, hole: int) -> None:
        label = self._select_scene_hole(hole)
        self.stage_status.setText(f"Moving to {label}.")
        self._move_to_selected_hole()

    def _on_scene_special_selected(self, target: str) -> None:
        self.stage_scene.set_special_target(target)
        if target == StageScene.LOAD_TARGET:
            self.stage_status.setText("Selected loading position. Double-click LOAD to move to the loading corner.")
            return
        self.stage_status.setText("Selected center reference. Double-click the center hole to home XY to center.")

    def _on_scene_special_activated(self, target: str) -> None:
        self.stage_scene.set_special_target(target)
        if target == StageScene.LOAD_TARGET:
            self.stage_status.setText("Moving to loading position.")
            self._move_xy_corner()
            return
        self.stage_status.setText("Homing XY to center.")
        self._home_xy_center()

    def _jog_axis(self, key: str, delta: int, velocity: int | None = None) -> None:
        if key == "updown" and self._xy_only_mode_enabled():
            QtWidgets.QMessageBox.information(self, "Up/Down Disabled", "XY-only mode is active. Re-enable up/down before jogging Z.")
            return
        try:
            result = self.controller.move_axis_relative(key, delta, velocity=velocity)
        except Exception as exc:
            QtWidgets.QMessageBox.warning(self, "Jog Error", str(exc))
            return
        self._last_live_positions[key] = result.final_position
        self._append(f"Jogged {key.upper()} by {delta}: final={result.final_position}.")
        self._poll_stage_state()

    def _render_overlay_payload(self) -> str:
        target = self.target_hole.value()
        current = self._current_stage_hole
        target_text = "Cup 46 / measurement drop-off" if target == StageScene.CENTER_DROP_HOLE else f"Cup {target}"
        if current is None:
            current_text = "Unknown"
        elif current == StageScene.CENTER_DROP_HOLE:
            current_text = "Cup 46 / measurement drop-off"
        else:
            current_text = f"Cup {current}"
        return (
            "Stage overlay: "
            f"target={target_text}; current={current_text}; calibrated={len(self._calibration_map())} cups"
        )

    def closeEvent(self, event: QtGui.QCloseEvent) -> None:
        self._save_config()
        self.controller.shutdown()
        super().closeEvent(event)


def main() -> int:
    app = QtWidgets.QApplication(sys.argv)
    apply_liquid_glass_theme(app)
    assets_dir = Path(__file__).resolve().parent.parent / "assets"
    set_app_icon(app, "changer_xy_control_icon.png", assets_dir)
    window = MainWindow()
    set_app_icon(window, "changer_xy_control_icon.png", assets_dir)
    window.show()
    return app.exec()
