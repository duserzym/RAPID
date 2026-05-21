from __future__ import annotations

import configparser
import json
import math
import re
import shutil
import sys
import time
from dataclasses import asdict, dataclass
from pathlib import Path

import serial
from PySide6 import QtCore, QtGui, QtWidgets
from serial.tools import list_ports

try:
    import pyqtgraph as pg
except Exception:  # pragma: no cover - optional at runtime
    pg = None


def _bootstrap_common_imports() -> None:
    root = Path(__file__).resolve().parents[2]
    if str(root) not in sys.path:
        sys.path.insert(0, str(root))


_bootstrap_common_imports()
from rapidpy_common.hardware import HardwareError, MotorAxisConfig, MotorControllerConfig, MotorSerialClient, MoveResult  # noqa: E402
from rapidpy_common.ui import apply_card_shadow, apply_liquid_glass_theme, set_app_icon  # noqa: E402


APP_SETTINGS_PATH = Path.home() / ".rapidpy_updown_settings.json"
DEFAULT_SETTINGS_PATH = Path(__file__).resolve().parents[3] / "VB6" / "settings" / "Paleomag_v3.INI"
TOP_SWITCH_BIT = 4
HIGH_SPEED_CONFIRM_CM_PER_SEC = 10.0
DEFAULT_MANUAL_SPEED_CM_PER_SEC = 1.0
POSITION_TOLERANCE_COUNTS = 150
FLOAT_RE = re.compile(r"[-+]?\d*\.?\d+(?:[eE][-+]?\d+)?")


@dataclass(slots=True)
class UpDownSettings:
    motor_port: str = ""
    squid_port: str = ""
    squid_baud: int = 1200
    settings_path: str = str(DEFAULT_SETTINGS_PATH)
    min_raw_count: int = 0
    max_raw_count: int = 40000
    pickup_raw: int = 425_000
    dropoff_raw: int = 582_500
    susceptibility_meter_raw: int = 20_000
    z_velocity_raw: int = 2_500_000
    jog_step_raw: int = 4000
    target_raw: int = 0
    sample_height_cm: float = 2.54
    scan_half_range_cm: float = 2.0
    scan_step_cm: float = 0.1
    scan_settle_s: float = 1.0


@dataclass(slots=True)
class SettingsProfile:
    path: Path
    motion_defaults: MotorControllerConfig
    updown_axis: MotorAxisConfig
    updown_motor_1cm: int
    zero_pos: int
    meas_pos: int
    af_pos: int
    irm_pos: int
    scoil_pos: int
    floor_pos: int
    sample_bottom: int
    sample_top: int
    sample_height_counts: int


@dataclass(slots=True)
class ProfileBand:
    label: str
    raw_top: int
    raw_bottom: int
    width: float
    fill_color: str
    outline_color: str
    side: str = "right"
    value_text: str | None = None


@dataclass(slots=True)
class ProfileIndicator:
    label: str
    raw_position: int
    side: str
    color: str
    value_text: str | None = None
    style: str = "bar"
    symbol: str = "dot"
    bar_half_width: float = 18.0
    emphasis: bool = False


@dataclass(slots=True)
class ProfileModel:
    range_top: int
    range_bottom: int
    top_switch_raw: int
    holder_bottom_raw: int
    sample_top_raw: int
    sample_bottom_raw: int
    zero_raw: int
    floor_raw: int
    live_raw: int | None
    bands: tuple[ProfileBand, ...]
    indicators: tuple[ProfileIndicator, ...]


@dataclass(slots=True)
class SquidCalibration:
    xcal: float = -3.410
    ycal: float = -3.470
    zcal: float = -2.516
    range_fact: float = 1e-5


@dataclass(slots=True)
class ScanPoint:
    index: int
    raw_position: int
    z_cm: float
    x_emu: float
    y_emu: float
    z_emu: float
    moment_emu: float


@dataclass(slots=True)
class ScanResult:
    points: list[ScanPoint]
    suggested_z_cm: float | None
    suggested_target_raw: int | None
    suggested_meas_pos_raw: int | None
    fit_method: str
    note: str = ""


class SquidCommunicationError(RuntimeError):
    pass


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


def load_settings(path: Path = APP_SETTINGS_PATH) -> UpDownSettings:
    if not path.exists():
        return UpDownSettings()
    try:
        payload = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return UpDownSettings()

    merged = asdict(UpDownSettings())
    for key in merged:
        if key in payload:
            merged[key] = payload[key]
    return UpDownSettings(**merged)


def save_settings(settings: UpDownSettings, path: Path = APP_SETTINGS_PATH) -> None:
    try:
        path.write_text(json.dumps(asdict(settings), indent=2, sort_keys=True), encoding="utf-8")
    except OSError:
        return


def _load_settings_profile(settings_path: Path) -> SettingsProfile:
    config = _load_settings_config(settings_path)

    motor_section = config["SteppingMotor"] if config.has_section("SteppingMotor") else {}
    program_section = config["MotorPrograms"] if config.has_section("MotorPrograms") else {}

    defaults = MotorControllerConfig(
        changer_speed=_parse_ini_int(motor_section.get("ChangerSpeed"), MotorControllerConfig().changer_speed),
        turner_speed=_parse_ini_int(motor_section.get("TurnerSpeed"), MotorControllerConfig().turner_speed),
        turning_motor_full_rotation=_parse_ini_int(
            motor_section.get("TurningMotorFullRotation"),
            MotorControllerConfig().turning_motor_full_rotation,
        ),
        turning_motor_1rps=_parse_ini_int(motor_section.get("TurningMotor1rps"), MotorControllerConfig().turning_motor_1rps),
        lift_speed_slow=_parse_ini_int(motor_section.get("LiftSpeedSlow"), MotorControllerConfig().lift_speed_slow),
        lift_speed_normal=_parse_ini_int(motor_section.get("LiftSpeedNormal"), MotorControllerConfig().lift_speed_normal),
        lift_speed_fast=_parse_ini_int(motor_section.get("LiftSpeedFast"), MotorControllerConfig().lift_speed_fast),
        lift_acceleration=_parse_ini_int(motor_section.get("LiftAcceleration"), MotorControllerConfig().lift_acceleration),
        meas_pos=_parse_ini_int(motor_section.get("MeasPos"), MotorControllerConfig().meas_pos),
        sample_bottom=_parse_ini_int(motor_section.get("SampleBottom"), MotorControllerConfig().sample_bottom),
        sample_height=_parse_ini_int(
            motor_section.get("SampleTop"), MotorControllerConfig().sample_bottom + MotorControllerConfig().sample_height,
        ) - _parse_ini_int(motor_section.get("SampleBottom"), MotorControllerConfig().sample_bottom),
        updown_torque_factor=_parse_ini_int(
            motor_section.get("UpDownTorqueFactor"),
            MotorControllerConfig().updown_torque_factor,
        ),
        pickup_torque_throttle=float(
            _parse_ini_number(motor_section.get("PickupTorqueThrottle"), MotorControllerConfig().pickup_torque_throttle)
        ),
    )
    sample_bottom = defaults.sample_bottom
    sample_top = _parse_ini_int(motor_section.get("SampleTop"), sample_bottom + defaults.sample_height)
    sample_height_counts = sample_top - sample_bottom
    updown_address = _parse_ini_int(program_section.get("MotorIDUpDown"), 16)
    updown_motor_1cm = _parse_ini_int(motor_section.get("UpDownMotor1cm"), 0)
    zero_pos = _parse_ini_int(motor_section.get("ZeroPos"), -50_000)
    meas_pos = _parse_ini_int(motor_section.get("MeasPos"), defaults.meas_pos)
    af_pos = _parse_ini_int(motor_section.get("AFPos"), -42_500)
    irm_pos = _parse_ini_int(motor_section.get("IRMPos"), -36_000)
    scoil_pos = _parse_ini_int(motor_section.get("SCoilPos"), -22_700)
    floor_pos = _parse_ini_int(motor_section.get("FloorPos"), -148_955)
    return SettingsProfile(
        path=settings_path,
        motion_defaults=defaults,
        updown_axis=MotorAxisConfig(name="UpDown", motor_id=3, address=updown_address),
        updown_motor_1cm=updown_motor_1cm,
        zero_pos=zero_pos,
        meas_pos=meas_pos,
        af_pos=af_pos,
        irm_pos=irm_pos,
        scoil_pos=scoil_pos,
        floor_pos=floor_pos,
        sample_bottom=sample_bottom,
        sample_top=sample_top,
        sample_height_counts=sample_height_counts,
    )


class VerticalProfileWidget(QtWidgets.QWidget):
    def __init__(self, parent: QtWidgets.QWidget | None = None) -> None:
        super().__init__(parent)
        self._model: ProfileModel | None = None
        self._scan_points: tuple[ScanPoint, ...] = ()
        self._scan_center_cm = 0.0
        self._scan_half_range_cm = 0.25
        self._suggested_z_cm: float | None = None
        self._suggested_target_raw: int | None = None
        self.setMinimumHeight(560)
        self.setSizePolicy(QtWidgets.QSizePolicy.Policy.Expanding, QtWidgets.QSizePolicy.Policy.Expanding)

    def set_profile(self, model: ProfileModel) -> None:
        self._model = model
        self.update()

    def set_scan_detail(
        self,
        points: list[ScanPoint],
        center_z_cm: float,
        half_range_cm: float,
        suggested_z_cm: float | None,
        suggested_target_raw: int | None,
    ) -> None:
        self._scan_points = tuple(points)
        self._scan_center_cm = center_z_cm
        self._scan_half_range_cm = max(half_range_cm, 0.05)
        self._suggested_z_cm = suggested_z_cm
        self._suggested_target_raw = suggested_target_raw
        self.update()

    def _adjust_label_positions(self, targets: list[float], top: float, bottom: float, gap: float) -> list[float]:
        adjusted: list[float] = []
        for target in targets:
            y_pos = target if not adjusted else max(target, adjusted[-1] + gap)
            adjusted.append(y_pos)
        for index in range(len(adjusted) - 2, -1, -1):
            max_y = adjusted[index + 1] - gap
            if adjusted[index] > max_y:
                adjusted[index] = max_y
        shift = 0.0
        if adjusted:
            if adjusted[0] < top:
                shift = top - adjusted[0]
            if adjusted[-1] > bottom:
                shift = min(shift, bottom - adjusted[-1]) if shift else bottom - adjusted[-1]
        if shift:
            adjusted = [value + shift for value in adjusted]
        return adjusted

    def paintEvent(self, event: QtGui.QPaintEvent) -> None:  # noqa: N802
        del event
        model = self._model
        if model is None or model.range_top == model.range_bottom:
            return

        painter = QtGui.QPainter(self)
        painter.setRenderHint(QtGui.QPainter.RenderHint.Antialiasing)
        rect = self.rect().adjusted(12, 10, -18, -10)
        painter.fillRect(self.rect(), QtCore.Qt.GlobalColor.transparent)

        panel = QtCore.QRectF(rect)
        painter.setPen(QtCore.Qt.PenStyle.NoPen)
        painter.setBrush(QtGui.QColor("#fffaf4"))
        painter.drawRoundedRect(panel, 24, 24)

        range_span = float(max(1, model.range_top - model.range_bottom))
        chart_top = panel.top() + 18
        chart_bottom = panel.bottom() - 18
        usable_top = chart_top + 20
        usable_bottom = chart_bottom - 18
        center_x = panel.center().x() - 38

        def to_y(raw_position: float) -> float:
            ratio = (model.range_top - raw_position) / range_span
            return usable_top + ratio * (usable_bottom - usable_top)

        top_band_raw = max(band.raw_top for band in model.bands)
        bottom_band_raw = min(band.raw_bottom for band in model.bands)
        top_margin_raw = max(1400, int(abs(top_band_raw - model.holder_bottom_raw) * 0.12))
        bottom_margin_raw = max(900, int(abs(bottom_band_raw - top_band_raw) * 0.08))
        outer_top_raw = top_band_raw + top_margin_raw
        outer_bottom_raw = bottom_band_raw - bottom_margin_raw
        outer_top_y = to_y(outer_top_raw)
        outer_bottom_y = to_y(outer_bottom_raw)
        outer_body = QtCore.QRectF(center_x - 68, outer_top_y, 136, max(outer_bottom_y - outer_top_y, 140.0))
        body_rect = outer_body.adjusted(10, 0, -10, 0)
        painter.setBrush(QtGui.QColor("#111111"))
        painter.drawRoundedRect(body_rect, 62, 62)

        shield_rect = body_rect.adjusted(18, 14, -18, -14)
        shield_gradient = QtGui.QLinearGradient(shield_rect.topLeft(), shield_rect.bottomRight())
        shield_gradient.setColorAt(0.0, QtGui.QColor("#fbf5ea"))
        shield_gradient.setColorAt(1.0, QtGui.QColor("#eadfce"))
        painter.setBrush(shield_gradient)
        painter.drawRoundedRect(shield_rect, 40, 40)

        bore_rect = QtCore.QRectF(center_x - 13, outer_top_y + 18, 26, max(outer_bottom_y - outer_top_y - 36, 120.0))
        painter.setBrush(QtGui.QColor("#2b1d18"))
        painter.drawRoundedRect(bore_rect, 14, 14)
        painter.setBrush(QtGui.QColor("#4a3128"))
        painter.drawRoundedRect(bore_rect.adjusted(6, 10, -6, -10), 8, 8)

        zone_bands = sorted(model.bands, key=lambda band: band.raw_bottom, reverse=True)
        band_height = 9.0
        for band in zone_bands:
            center_y = to_y(band.raw_bottom)
            band_rect = QtCore.QRectF(center_x - band.width / 2.0, center_y - band_height / 2.0, band.width, band_height)
            fill = QtGui.QLinearGradient(band_rect.topLeft(), band_rect.bottomLeft())
            fill.setColorAt(0.0, QtGui.QColor(band.fill_color).lighter(112))
            fill.setColorAt(1.0, QtGui.QColor(band.fill_color).darker(112))
            painter.setPen(QtGui.QPen(QtGui.QColor(band.outline_color), 1.2))
            painter.setBrush(fill)
            painter.drawRoundedRect(band_rect, 3.0, 3.0)

        zero_y = to_y(model.zero_raw)
        painter.setPen(QtGui.QPen(QtGui.QColor("#8a6a44"), 1.0, QtCore.Qt.PenStyle.DashLine))
        painter.drawLine(QtCore.QPointF(bore_rect.left() - 18, zero_y), QtCore.QPointF(bore_rect.right() + 18, zero_y))

        holder_top_y = to_y(model.top_switch_raw) + 10
        holder_bottom_y = to_y(model.holder_bottom_raw)
        holder_rect = QtCore.QRectF(center_x - 5.5, holder_top_y, 11.0, max(holder_bottom_y - holder_top_y - 6.0, 18.0))
        painter.setPen(QtGui.QPen(QtGui.QColor("#725243"), 1.0))
        painter.setBrush(QtGui.QColor("#bea48f"))
        painter.drawRoundedRect(holder_rect, 5.5, 5.5)

        sample_bottom_y = to_y(model.sample_bottom_raw)
        stage_rect = QtCore.QRectF(center_x - 23.0, sample_bottom_y - 4.5, 46.0, 9.0)
        painter.setPen(QtGui.QPen(QtGui.QColor("#31566d"), 1.1))
        painter.setBrush(QtGui.QColor("#7fa2b7"))
        painter.drawRoundedRect(stage_rect, 4.5, 4.5)

        measurement_band = next((band for band in model.bands if "Measurement" in band.label), None)
        if measurement_band is not None:
            measurement_y = to_y(measurement_band.raw_bottom)
            plot_width = min(104.0, max(84.0, body_rect.left() - panel.left() - 146.0))
            plot_height = min(
                222.0,
                max(156.0, abs(to_y(measurement_band.raw_top) - measurement_y) * 3.4),
            )
            plot_left = max(panel.left() + 112.0, body_rect.left() - plot_width - 34.0)
            plot_top = min(
                max(measurement_y - plot_height / 2.0, panel.top() + 116.0),
                panel.bottom() - plot_height - 28.0,
            )
            plot_rect = QtCore.QRectF(plot_left, plot_top, plot_width, plot_height)
            plot_inner = plot_rect.adjusted(24.0, 18.0, -10.0, -24.0)

            moments = [point.moment_emu for point in self._scan_points]
            x_min = min(moments, default=0.0)
            x_max = max(moments, default=0.45)
            x_min = min(0.0, x_min)
            if math.isclose(x_max, x_min):
                x_max = x_min + max(abs(x_max) * 0.35, 0.35)
            else:
                x_max += max((x_max - x_min) * 0.12, 0.05)
            if x_max - x_min < 0.18:
                x_max = x_min + 0.18

            y_min = self._scan_center_cm - self._scan_half_range_cm
            y_max = self._scan_center_cm + self._scan_half_range_cm

            def to_plot(moment_emu: float, z_cm: float) -> QtCore.QPointF:
                x_ratio = (moment_emu - x_min) / max(x_max - x_min, 1e-6)
                y_ratio = (y_max - z_cm) / max(y_max - y_min, 1e-6)
                return QtCore.QPointF(
                    plot_inner.left() + x_ratio * plot_inner.width(),
                    plot_inner.top() + y_ratio * plot_inner.height(),
                )

            connector_anchor = QtCore.QPointF(body_rect.left() - 4.0, measurement_y)
            painter.setPen(QtGui.QPen(QtGui.QColor(180, 143, 103, 184), 1.05))
            painter.drawLine(plot_rect.topRight(), connector_anchor)
            painter.drawLine(plot_rect.bottomRight(), connector_anchor)

            painter.setPen(QtGui.QPen(QtGui.QColor("#d7c6b8"), 1.0))
            painter.setBrush(QtGui.QColor(255, 252, 248, 236))
            painter.drawRoundedRect(plot_rect, 16, 16)

            grid_pen = QtGui.QPen(QtGui.QColor(145, 124, 112, 52), 0.9)
            for fraction in (0.0, 0.25, 0.5, 0.75, 1.0):
                y_line = plot_inner.top() + fraction * plot_inner.height()
                painter.setPen(grid_pen)
                painter.drawLine(QtCore.QPointF(plot_inner.left(), y_line), QtCore.QPointF(plot_inner.right(), y_line))
            for fraction in (0.0, 0.5, 1.0):
                x_line = plot_inner.left() + fraction * plot_inner.width()
                painter.setPen(grid_pen)
                painter.drawLine(QtCore.QPointF(x_line, plot_inner.top()), QtCore.QPointF(x_line, plot_inner.bottom()))

            painter.setPen(QtGui.QPen(QtGui.QColor("#6f5a52"), 1.05))
            painter.drawLine(plot_inner.bottomLeft(), plot_inner.topLeft())
            painter.drawLine(plot_inner.bottomLeft(), plot_inner.bottomRight())

            axis_font = QtGui.QFont(self.font())
            axis_font.setPointSizeF(max(6.6, axis_font.pointSizeF() - 1.2))
            painter.setFont(axis_font)
            painter.setPen(QtGui.QColor("#6f5a52"))
            painter.drawText(
                QtCore.QRectF(plot_rect.left() + 4, plot_rect.top() + 4, plot_rect.width() - 8, 12),
                QtCore.Qt.AlignmentFlag.AlignHCenter | QtCore.Qt.AlignmentFlag.AlignVCenter,
                "Meas zone",
            )
            painter.drawText(
                QtCore.QRectF(plot_inner.left() - 36, plot_inner.top() - 4, 32, 12),
                QtCore.Qt.AlignmentFlag.AlignRight | QtCore.Qt.AlignmentFlag.AlignVCenter,
                f"{y_max:+.2f}",
            )
            painter.drawText(
                QtCore.QRectF(plot_inner.left() - 36, plot_inner.bottom() - 8, 32, 12),
                QtCore.Qt.AlignmentFlag.AlignRight | QtCore.Qt.AlignmentFlag.AlignVCenter,
                f"{y_min:+.2f}",
            )
            painter.drawText(
                QtCore.QRectF(plot_inner.left() - 6, plot_inner.bottom() + 6, 20, 12),
                QtCore.Qt.AlignmentFlag.AlignLeft | QtCore.Qt.AlignmentFlag.AlignVCenter,
                f"{x_min:.1f}",
            )
            painter.drawText(
                QtCore.QRectF(plot_inner.right() - 20, plot_inner.bottom() + 6, 28, 12),
                QtCore.Qt.AlignmentFlag.AlignRight | QtCore.Qt.AlignmentFlag.AlignVCenter,
                f"{x_max:.1f}",
            )
            painter.drawText(
                QtCore.QRectF(plot_inner.left(), plot_inner.bottom() + 18, plot_inner.width(), 12),
                QtCore.Qt.AlignmentFlag.AlignHCenter | QtCore.Qt.AlignmentFlag.AlignVCenter,
                "Moment",
            )
            painter.save()
            painter.translate(plot_rect.left() + 8, plot_inner.center().y())
            painter.rotate(-90)
            painter.drawText(
                QtCore.QRectF(-plot_inner.height() / 2.0, -10.0, plot_inner.height(), 12.0),
                QtCore.Qt.AlignmentFlag.AlignHCenter | QtCore.Qt.AlignmentFlag.AlignVCenter,
                "Z position",
            )
            painter.restore()

            if self._scan_points:
                ordered_points = sorted(self._scan_points, key=lambda point: point.z_cm)
                polyline = QtGui.QPainterPath()
                first_point = to_plot(ordered_points[0].moment_emu, ordered_points[0].z_cm)
                polyline.moveTo(first_point)
                for point in ordered_points[1:]:
                    polyline.lineTo(to_plot(point.moment_emu, point.z_cm))
                painter.setPen(QtGui.QPen(QtGui.QColor("#7a0219"), 1.9))
                painter.drawPath(polyline)
                painter.setBrush(QtGui.QColor("#ffca3a"))
                painter.setPen(QtGui.QPen(QtGui.QColor("#7a0219"), 0.9))
                for point in ordered_points:
                    painter.drawEllipse(to_plot(point.moment_emu, point.z_cm), 2.8, 2.8)
            else:
                painter.setPen(QtGui.QColor(111, 90, 82, 140))
                painter.drawText(plot_inner, QtCore.Qt.AlignmentFlag.AlignCenter, "No scan")

            if self._suggested_z_cm is not None and y_min <= self._suggested_z_cm <= y_max:
                suggested_y = to_plot(x_min, self._suggested_z_cm).y()
                painter.setPen(QtGui.QPen(QtGui.QColor("#31566d"), 1.15, QtCore.Qt.PenStyle.DashLine))
                painter.drawLine(
                    QtCore.QPointF(plot_inner.left(), suggested_y),
                    QtCore.QPointF(plot_inner.right(), suggested_y),
                )
                painter.setPen(QtGui.QColor("#31566d"))
                painter.drawText(
                    QtCore.QRectF(plot_inner.left() + 4, suggested_y - 14, plot_inner.width() - 8, 12),
                    QtCore.Qt.AlignmentFlag.AlignLeft | QtCore.Qt.AlignmentFlag.AlignVCenter,
                    "Opt",
                )

        label_height = 58.0
        gap = 12.0
        left_indicators = [indicator for indicator in model.indicators if indicator.side == "left"]
        right_entries: list[tuple[str, object, float]] = [
            ("indicator", indicator, float(indicator.raw_position)) for indicator in model.indicators if indicator.side != "left"
        ]
        right_entries.extend(("band", band, (band.raw_top + band.raw_bottom) / 2.0) for band in model.bands)
        right_entries.sort(key=lambda item: item[2], reverse=True)
        left_indicators.sort(key=lambda indicator: indicator.raw_position, reverse=True)

        left_targets = [to_y(indicator.raw_position) - label_height / 2.0 for indicator in left_indicators]
        right_targets = [to_y(position) - label_height / 2.0 for _, _, position in right_entries]
        left_y = self._adjust_label_positions(left_targets, panel.top() + 6, panel.bottom() - label_height - 6, label_height + gap)
        right_y = self._adjust_label_positions(right_targets, panel.top() + 6, panel.bottom() - label_height - 6, label_height + gap)

        title_font = QtGui.QFont(self.font())
        title_font.setPointSizeF(max(7.9, title_font.pointSizeF() - 0.5))
        title_font.setBold(True)
        meta_font = QtGui.QFont(title_font)
        meta_font.setBold(False)
        meta_font.setPointSizeF(max(7.0, meta_font.pointSizeF() - 0.7))

        def draw_indicator(indicator: ProfileIndicator, y_value: float, label_y: float) -> None:
            color = QtGui.QColor(indicator.color)
            if indicator.side == "left":
                anchor_x = center_x - indicator.bar_half_width
                tick_end = body_rect.left() - 18
                label_rect = QtCore.QRectF(panel.left() + 8, label_y, tick_end - panel.left() - 16, label_height)
                line_end = label_rect.right() + 4
                text_align = QtCore.Qt.AlignmentFlag.AlignRight | QtCore.Qt.AlignmentFlag.AlignVCenter
            else:
                anchor_x = center_x + indicator.bar_half_width
                tick_end = body_rect.right() + 18
                label_rect = QtCore.QRectF(tick_end + 14, label_y, panel.right() - tick_end - 22, label_height)
                line_end = label_rect.left() - 4
                text_align = QtCore.Qt.AlignmentFlag.AlignLeft | QtCore.Qt.AlignmentFlag.AlignVCenter

            painter.setPen(QtGui.QPen(color, 1.55 if indicator.emphasis else 1.15))
            if indicator.style == "bar":
                painter.drawLine(QtCore.QPointF(center_x - indicator.bar_half_width, y_value), QtCore.QPointF(center_x + indicator.bar_half_width, y_value))
            painter.drawLine(QtCore.QPointF(anchor_x, y_value), QtCore.QPointF(tick_end, y_value))
            painter.drawLine(QtCore.QPointF(tick_end, y_value), QtCore.QPointF(line_end, label_rect.center().y()))
            painter.setBrush(color)
            if indicator.symbol == "rect":
                painter.drawRoundedRect(QtCore.QRectF(center_x - 7.0, y_value - 4.5, 14.0, 9.0), 3.0, 3.0)
            else:
                dot_radius = 4.2 if indicator.emphasis else 3.0
                painter.drawEllipse(QtCore.QPointF(center_x, y_value), dot_radius, dot_radius)

            painter.setPen(QtGui.QPen(QtGui.QColor(255, 255, 255, 224), 1))
            painter.setBrush(QtGui.QColor(255, 255, 255, 232))
            painter.drawRoundedRect(label_rect, 10, 10)
            painter.setPen(QtGui.QColor("#402f2b"))
            painter.setFont(title_font)
            painter.drawText(
                label_rect.adjusted(8, 4, -8, -18),
                text_align | QtCore.Qt.TextFlag.TextWordWrap,
                indicator.label,
            )
            painter.setPen(QtGui.QColor("#7a625c"))
            painter.setFont(meta_font)
            painter.drawText(label_rect.adjusted(8, 34, -8, -4), text_align, indicator.value_text or f"{indicator.raw_position:,}")

        def draw_band_label(band: ProfileBand, label_y: float) -> None:
            center_y = to_y((band.raw_top + band.raw_bottom) / 2.0)
            band_edge_x = center_x + band.width / 2.0
            tick_end = body_rect.right() + 18
            label_rect = QtCore.QRectF(tick_end + 14, label_y, panel.right() - tick_end - 22, label_height)
            painter.setPen(QtGui.QPen(QtGui.QColor(band.outline_color), 1.1))
            painter.drawLine(QtCore.QPointF(band_edge_x, center_y), QtCore.QPointF(tick_end, center_y))
            painter.drawLine(QtCore.QPointF(tick_end, center_y), QtCore.QPointF(label_rect.left() - 4, label_rect.center().y()))
            painter.setPen(QtGui.QPen(QtGui.QColor(255, 255, 255, 224), 1))
            painter.setBrush(QtGui.QColor(255, 255, 255, 232))
            painter.drawRoundedRect(label_rect, 10, 10)
            painter.setPen(QtGui.QColor("#402f2b"))
            painter.setFont(title_font)
            painter.drawText(
                label_rect.adjusted(8, 4, -8, -18),
                QtCore.Qt.AlignmentFlag.AlignLeft | QtCore.Qt.AlignmentFlag.AlignVCenter | QtCore.Qt.TextFlag.TextWordWrap,
                band.label,
            )
            painter.setPen(QtGui.QColor("#7a625c"))
            painter.setFont(meta_font)
            painter.drawText(label_rect.adjusted(8, 34, -8, -4), QtCore.Qt.AlignmentFlag.AlignLeft | QtCore.Qt.AlignmentFlag.AlignVCenter, band.value_text or f"{band.raw_bottom:,}")

        for indicator, label_y in zip(left_indicators, left_y):
            draw_indicator(indicator, to_y(indicator.raw_position), label_y)

        for (entry_type, entry, position), label_y in zip(right_entries, right_y):
            if entry_type == "indicator":
                draw_indicator(entry, to_y(position), label_y)
            else:
                draw_band_label(entry, label_y)
        painter.end()


def read_calibration_from_settings(settings_path: Path) -> SquidCalibration:
    try:
        config = _load_settings_config(settings_path)
    except Exception:
        return SquidCalibration()

    section = config["MagnetometerCalibration"] if config.has_section("MagnetometerCalibration") else {}
    try:
        return SquidCalibration(
            xcal=float(section.get("XCal", "-3.410")),
            ycal=float(section.get("YCal", "-3.470")),
            zcal=float(section.get("ZCal", "-2.516")),
            range_fact=float(section.get("RangeFact", "0.00001")),
        )
    except ValueError:
        return SquidCalibration()


def read_calibration_from_ini(ini_path: Path) -> SquidCalibration:
    return read_calibration_from_settings(ini_path)


def _probe_squid_port(port: str) -> bool:
    try:
        with serial.Serial(
            port=port,
            baudrate=1200,
            bytesize=serial.EIGHTBITS,
            parity=serial.PARITY_NONE,
            stopbits=serial.STOPBITS_ONE,
            timeout=0.25,
            write_timeout=0.25,
        ) as ser:
            ser.reset_input_buffer()
            ser.reset_output_buffer()
            ser.write(b"\rALC\r")
            ser.flush()
            time.sleep(0.08)
            ser.write(b"\rALD\r")
            ser.flush()
            time.sleep(0.12)

            def query(command: str, timeout_s: float = 0.35) -> str:
                ser.reset_input_buffer()
                ser.write(f"\r{command}\r".encode("ascii", errors="ignore"))
                ser.flush()
                deadline = time.monotonic() + timeout_s
                chunks = bytearray()
                while time.monotonic() < deadline:
                    byte = ser.read(1)
                    if not byte:
                        continue
                    if byte == b"\r":
                        if chunks:
                            break
                        continue
                    chunks.extend(byte)
                return chunks.decode("ascii", errors="ignore").strip()

            count_response = query("XSC")
            data_response = query("XSD")
            return FLOAT_RE.search(count_response) is not None and FLOAT_RE.search(data_response) is not None
    except (OSError, serial.SerialException):
        return False


def _moment_magnitude(x_emu: float, y_emu: float, z_emu: float) -> float:
    return math.sqrt(x_emu * x_emu + y_emu * y_emu + z_emu * z_emu)


def fit_best_measurement_position(points: list[ScanPoint]) -> tuple[float | None, str]:
    if not points:
        return None, "no-data"
    ordered = sorted(points, key=lambda point: point.z_cm)
    peak_index = max(range(len(ordered)), key=lambda idx: ordered[idx].moment_emu)
    peak_point = ordered[peak_index]
    best_z = peak_point.z_cm
    method = "peak"

    if 0 < peak_index < len(ordered) - 1:
        left = ordered[peak_index - 1]
        mid = ordered[peak_index]
        right = ordered[peak_index + 1]
        x1, y1 = left.z_cm, left.moment_emu
        x2, y2 = mid.z_cm, mid.moment_emu
        x3, y3 = right.z_cm, right.moment_emu
        denom = (x1 - x2) * (x1 - x3) * (x2 - x3)
        if abs(denom) > 1e-12:
            a = (x3 * (y2 - y1) + x2 * (y1 - y3) + x1 * (y3 - y2)) / denom
            b = (x3 * x3 * (y1 - y2) + x2 * x2 * (y3 - y1) + x1 * x1 * (y2 - y3)) / denom
            if abs(a) > 1e-12:
                vertex = -b / (2.0 * a)
                if min(x1, x3) <= vertex <= max(x1, x3):
                    best_z = vertex
                    method = "quadratic"

    if method == "peak":
        threshold = peak_point.moment_emu * 0.95
        strong_points = [point for point in ordered if point.moment_emu >= threshold]
        weight_sum = sum(point.moment_emu for point in strong_points)
        if strong_points and weight_sum > 0:
            best_z = sum(point.z_cm * point.moment_emu for point in strong_points) / weight_sum
            method = "weighted"
    return best_z, method


class RawSquidClient:
    def __init__(self) -> None:
        self._serial: serial.Serial | None = None

    @property
    def is_connected(self) -> bool:
        return self._serial is not None and self._serial.is_open

    def connect(
        self,
        port: str,
        baudrate: int = 1200,
        bytesize: int = serial.EIGHTBITS,
        parity: str = serial.PARITY_NONE,
        stopbits: float = serial.STOPBITS_ONE,
        timeout: float = 1.0,
    ) -> None:
        self.disconnect()
        self._serial = serial.Serial(
            port=port,
            baudrate=baudrate,
            bytesize=bytesize,
            parity=parity,
            stopbits=stopbits,
            timeout=timeout,
            write_timeout=timeout,
        )
        self._serial.reset_input_buffer()
        self._serial.reset_output_buffer()

    def disconnect(self) -> None:
        if self._serial is not None:
            try:
                self._serial.close()
            finally:
                self._serial = None

    def _require_serial(self) -> serial.Serial:
        if self._serial is None or not self._serial.is_open:
            raise SquidCommunicationError("SQUID serial port is not connected.")
        return self._serial

    def _send(self, command: str) -> None:
        port = self._require_serial()
        payload = f"\r{command}\r".encode("ascii", errors="ignore")
        port.write(payload)
        port.flush()

    def _read_response(self, timeout_s: float = 1.0) -> str:
        port = self._require_serial()
        deadline = time.monotonic() + timeout_s
        chunks = bytearray()
        while time.monotonic() < deadline:
            byte = port.read(1)
            if not byte:
                continue
            if byte == b"\r":
                break
            chunks.extend(byte)
        if not chunks:
            raise SquidCommunicationError("Timed out waiting for SQUID response.")
        return chunks.decode("ascii", errors="ignore").strip()

    def _query_float(self, command: str) -> float:
        self._send(command)
        response = self._read_response()
        match = FLOAT_RE.search(response)
        if not match:
            raise SquidCommunicationError(f"No numeric value in SQUID response for {command!r}: {response!r}")
        return float(match.group(0))

    def read_xyz_raw(self) -> tuple[float, float, float]:
        self._send("ALC")
        time.sleep(0.10)
        self._send("ALD")
        time.sleep(0.12)
        x = -(self._query_float("XSD") + self._query_float("XSC"))
        y = -(self._query_float("YSD") + self._query_float("YSC"))
        z = -(self._query_float("ZSD") + self._query_float("ZSC"))
        return x, y, z


class SquidMomentReader:
    def __init__(self) -> None:
        self._client = RawSquidClient()

    @property
    def is_connected(self) -> bool:
        return self._client.is_connected

    def connect(self, port: str, baudrate: int = 1200) -> None:
        self._client.connect(port, baudrate=baudrate)

    def disconnect(self) -> None:
        self._client.disconnect()

    def take_baseline(self) -> tuple[float, float, float]:
        return self._client.read_xyz_raw()

    def read_moment(
        self,
        calibration: SquidCalibration,
        baseline_raw: tuple[float, float, float] | None,
    ) -> tuple[float, float, float, float]:
        x_raw, y_raw, z_raw = self._client.read_xyz_raw()
        bx, by, bz = baseline_raw if baseline_raw is not None else (0.0, 0.0, 0.0)
        x_emu = (x_raw - bx) * calibration.xcal * calibration.range_fact
        y_emu = (y_raw - by) * calibration.ycal * calibration.range_fact
        z_emu = (z_raw - bz) * calibration.zcal * calibration.range_fact
        return x_emu, y_emu, z_emu, _moment_magnitude(x_emu, y_emu, z_emu)


class UpDownController:
    def __init__(self, profile: SettingsProfile) -> None:
        self.profile = profile
        self.motor = MotorSerialClient(config=profile.motion_defaults)

    def apply_settings_profile(self, profile: SettingsProfile) -> None:
        self.profile = profile
        self.motor.config = profile.motion_defaults

    @property
    def is_connected(self) -> bool:
        return self.motor.is_connected

    def connect(self, port: str) -> None:
        self.motor.connect(port, baudrate=57600)

    def disconnect(self) -> None:
        self.motor.disconnect()

    def read_position(self) -> int:
        return self.motor.read_position(self.profile.updown_axis)

    def top_switch_active(self) -> bool:
        return self.motor.check_internal_status(self.profile.updown_axis, TOP_SWITCH_BIT) == 1

    def halt(self) -> None:
        self.motor.halt(self.profile.updown_axis)

    def home_to_top(self) -> MoveResult:
        return self.motor.home_to_top(self.profile.updown_axis)

    def move_to_raw(self, target: int, velocity: int) -> MoveResult:
        return self.motor.move_motor(
            self.profile.updown_axis,
            target=target,
            velocity=int(velocity),
            wait_for_stop=True,
            acceleration=96637,
            relative_mode=False,
        )

    def jog_relative(self, delta: int, velocity: int) -> MoveResult:
        current = self.read_position()
        result = self.motor.move_motor(
            self.profile.updown_axis,
            target=int(delta),
            velocity=int(velocity),
            wait_for_stop=True,
            acceleration=96637,
            relative_mode=True,
        )
        return MoveResult(target=current + int(delta), final_position=result.final_position, success=result.success)

    def sample_pickup(self) -> MoveResult:
        return self.motor.sample_pickup(self.profile.updown_axis)

    def sample_dropoff(self) -> MoveResult:
        return self.motor.sample_dropoff(self.profile.updown_axis)


class ScanWorker(QtCore.QThread):
    point_acquired = QtCore.Signal(object)
    scan_complete = QtCore.Signal(object)
    scan_failed = QtCore.Signal(str)
    log_message = QtCore.Signal(str)

    def __init__(
        self,
        controller: UpDownController,
        squid: SquidMomentReader,
        calibration: SquidCalibration,
        baseline_raw: tuple[float, float, float] | None,
        target_positions: list[int],
        settle_s: float,
        velocity_raw: int,
        sample_height_cm: float,
        counts_per_cm: int,
        target_meas_center_raw: int,
        safe_min_raw: int,
        safe_max_raw: int,
        parent: QtCore.QObject | None = None,
    ) -> None:
        super().__init__(parent)
        self._controller = controller
        self._squid = squid
        self._calibration = calibration
        self._baseline_raw = baseline_raw
        self._target_positions = target_positions
        self._settle_s = settle_s
        self._velocity_raw = velocity_raw
        self._sample_height_cm = sample_height_cm
        self._counts_per_cm = counts_per_cm
        self._target_meas_center_raw = target_meas_center_raw
        self._safe_min_raw = min(safe_min_raw, safe_max_raw)
        self._safe_max_raw = max(safe_min_raw, safe_max_raw)
        self._stop_requested = False

    def request_stop(self) -> None:
        self._stop_requested = True
        try:
            self._controller.halt()
        except Exception:
            pass

    def _check_abort(self) -> None:
        if self._stop_requested:
            raise RuntimeError("Scan cancelled.")

    def _ensure_bounds(self, target: int) -> None:
        if not (self._safe_min_raw <= target <= self._safe_max_raw):
            raise RuntimeError(
                f"Requested scan target {target:,} is outside the enforced safe range "
                f"[{self._safe_min_raw:,}, {self._safe_max_raw:,}]."
            )

    def run(self) -> None:
        try:
            points: list[ScanPoint] = []
            tolerance = max(POSITION_TOLERANCE_COUNTS, abs(self._counts_per_cm) // 10, 100)
            for index, target in enumerate(self._target_positions, start=1):
                self._check_abort()
                self._ensure_bounds(target)
                self.log_message.emit(f"Scan move {index}/{len(self._target_positions)} to {target:,} raw counts.")
                result = self._controller.move_to_raw(target, self._velocity_raw)
                final_position = result.final_position
                if abs(final_position - target) > tolerance:
                    try:
                        self._controller.halt()
                    except Exception:
                        pass
                    raise RuntimeError(
                        f"Z move did not settle at the requested target. "
                        f"Requested {target:,}, reached {final_position:,}."
                    )
                if self._controller.top_switch_active() and abs(final_position) > tolerance:
                    try:
                        self._controller.halt()
                    except Exception:
                        pass
                    raise RuntimeError(
                        "Top switch became active away from the expected top region. "
                        "The scan stopped to protect the holder."
                    )
                self._check_abort()
                time.sleep(self._settle_s)
                x_emu, y_emu, z_emu, moment_emu = self._squid.read_moment(self._calibration, self._baseline_raw)
                point = ScanPoint(
                    index=index,
                    raw_position=final_position,
                    z_cm=final_position / self._counts_per_cm,
                    x_emu=x_emu,
                    y_emu=y_emu,
                    z_emu=z_emu,
                    moment_emu=moment_emu,
                )
                points.append(point)
                self.point_acquired.emit(point)
                self.log_message.emit(
                    f"Scan point {index}: raw={final_position:,}, z={point.z_cm:+.3f} cm, moment={moment_emu:.3e} emu"
                )
            suggested_z_cm, fit_method = fit_best_measurement_position(points)
            suggested_target_raw = None
            suggested_meas_pos_raw = None
            note = ""
            if suggested_z_cm is not None:
                suggested_target_raw = int(round(suggested_z_cm * self._counts_per_cm))
                half_height_counts = int(round(self._sample_height_cm * self._counts_per_cm / 2.0))
                suggested_meas_pos_raw = suggested_target_raw - half_height_counts
                if not (self._safe_min_raw <= suggested_target_raw <= self._safe_max_raw):
                    note = (
                        "Best-fit target falls outside the enforced safety range; "
                        "review the scan window before applying it."
                    )
            self.scan_complete.emit(
                ScanResult(
                    points=points,
                    suggested_z_cm=suggested_z_cm,
                    suggested_target_raw=suggested_target_raw,
                    suggested_meas_pos_raw=suggested_meas_pos_raw,
                    fit_method=fit_method,
                    note=note,
                )
            )
        except Exception as exc:
            self.scan_failed.emit(str(exc))


class MainWindow(QtWidgets.QMainWindow):
    def __init__(self) -> None:
        super().__init__()
        self.setWindowTitle("RapidPy Up/Down Control")
        self._settings = load_settings()
        self.settings_profile = self._load_profile(Path(self._settings.settings_path))
        self.controller = UpDownController(self.settings_profile)
        self.squid = SquidMomentReader()
        self._calibration = read_calibration_from_settings(self.settings_profile.path)
        self._baseline_raw: tuple[float, float, float] | None = None
        self._scan_worker: ScanWorker | None = None
        self._scan_points: list[ScanPoint] = []
        self._plot_suggested_z_cm: float | None = None
        self._plot_suggested_target_raw: int | None = None
        self._poll_timer = QtCore.QTimer(self)
        self._poll_timer.setInterval(500)
        self._poll_timer.timeout.connect(self._poll_live_state)
        self._suppress_speed_confirm = False
        self._last_confirmed_speed = 0
        self._last_live_raw: int | None = None
        self._pending_meas_pos_suggestion: int | None = None
        self._build_ui()
        self._apply_local_style()
        self.setMinimumSize(1560, 820)
        self.resize(1640, 920)
        self._refresh_ports()
        self._populate_squid_ports()
        self._load_settings_into_widgets()
        self._autodetect_squid_port()
        self._apply_profile_to_ui(reset_motion=True)
        self._poll_timer.start()

    def _load_profile(self, path: Path) -> SettingsProfile:
        try:
            return _load_settings_profile(path)
        except Exception:
            fallback = DEFAULT_SETTINGS_PATH if DEFAULT_SETTINGS_PATH.exists() else path
            if fallback.exists():
                return _load_settings_profile(fallback)
            return SettingsProfile(
                path=fallback,
                motion_defaults=MotorControllerConfig(),
                updown_axis=MotorAxisConfig(name="UpDown", motor_id=3, address=16),
                updown_motor_1cm=10000,
                zero_pos=-50_000,
                meas_pos=30000,
                af_pos=-42_500,
                irm_pos=-36_000,
                scoil_pos=-22_700,
                floor_pos=-148_955,
                sample_bottom=425000,
                sample_top=582500,
                sample_height_counts=157500,
            )

    def _build_ui(self) -> None:
        root = QtWidgets.QWidget(self)
        self.setCentralWidget(root)
        outer = QtWidgets.QVBoxLayout(root)
        outer.setContentsMargins(12, 12, 12, 12)
        outer.setSpacing(10)

        header = QtWidgets.QFrame()
        header.setObjectName("card")
        header_layout = QtWidgets.QHBoxLayout(header)
        header_layout.setContentsMargins(18, 16, 18, 16)
        header_layout.setSpacing(12)
        title_col = QtWidgets.QVBoxLayout()
        title = QtWidgets.QLabel("Up/Down Axis")
        title.setObjectName("title")
        subtitle = QtWidgets.QLabel(
            "Compact Z-axis control with VB6 raw-count tuning, top-switch status, and Z scanning."
        )
        subtitle.setObjectName("subtitle")
        subtitle.setWordWrap(True)
        title_col.addWidget(title)
        title_col.addWidget(subtitle)
        header_layout.addLayout(title_col, 1)

        self.current_position_pill = self._make_pill("Z -- raw")
        self.top_switch_pill = self._make_pill("Z TOP --")
        self.meas_pos_pill = self._make_pill("MeasPos --")
        self.live_raw_label = self._make_pill("Raw --")
        self.live_cm_label = self._make_pill("Z -- cm")
        self.live_target_label = self._make_pill("Target --")
        self.safety_label = QtWidgets.QLabel()
        self.safety_label.setObjectName("tableHint")
        self.safety_label.setWordWrap(True)
        header_status = QtWidgets.QVBoxLayout()
        header_status.setObjectName("headerStatusHost")
        header_status.setSpacing(8)
        pill_grid = QtWidgets.QGridLayout()
        pill_grid.setContentsMargins(0, 0, 0, 0)
        pill_grid.setHorizontalSpacing(10)
        pill_grid.setVerticalSpacing(8)
        pill_grid.addWidget(self.current_position_pill, 0, 0)
        pill_grid.addWidget(self.top_switch_pill, 0, 1)
        pill_grid.addWidget(self.meas_pos_pill, 0, 2)
        pill_grid.addWidget(self.live_raw_label, 1, 0)
        pill_grid.addWidget(self.live_cm_label, 1, 1)
        pill_grid.addWidget(self.live_target_label, 1, 2)
        pill_grid.setColumnStretch(0, 1)
        pill_grid.setColumnStretch(1, 1)
        pill_grid.setColumnStretch(2, 1)
        header_status.addLayout(pill_grid)
        header_status.addWidget(self.safety_label)
        header_layout.addLayout(header_status, 2)
        outer.addWidget(header)
        apply_card_shadow(header)

        shell = QtWidgets.QHBoxLayout()
        shell.setSpacing(10)
        outer.addLayout(shell, 1)

        left_host = QtWidgets.QWidget()
        left_host.setObjectName("columnHost")
        left_host.setMinimumWidth(296)
        left_host.setMaximumWidth(328)
        left_host_layout = QtWidgets.QVBoxLayout(left_host)
        left_host_layout.setContentsMargins(0, 0, 0, 0)
        left_host_layout.setSpacing(10)
        shell.addWidget(left_host)

        center_host = QtWidgets.QWidget()
        center_host.setObjectName("columnHost")
        center_host.setMinimumWidth(660)
        center_layout = QtWidgets.QVBoxLayout(center_host)
        center_layout.setContentsMargins(0, 0, 0, 0)
        center_layout.setSpacing(10)
        shell.addWidget(center_host, 2)

        motion_host = QtWidgets.QWidget()
        motion_host.setObjectName("columnHost")
        motion_host.setMinimumWidth(270)
        motion_host.setMaximumWidth(296)
        motion_layout = QtWidgets.QVBoxLayout(motion_host)
        motion_layout.setContentsMargins(0, 0, 0, 0)
        motion_layout.setSpacing(10)
        shell.addWidget(motion_host)

        scan_host = QtWidgets.QWidget()
        scan_host.setObjectName("columnHost")
        scan_host.setMinimumWidth(276)
        scan_host.setMaximumWidth(302)
        scan_layout = QtWidgets.QVBoxLayout(scan_host)
        scan_layout.setContentsMargins(0, 0, 0, 0)
        scan_layout.setSpacing(10)
        shell.addWidget(scan_host)

        self._build_connections_card(left_host_layout)
        self._build_settings_card(left_host_layout)
        left_host_layout.addStretch(1)

        self._build_profile_card(center_layout)

        self._build_motion_card(motion_layout)
        self._build_console_card(motion_layout)
        motion_layout.addStretch(1)

        self._build_scan_card(scan_layout)
        scan_layout.addStretch(1)

    def _build_connections_card(self, parent: QtWidgets.QVBoxLayout) -> None:
        card, layout = self._build_card(
            "Connections",
            "Motor and SQUID ports with live top-switch context.",
        )
        self.motor_port_combo = QtWidgets.QComboBox()
        self.motor_port_combo.setEditable(True)
        self.squid_port_combo = QtWidgets.QComboBox()
        self.squid_port_combo.setEditable(True)
        self.refresh_ports_btn = QtWidgets.QPushButton("Refresh")
        self.connect_motor_btn = QtWidgets.QPushButton("Connect")
        self.connect_motor_btn.setObjectName("accent")
        self.disconnect_motor_btn = QtWidgets.QPushButton("Disconnect")
        self.connect_squid_btn = QtWidgets.QPushButton("Connect")
        self.disconnect_squid_btn = QtWidgets.QPushButton("Disconnect")

        motor_row = QtWidgets.QHBoxLayout()
        motor_row.setContentsMargins(0, 0, 0, 0)
        motor_row.setSpacing(4)
        motor_row.addWidget(self.motor_port_combo, 1)
        motor_row.addWidget(self.refresh_ports_btn)
        layout.addRow("Motor port", self._layout_widget(motor_row))
        motor_btn_row = QtWidgets.QHBoxLayout()
        motor_btn_row.setContentsMargins(0, 0, 0, 0)
        motor_btn_row.setSpacing(4)
        motor_btn_row.addWidget(self.connect_motor_btn)
        motor_btn_row.addWidget(self.disconnect_motor_btn)
        layout.addRow("", self._layout_widget(motor_btn_row))

        squid_row = QtWidgets.QHBoxLayout()
        squid_row.setContentsMargins(0, 0, 0, 0)
        squid_row.setSpacing(4)
        squid_row.addWidget(self.squid_port_combo, 1)
        layout.addRow("SQUID port", self._layout_widget(squid_row))
        squid_btn_row = QtWidgets.QHBoxLayout()
        squid_btn_row.setContentsMargins(0, 0, 0, 0)
        squid_btn_row.setSpacing(4)
        squid_btn_row.addWidget(self.connect_squid_btn)
        squid_btn_row.addWidget(self.disconnect_squid_btn)
        layout.addRow("", self._layout_widget(squid_btn_row))
        self.connections_status = QtWidgets.QLabel()
        self.connections_status.setObjectName("tableHint")
        self.connections_status.setWordWrap(True)
        layout.addRow("", self.connections_status)
        parent.addWidget(card)

        self.refresh_ports_btn.clicked.connect(self._refresh_ports)
        self.connect_motor_btn.clicked.connect(self._connect_motor)
        self.disconnect_motor_btn.clicked.connect(self._disconnect_motor)
        self.connect_squid_btn.clicked.connect(self._connect_squid)
        self.disconnect_squid_btn.clicked.connect(self._disconnect_squid)

    def _build_settings_card(self, parent: QtWidgets.QVBoxLayout) -> None:
        card, layout = self._build_card(
            "Settings And References",
            "Settings file, MeasPos, and VB6 reference heights for the center profile.",
        )
        self.settings_path_edit = QtWidgets.QLineEdit(str(self.settings_profile.path))
        self.settings_browse_btn = QtWidgets.QPushButton("Browse")
        self.settings_reload_btn = QtWidgets.QPushButton("Reload")
        self.save_settings_btn = QtWidgets.QPushButton("Save File")
        self.save_settings_btn.setObjectName("accent")

        self.meas_pos_spin = QtWidgets.QSpinBox()
        self.meas_pos_spin.setRange(-2_000_000_000, 2_000_000_000)
        self.apply_suggestion_btn = QtWidgets.QPushButton("Use Suggestion")
        self.apply_suggestion_btn.setEnabled(False)
        self.assumed_target_label = QtWidgets.QLabel()
        self.assumed_target_label.setObjectName("tableHint")
        self.assumed_target_label.setWordWrap(True)
        self.reference_positions_label = QtWidgets.QLabel()
        self.reference_positions_label.setObjectName("tableHint")
        self.reference_positions_label.setWordWrap(True)

        layout.addRow("Settings file", self.settings_path_edit)
        settings_btn_row = QtWidgets.QHBoxLayout()
        settings_btn_row.setContentsMargins(0, 0, 0, 0)
        settings_btn_row.setSpacing(4)
        settings_btn_row.addWidget(self.settings_browse_btn)
        settings_btn_row.addWidget(self.settings_reload_btn)
        layout.addRow("", self._layout_widget(settings_btn_row))
        layout.addRow("MeasPos", self.meas_pos_spin)
        layout.addRow("", self.assumed_target_label)
        layout.addRow("", self.reference_positions_label)
        save_row = QtWidgets.QHBoxLayout()
        save_row.setContentsMargins(0, 0, 0, 0)
        save_row.setSpacing(4)
        save_row.addWidget(self.apply_suggestion_btn)
        save_row.addWidget(self.save_settings_btn)
        layout.addRow("", self._layout_widget(save_row))
        parent.addWidget(card)

        self.settings_browse_btn.clicked.connect(self._browse_settings_path)
        self.settings_reload_btn.clicked.connect(self._reload_settings_profile)
        self.save_settings_btn.clicked.connect(self._save_settings_file)
        self.apply_suggestion_btn.clicked.connect(self._apply_scan_suggestion)
        self.meas_pos_spin.valueChanged.connect(self._update_assumed_target_text)

    def _build_profile_card(self, parent: QtWidgets.QVBoxLayout) -> None:
        card = QtWidgets.QFrame()
        card.setObjectName("card")
        apply_card_shadow(card)
        layout = QtWidgets.QVBoxLayout(card)
        layout.setContentsMargins(14, 14, 14, 14)
        layout.setSpacing(8)
        title = QtWidgets.QLabel("Magnetometer Z Profile")
        title.setObjectName("consoleTitle")
        subtitle = QtWidgets.QLabel(
            "Holder-bottom referenced bore cartoon using VB6 Z references."
        )
        subtitle.setObjectName("subtitle")
        subtitle.setWordWrap(True)
        self.profile_scene = VerticalProfileWidget()
        self.profile_caption = QtWidgets.QLabel()
        self.profile_caption.setObjectName("tableHint")
        self.profile_caption.setWordWrap(True)

        layout.addWidget(title)
        layout.addWidget(subtitle)
        layout.addWidget(self.profile_scene, 1)
        layout.addWidget(self.profile_caption)
        parent.addWidget(card, 1)

    def _build_motion_card(self, parent: QtWidgets.QVBoxLayout) -> None:
        card, layout = self._build_card(
            "Velocity And Jog",
            "Raw-count motion with live cm estimates.",
        )
        self.z_velocity_spin = QtWidgets.QSpinBox()
        self.z_velocity_spin.setRange(1, 50_000_000)
        self.z_velocity_spin.setSingleStep(100_000)
        self.z_velocity_spin.setGroupSeparatorShown(True)
        self.z_velocity_estimate = QtWidgets.QLabel()
        self.z_velocity_estimate.setObjectName("tableHint")
        self.z_velocity_estimate.setWordWrap(True)

        self.jog_step_spin = QtWidgets.QSpinBox()
        self.jog_step_spin.setRange(1, 2_000_000_000)
        self.jog_step_spin.setSingleStep(100)
        self.jog_step_spin.setGroupSeparatorShown(True)
        self.jog_step_estimate = QtWidgets.QLabel()
        self.jog_step_estimate.setObjectName("tableHint")
        self.jog_step_estimate.setWordWrap(True)

        self.target_raw_spin = QtWidgets.QSpinBox()
        self.target_raw_spin.setRange(-2_000_000_000, 2_000_000_000)
        self.target_raw_spin.setGroupSeparatorShown(True)
        self.target_raw_hint = QtWidgets.QLabel()
        self.target_raw_hint.setObjectName("tableHint")
        self.target_raw_hint.setWordWrap(True)

        self.move_target_btn = QtWidgets.QPushButton("Go Target")
        self.move_target_btn.setObjectName("accent")
        self.move_meas_btn = QtWidgets.QPushButton("Go Meas Z")
        self.home_top_btn = QtWidgets.QPushButton("Home Top")
        self.pickup_btn = QtWidgets.QPushButton("Pickup")
        self.dropoff_btn = QtWidgets.QPushButton("Dropoff")
        self.susceptibility_btn = QtWidgets.QPushButton("Susc. Meter")

        self.pickup_raw_spin = QtWidgets.QSpinBox()
        self.pickup_raw_spin.setRange(-2_000_000_000, 2_000_000_000)
        self.pickup_raw_spin.setGroupSeparatorShown(True)
        self.dropoff_raw_spin = QtWidgets.QSpinBox()
        self.dropoff_raw_spin.setRange(-2_000_000_000, 2_000_000_000)
        self.dropoff_raw_spin.setGroupSeparatorShown(True)
        self.susceptibility_raw_spin = QtWidgets.QSpinBox()
        self.susceptibility_raw_spin.setRange(-2_000_000_000, 2_000_000_000)
        self.susceptibility_raw_spin.setGroupSeparatorShown(True)

        self.jog_up_btn = QtWidgets.QPushButton("Jog Up")
        self.jog_down_btn = QtWidgets.QPushButton("Jog Down")

        grid = QtWidgets.QGridLayout()
        grid.setHorizontalSpacing(6)
        grid.setVerticalSpacing(6)
        grid.addWidget(QtWidgets.QLabel("Z velocity"), 0, 0)
        grid.addWidget(self.z_velocity_spin, 0, 1)
        grid.addWidget(self.z_velocity_estimate, 1, 0, 1, 2)
        grid.addWidget(QtWidgets.QLabel("Jog step"), 2, 0)
        grid.addWidget(self.jog_step_spin, 2, 1)
        grid.addWidget(self.jog_step_estimate, 3, 0, 1, 2)
        grid.addWidget(QtWidgets.QLabel("Raw target"), 4, 0)
        grid.addWidget(self.target_raw_spin, 4, 1)
        grid.addWidget(self.target_raw_hint, 5, 0, 1, 2)
        grid.addWidget(self.move_target_btn, 6, 0)
        grid.addWidget(self.move_meas_btn, 6, 1)
        grid.addWidget(self.jog_up_btn, 7, 0)
        grid.addWidget(self.jog_down_btn, 7, 1)
        grid.addWidget(self.home_top_btn, 8, 0, 1, 2)
        grid.addWidget(self.pickup_btn, 9, 0)
        grid.addWidget(self.pickup_raw_spin, 9, 1)
        grid.addWidget(self.dropoff_btn, 10, 0)
        grid.addWidget(self.dropoff_raw_spin, 10, 1)
        grid.addWidget(self.susceptibility_btn, 11, 0)
        grid.addWidget(self.susceptibility_raw_spin, 11, 1)
        layout.addRow(self._layout_widget(grid))
        parent.addWidget(card)

        self.z_velocity_spin.valueChanged.connect(self._on_velocity_changed)
        self.jog_step_spin.valueChanged.connect(self._update_motion_hints)
        self.target_raw_spin.valueChanged.connect(self._update_motion_hints)
        self.move_target_btn.clicked.connect(self._move_to_target_raw)
        self.move_meas_btn.clicked.connect(self._move_to_assumed_measurement)
        self.home_top_btn.clicked.connect(self._home_to_top)
        self.pickup_btn.clicked.connect(lambda: self._move_to_preset("pickup", self.pickup_raw_spin.value(), use_pickup=True))
        self.dropoff_btn.clicked.connect(lambda: self._move_to_preset("dropoff", self.dropoff_raw_spin.value(), use_dropoff=True))
        self.susceptibility_btn.clicked.connect(
            lambda: self._move_to_preset("susceptibility", self.susceptibility_raw_spin.value())
        )
        self.jog_up_btn.clicked.connect(lambda: self._jog_relative(upward=True))
        self.jog_down_btn.clicked.connect(lambda: self._jog_relative(upward=False))

    def _build_scan_card(self, parent: QtWidgets.QVBoxLayout) -> None:
        card, layout = self._build_card(
            "Measurement Z Optimization",
            "Scan around the assumed measurement target and fit the best Z position.",
        )
        self.sample_height_spin = QtWidgets.QDoubleSpinBox()
        self.sample_height_spin.setRange(0.1, 20.0)
        self.sample_height_spin.setDecimals(3)
        self.sample_height_spin.setSingleStep(0.1)
        self.sample_height_spin.setSuffix(" cm")

        self.scan_half_range_spin = QtWidgets.QDoubleSpinBox()
        self.scan_half_range_spin.setRange(0.1, 10.0)
        self.scan_half_range_spin.setDecimals(2)
        self.scan_half_range_spin.setSingleStep(0.1)
        self.scan_half_range_spin.setSuffix(" cm")

        self.scan_step_spin = QtWidgets.QDoubleSpinBox()
        self.scan_step_spin.setRange(0.01, 1.0)
        self.scan_step_spin.setDecimals(3)
        self.scan_step_spin.setSingleStep(0.05)
        self.scan_step_spin.setSuffix(" cm")

        self.scan_settle_spin = QtWidgets.QDoubleSpinBox()
        self.scan_settle_spin.setRange(0.1, 10.0)
        self.scan_settle_spin.setDecimals(2)
        self.scan_settle_spin.setSingleStep(0.1)
        self.scan_settle_spin.setSuffix(" s")

        self.scan_window_hint = QtWidgets.QLabel()
        self.scan_window_hint.setObjectName("tableHint")
        self.scan_window_hint.setWordWrap(True)

        self.take_baseline_btn = QtWidgets.QPushButton("Take Baseline")
        self.take_baseline_btn.setObjectName("accent")
        self.baseline_label = QtWidgets.QLabel("Baseline not captured")
        self.baseline_label.setWordWrap(True)
        self.scan_start_btn = QtWidgets.QPushButton("Start Scan")
        self.scan_start_btn.setObjectName("accent")
        self.scan_stop_btn = QtWidgets.QPushButton("Stop")
        self.scan_stop_btn.setEnabled(False)
        self.scan_result_label = QtWidgets.QLabel("No scan result yet.")
        self.scan_result_label.setObjectName("tableHint")
        self.scan_result_label.setWordWrap(True)

        form = QtWidgets.QFormLayout()
        form.addRow("Sample ht", self.sample_height_spin)
        form.addRow("Half-range", self.scan_half_range_spin)
        form.addRow("Step", self.scan_step_spin)
        form.addRow("Settle", self.scan_settle_spin)
        form.addRow("", self.scan_window_hint)
        form.addRow("", self.take_baseline_btn)
        form.addRow("", self.baseline_label)
        scan_btn_row = QtWidgets.QHBoxLayout()
        scan_btn_row.addWidget(self.scan_start_btn)
        scan_btn_row.addWidget(self.scan_stop_btn)
        form.addRow("", self._layout_widget(scan_btn_row))
        form.addRow("", self.scan_result_label)
        layout.addRow(self._layout_widget(form))
        parent.addWidget(card)

        self.sample_height_spin.valueChanged.connect(self._update_assumed_target_text)
        self.sample_height_spin.valueChanged.connect(self._update_scan_window_hint)
        self.scan_half_range_spin.valueChanged.connect(self._update_scan_window_hint)
        self.scan_step_spin.valueChanged.connect(self._update_scan_window_hint)
        self.take_baseline_btn.clicked.connect(self._take_squid_baseline)
        self.scan_start_btn.clicked.connect(self._start_scan)
        self.scan_stop_btn.clicked.connect(self._stop_scan)

    def _build_live_card(self, parent: QtWidgets.QVBoxLayout) -> None:
        card = QtWidgets.QFrame()
        card.setObjectName("card")
        apply_card_shadow(card)
        layout = QtWidgets.QVBoxLayout(card)
        layout.setContentsMargins(14, 14, 14, 14)
        layout.setSpacing(8)

        title = QtWidgets.QLabel("Live State")
        title.setObjectName("consoleTitle")
        layout.addWidget(title)

        values = QtWidgets.QHBoxLayout()
        values.setSpacing(6)
        self.live_raw_label = self._make_pill("Raw --")
        self.live_cm_label = self._make_pill("Z -- cm")
        self.live_target_label = self._make_pill("Target --")
        values.addWidget(self.live_raw_label)
        values.addWidget(self.live_cm_label)
        values.addWidget(self.live_target_label)
        layout.addLayout(values)

        self.safety_label = QtWidgets.QLabel()
        self.safety_label.setObjectName("tableHint")
        self.safety_label.setWordWrap(True)
        layout.addWidget(self.safety_label)
        parent.addWidget(card)

    def _build_plot_card(self, parent: QtWidgets.QVBoxLayout) -> None:
        card = QtWidgets.QFrame()
        card.setObjectName("card")
        apply_card_shadow(card)
        layout = QtWidgets.QVBoxLayout(card)
        layout.setContentsMargins(14, 14, 14, 14)
        layout.setSpacing(8)
        title = QtWidgets.QLabel("Z Scan Plot")
        title.setObjectName("consoleTitle")
        layout.addWidget(title)
        if pg is None:
            self.scan_plot = None
            placeholder = QtWidgets.QLabel("pyqtgraph is not installed, so the scan plot is unavailable.")
            placeholder.setWordWrap(True)
            layout.addWidget(placeholder)
        else:
            self.scan_plot = pg.PlotWidget()
            self.scan_plot.setBackground((0, 0, 0, 0))
            self.scan_plot.setLabel("left", "Moment (emu)")
            self.scan_plot.setLabel("bottom", "Z position (cm)")
            self.scan_plot.showGrid(x=True, y=True, alpha=0.18)
            self.scan_curve = self.scan_plot.plot([], [], pen=pg.mkPen("#7A0219", width=3), symbol="o", symbolBrush="#FFCA3A", symbolSize=8)
            self.suggestion_line = pg.InfiniteLine(angle=90, movable=False, pen=pg.mkPen("#31566D", width=2, style=QtCore.Qt.PenStyle.DashLine))
            self.scan_plot.addItem(self.suggestion_line)
            self.suggestion_line.hide()
            self.suggestion_label = pg.TextItem(anchor=(0, 1), color="#31566D", fill=pg.mkBrush(255, 255, 255, 228))
            self.scan_plot.addItem(self.suggestion_label)
            self.suggestion_label.hide()
            layout.addWidget(self.scan_plot, 1)
        parent.addWidget(card, 2)

    def _build_console_card(self, parent: QtWidgets.QVBoxLayout) -> None:
        card = QtWidgets.QFrame()
        card.setObjectName("card")
        apply_card_shadow(card)
        layout = QtWidgets.QVBoxLayout(card)
        layout.setContentsMargins(14, 14, 14, 14)
        layout.setSpacing(8)
        title = QtWidgets.QLabel("Console")
        title.setObjectName("consoleTitle")
        layout.addWidget(title)
        self.console = QtWidgets.QPlainTextEdit()
        self.console.setObjectName("console")
        self.console.setReadOnly(True)
        self.console.setMinimumHeight(92)
        self.console.setMaximumHeight(132)
        layout.addWidget(self.console, 1)
        parent.addWidget(card)

    def _build_card(self, title_text: str, subtitle_text: str) -> tuple[QtWidgets.QFrame, QtWidgets.QFormLayout]:
        card = QtWidgets.QFrame()
        card.setObjectName("card")
        apply_card_shadow(card)
        outer = QtWidgets.QVBoxLayout(card)
        outer.setContentsMargins(12, 12, 12, 12)
        outer.setSpacing(6)
        title = QtWidgets.QLabel(title_text)
        title.setObjectName("consoleTitle")
        subtitle = QtWidgets.QLabel(subtitle_text)
        subtitle.setObjectName("subtitle")
        subtitle.setWordWrap(True)
        outer.addWidget(title)
        outer.addWidget(subtitle)
        form = QtWidgets.QFormLayout()
        form.setFieldGrowthPolicy(QtWidgets.QFormLayout.FieldGrowthPolicy.ExpandingFieldsGrow)
        form.setLabelAlignment(QtCore.Qt.AlignmentFlag.AlignLeft)
        form.setFormAlignment(QtCore.Qt.AlignmentFlag.AlignTop)
        form.setContentsMargins(0, 4, 0, 0)
        form.setSpacing(4)
        outer.addLayout(form)
        return card, form

    def _make_pill(self, text: str) -> QtWidgets.QLabel:
        label = QtWidgets.QLabel(text)
        label.setObjectName("valuePill")
        label.setAlignment(QtCore.Qt.AlignmentFlag.AlignCenter)
        label.setSizePolicy(QtWidgets.QSizePolicy.Policy.Expanding, QtWidgets.QSizePolicy.Policy.Preferred)
        label.setMinimumHeight(38)
        return label

    def _layout_widget(self, layout: QtWidgets.QLayout) -> QtWidgets.QWidget:
        widget = QtWidgets.QWidget()
        widget.setObjectName("layoutWrapper")
        widget.setLayout(layout)
        return widget

    def _apply_local_style(self) -> None:
        compact_font = QtGui.QFont(self.font())
        compact_size = compact_font.pointSizeF()
        if compact_size > 0:
            compact_font.setPointSizeF(max(8.2, compact_size - 0.9))
            self.setFont(compact_font)
        self.setStyleSheet(
            self.styleSheet()
            + """
            QWidget#columnHost, QWidget#layoutWrapper {
                background: transparent;
            }
            QScrollArea#panelScroll { background: transparent; border: none; }
            QScrollArea#panelScroll > QWidget > QWidget { background: transparent; }
            QLabel#subtitle {
                color: #6d5a55;
                font-size: 10px;
            }
            QLabel#tableHint {
                color: #5e4b47;
                background: rgba(255, 252, 248, 0.78);
                border: 1px solid rgba(122, 2, 25, 0.10);
                border-radius: 12px;
                padding: 6px 8px;
                font-size: 10px;
            }
            QLabel#consoleTitle {
                color: #5d0013;
                font-size: 14px;
                font-weight: 700;
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
            QPushButton {
                min-height: 26px;
                padding: 4px 9px;
                font-size: 11px;
            }
            QPushButton#accent {
                min-height: 26px;
                padding: 4px 10px;
                font-size: 11px;
            }
            QSpinBox, QDoubleSpinBox, QLineEdit, QComboBox {
                min-height: 24px;
                font-size: 10px;
            }
            QLabel#valuePill {
                background: rgba(255, 255, 255, 0.90);
                border: 1px solid rgba(122, 2, 25, 0.14);
                border-radius: 18px;
                font-size: 11px;
                font-weight: 650;
                padding: 8px 12px;
                min-width: 116px;
            }
            QFrame#card {
                margin: 0px;
            }
            """
        )

    def _append(self, message: str) -> None:
        self.console.appendPlainText(message)

    def _available_ports(self) -> list[str]:
        return sorted(port.device for port in list_ports.comports())

    def _populate_port_combo(self, combo: QtWidgets.QComboBox, ports: list[str], wanted: str = "") -> None:
        current = combo.currentText().strip()
        combo.clear()
        combo.addItem("")
        combo.addItems(ports)
        target = wanted.strip() or current
        if target:
            if combo.findText(target) < 0:
                combo.addItem(target)
            combo.setCurrentText(target)
        else:
            combo.setCurrentIndex(0)

    def _refresh_ports(self) -> None:
        if not hasattr(self, "motor_port_combo"):
            return
        self._populate_port_combo(self.motor_port_combo, self._available_ports(), self._settings.motor_port)

    def _populate_squid_ports(self) -> None:
        if not hasattr(self, "squid_port_combo"):
            return
        self._populate_port_combo(self.squid_port_combo, self._available_ports(), self._settings.squid_port)

    def _autodetect_squid_port(self) -> None:
        if not hasattr(self, "squid_port_combo"):
            return
        ports = self._available_ports()
        existing = self.squid_port_combo.currentText().strip() or self._settings.squid_port.strip()
        if existing and existing in ports and _probe_squid_port(existing):
            self._populate_port_combo(self.squid_port_combo, ports, existing)
            return

        detected = next((port for port in ports if _probe_squid_port(port)), "")
        self._populate_port_combo(self.squid_port_combo, ports, detected)
        if detected:
            self._settings.squid_port = detected

    def _load_settings_into_widgets(self) -> None:
        self.motor_port_combo.setCurrentText(self._settings.motor_port)
        self.squid_port_combo.setCurrentText(self._settings.squid_port)
        self.settings_path_edit.setText(self._settings.settings_path)
        self.meas_pos_spin.setValue(self.settings_profile.meas_pos)
        self.pickup_raw_spin.setValue(self._settings.pickup_raw)
        self.dropoff_raw_spin.setValue(self._settings.dropoff_raw)
        self.susceptibility_raw_spin.setValue(self._settings.susceptibility_meter_raw)
        self.z_velocity_spin.setValue(self._settings.z_velocity_raw)
        self.jog_step_spin.setValue(self._settings.jog_step_raw)
        self.target_raw_spin.setValue(self._settings.target_raw)
        self.sample_height_spin.setValue(self._settings.sample_height_cm)
        self.scan_half_range_spin.setValue(self._settings.scan_half_range_cm)
        self.scan_step_spin.setValue(self._settings.scan_step_cm)
        self.scan_settle_spin.setValue(self._settings.scan_settle_s)

    def _apply_profile_to_ui(self, reset_motion: bool = False) -> None:
        if reset_motion:
            self.pickup_raw_spin.setValue(self.settings_profile.sample_bottom)
            self.dropoff_raw_spin.setValue(self.settings_profile.sample_bottom + int(self.settings_profile.sample_height_counts * 0.9))
            self.susceptibility_raw_spin.setValue(self._settings.susceptibility_meter_raw)
            self.z_velocity_spin.setValue(self._default_z_speed_raw())
            self.jog_step_spin.setValue(self._safe_jog_step())
            self.target_raw_spin.setValue(self.settings_profile.meas_pos)
        self.meas_pos_spin.setValue(self.settings_profile.meas_pos)
        self.meas_pos_pill.setText(f"MeasPos {self.settings_profile.meas_pos:,}")
        self._calibration = read_calibration_from_settings(self.settings_profile.path)
        self._pending_meas_pos_suggestion = None
        self.apply_suggestion_btn.setEnabled(False)
        self._update_motion_hints()
        self._update_assumed_target_text()
        self._update_scan_window_hint()
        self._update_safety_label(None)
        self._update_reference_positions_text()
        self._update_connections_status()
        self._refresh_profile_model(None)

    def _counts_per_cm(self) -> int:
        return self.settings_profile.updown_motor_1cm

    def _raw_to_cm(self, raw: int) -> float | None:
        counts_per_cm = self._counts_per_cm()
        if counts_per_cm == 0:
            return None
        return raw / counts_per_cm

    def _cm_to_raw(self, cm_value: float) -> int | None:
        counts_per_cm = self._counts_per_cm()
        if counts_per_cm == 0:
            return None
        return int(round(cm_value * counts_per_cm))

    def _measurement_sign(self) -> int:
        counts_per_cm = self._counts_per_cm()
        if counts_per_cm > 0:
            return 1
        if counts_per_cm < 0:
            return -1
        return 1 if self.settings_profile.meas_pos >= 0 else -1

    def _safe_raw_bounds(self) -> tuple[int, int]:
        sign = self._measurement_sign()
        vb6_bottom_limit = sign * int(round(abs(self.settings_profile.meas_pos) * 1.15))
        logical_min = min(0, vb6_bottom_limit)
        logical_max = max(0, vb6_bottom_limit)
        return logical_min, logical_max

    def _settings_snapshot_dir(self, target: Path) -> Path:
        return target.parent / ".rapidpy_history" / target.stem

    def _create_settings_snapshot(self, target: Path) -> Path | None:
        if not target.exists():
            return None
        history_dir = self._settings_snapshot_dir(target)
        history_dir.mkdir(parents=True, exist_ok=True)
        timestamp = time.strftime("%Y-%m-%d_%H-%M-%S")
        snapshot = history_dir / f"{timestamp}{target.suffix}"
        counter = 1
        while snapshot.exists():
            snapshot = history_dir / f"{timestamp}_{counter}{target.suffix}"
            counter += 1
        shutil.copy2(target, snapshot)
        return snapshot

    def _browse_settings_path(self) -> None:
        path, _ = QtWidgets.QFileDialog.getOpenFileName(
            self,
            "Select Settings File",
            self.settings_path_edit.text().strip() or str(DEFAULT_SETTINGS_PATH),
            "Settings Files (*.ini *.json);;INI Files (*.ini);;JSON Files (*.json);;All Files (*)",
        )
        if path:
            self.settings_path_edit.setText(path)

    def _build_current_settings_config(self, source_path: Path) -> configparser.ConfigParser:
        config = _load_settings_config(source_path) if source_path.exists() else _new_settings_config()
        if not config.has_section("SteppingMotor"):
            config.add_section("SteppingMotor")
        config["SteppingMotor"]["MeasPos"] = str(self.meas_pos_spin.value())
        return config

    def _reload_settings_profile(self) -> None:
        path = Path(self.settings_path_edit.text().strip())
        if not path.exists():
            QtWidgets.QMessageBox.warning(self, "Missing Settings", f"Could not find settings file:\n{path}")
            return
        try:
            profile = _load_settings_profile(path)
        except Exception as exc:
            QtWidgets.QMessageBox.warning(self, "Load Failed", str(exc))
            return
        self.settings_profile = profile
        self.controller.apply_settings_profile(profile)
        self._settings.settings_path = str(path)
        self._append(f"Loaded settings from {path}")
        self._apply_profile_to_ui(reset_motion=True)

    def _save_settings_file(self) -> None:
        path = Path(self.settings_path_edit.text().strip())
        if not path.suffix:
            path = path.with_suffix(".ini")
            self.settings_path_edit.setText(str(path))
        source_path = path if path.exists() else self.settings_profile.path
        try:
            config = self._build_current_settings_config(source_path)
        except Exception as exc:
            QtWidgets.QMessageBox.warning(self, "Save Failed", str(exc))
            return
        snapshot = self._create_settings_snapshot(path)
        try:
            if path.suffix.lower() == ".json":
                path.write_text(json.dumps(_settings_json_payload_from_config(config), indent=2), encoding="utf-8")
            else:
                with path.open("w", encoding="utf-8", newline="\n") as handle:
                    config.write(handle)
        except OSError as exc:
            QtWidgets.QMessageBox.warning(self, "Save Failed", str(exc))
            return
        self.settings_profile = _load_settings_profile(path)
        self.controller.apply_settings_profile(self.settings_profile)
        self._settings.settings_path = str(path)
        self._append(f"Saved MeasPos {self.meas_pos_spin.value():,} to {path}")
        if snapshot is not None:
            self._append(f"Snapshot saved to {snapshot}")
        self._apply_profile_to_ui(reset_motion=False)

    def _connect_motor(self) -> None:
        port = self.motor_port_combo.currentText().strip()
        if not port:
            QtWidgets.QMessageBox.warning(self, "Missing Port", "Select the motor COM port first.")
            return
        try:
            self.controller.connect(port)
        except Exception as exc:
            QtWidgets.QMessageBox.critical(self, "Motor Connection Error", str(exc))
            return
        self._settings.motor_port = port
        self._append(f"Connected motor on {port}")
        self._update_connections_status()
        self._poll_live_state()

    def _disconnect_motor(self) -> None:
        self.controller.disconnect()
        self._append("Disconnected motor")
        self._update_connections_status()
        self._poll_live_state()

    def _connect_squid(self) -> None:
        port = self.squid_port_combo.currentText().strip()
        if not port:
            QtWidgets.QMessageBox.warning(self, "Missing Port", "Select the SQUID COM port first.")
            return
        try:
            self.squid.connect(port, baudrate=int(self._settings.squid_baud))
        except Exception as exc:
            QtWidgets.QMessageBox.critical(self, "SQUID Connection Error", str(exc))
            return
        self._settings.squid_port = port
        self._append(f"Connected SQUID on {port}")
        self._update_connections_status()

    def _disconnect_squid(self) -> None:
        self.squid.disconnect()
        self._append("Disconnected SQUID")
        self._update_connections_status()

    def _velocity_raw_scale(self) -> float:
        full_rotation = abs(float(self.settings_profile.motion_defaults.turning_motor_full_rotation))
        one_rps = abs(float(self.settings_profile.motion_defaults.turning_motor_1rps))
        if full_rotation > 0 and one_rps > 0:
            return one_rps / full_rotation
        return 2000.0

    def _raw_velocity_to_position_counts_per_second(self, raw_velocity: int | float) -> float:
        return float(raw_velocity) / self._velocity_raw_scale()

    def _position_counts_per_second_to_raw_velocity(self, position_counts_per_second: float) -> int:
        return max(1, int(round(position_counts_per_second * self._velocity_raw_scale())))

    def _default_z_speed_raw(self) -> int:
        counts_per_cm = abs(self._counts_per_cm())
        if counts_per_cm <= 0:
            return self.settings_profile.motion_defaults.lift_speed_slow
        raw_velocity = self._position_counts_per_second_to_raw_velocity(counts_per_cm * DEFAULT_MANUAL_SPEED_CM_PER_SEC)
        return max(1, min(raw_velocity, int(self.settings_profile.motion_defaults.lift_speed_fast)))

    def _safe_jog_step(self) -> int:
        counts_per_cm = abs(self._counts_per_cm())
        if counts_per_cm <= 0:
            return 4000
        return max(250, min(50_000, int(round(counts_per_cm / 10.0))))

    def _estimated_speed_cm_per_second(self, raw_velocity: int) -> float | None:
        counts_per_cm = abs(self._counts_per_cm())
        if counts_per_cm <= 0:
            return None
        return self._raw_velocity_to_position_counts_per_second(raw_velocity) / counts_per_cm

    def _format_z_velocity_estimate(self, raw_velocity: int) -> str:
        speed_cm_s = self._estimated_speed_cm_per_second(raw_velocity)
        if speed_cm_s is None:
            return "Estimated Z speed is unavailable until UpDownMotor1cm is loaded from the INI."
        return f"Estimated Z actual speed: ~{speed_cm_s:.2f} cm/s from {raw_velocity:,} raw counts/s."

    def _format_jog_step_estimate(self, step_counts: int) -> str:
        cm_value = self._raw_to_cm(step_counts)
        if cm_value is None:
            return f"Jog step: {step_counts:,} raw counts. Centimeter translation appears after UpDownMotor1cm loads."
        return f"Jog step: {step_counts:,} raw counts (~{abs(cm_value):.3f} cm)."

    def _format_target_hint(self, raw_target: int) -> str:
        cm_value = self._raw_to_cm(raw_target)
        low, high = self._safe_raw_bounds()
        if cm_value is None:
            return f"Target {raw_target:,} raw counts. Safe range [{low:,}, {high:,}]."
        return f"Target {raw_target:,} raw counts (~{cm_value:+.3f} cm). Safe range [{low:,}, {high:,}]."

    def _sample_top_for_length_cm(self, length_cm: float) -> int | None:
        sample_length_raw = self._cm_to_raw(length_cm)
        if sample_length_raw is None:
            return None
        return self.meas_pos_spin.value() + sample_length_raw

    def _update_reference_positions_text(self) -> None:
        self.reference_positions_label.setText(
            " | ".join(
                (
                    f"Zero {self.settings_profile.zero_pos:,}",
                    f"Meas {self.meas_pos_spin.value():,}",
                    f"AF {self.settings_profile.af_pos:,}",
                    f"IRM {self.settings_profile.irm_pos:,}",
                    f"S coil {self.settings_profile.scoil_pos:,}",
                    f"Floor {self.settings_profile.floor_pos:,}",
                )
            )
        )

    def _update_connections_status(self) -> None:
        motor_status = "connected" if self.controller.is_connected else "disconnected"
        squid_status = "connected" if self.squid.is_connected else "disconnected"
        self.connections_status.setText(
            f"Motor {motor_status}; SQUID {squid_status}. Top switch reads status bit 4."
        )

    def _profile_model(self, current_raw: int | None) -> ProfileModel:
        counts_per_cm = abs(self._counts_per_cm()) or 961
        susc_height = max(int(round(counts_per_cm * 0.80)), 4200)
        af_height = max(int(round(counts_per_cm * 1.15)), 5600)
        meas_height = max(self.settings_profile.sample_height_counts, int(round(counts_per_cm * 1.35)), 7600)
        measurement_cutoff_margin = max(int(round(counts_per_cm * 0.55)), 900)
        bands = (
            ProfileBand(
                "Susc.\nmeter",
                self.settings_profile.scoil_pos + susc_height,
                self.settings_profile.scoil_pos,
                18.0,
                "#6c9f7a",
                "#2d6a4f",
                value_text=f"{self.settings_profile.scoil_pos:,}",
            ),
            ProfileBand(
                "AF\ncoils",
                self.settings_profile.af_pos + af_height,
                self.settings_profile.af_pos,
                18.0,
                "#c99864",
                "#8c5b34",
                value_text=f"{self.settings_profile.af_pos:,}",
            ),
            ProfileBand(
                "Measurement\nzone",
                self.meas_pos_spin.value() + meas_height,
                self.meas_pos_spin.value(),
                18.0,
                "#d68d72",
                "#7a2c26",
                value_text=f"{self.meas_pos_spin.value():,}",
            ),
        )
        indicators = [
            ProfileIndicator("Top switch\nmotor 0", 0, "left", "#7a0219", "0", bar_half_width=18.0, emphasis=True),
            ProfileIndicator(
                "1.0 in\ntop",
                self.settings_profile.sample_top,
                "right",
                "#8f4b45",
                f"{self.settings_profile.sample_top:,}",
                symbol="rect",
                bar_half_width=18.0,
                emphasis=True,
            ),
            ProfileIndicator(
                "XY stage\nholder bottom",
                self.settings_profile.sample_bottom,
                "left",
                "#31566d",
                f"{self.settings_profile.sample_bottom:,}",
                bar_half_width=18.0,
            ),
            ProfileIndicator(
                "Zero\nbaseline",
                self.settings_profile.zero_pos,
                "right",
                "#8a6a44",
                f"{self.settings_profile.zero_pos:,}",
                bar_half_width=18.0,
            ),
        ]
        if current_raw is not None:
            indicators.append(
                ProfileIndicator("Live Z", current_raw, "right", "#31566d", f"{current_raw:,}", bar_half_width=18.0, emphasis=True)
            )
        range_bottom = self.meas_pos_spin.value() - measurement_cutoff_margin
        return ProfileModel(
            range_top=0,
            range_bottom=range_bottom,
            top_switch_raw=0,
            holder_bottom_raw=self.settings_profile.sample_bottom,
            sample_top_raw=self.settings_profile.sample_top,
            sample_bottom_raw=self.settings_profile.sample_bottom,
            zero_raw=self.settings_profile.zero_pos,
            floor_raw=self.settings_profile.floor_pos,
            live_raw=current_raw,
            bands=bands,
            indicators=tuple(indicators),
        )

    def _refresh_profile_model(self, current_raw: int | None) -> None:
        self._last_live_raw = current_raw
        model = self._profile_model(current_raw)
        self.profile_scene.set_profile(model)
        center_cm = self._raw_to_cm(self.meas_pos_spin.value())
        self.profile_scene.set_scan_detail(
            self._scan_points,
            0.0 if center_cm is None else center_cm,
            float(self.scan_half_range_spin.value()),
            self._plot_suggested_z_cm,
            self._plot_suggested_target_raw,
        )
        counts_per_cm = self._counts_per_cm()
        caption = (
            f"Reference is holder bottom / XY stage at {self.settings_profile.sample_bottom:,} raw. "
            f"Top switch is motor 0 above the bore; stage bands mark S coil {self.settings_profile.scoil_pos:,}, "
            f"AF {self.settings_profile.af_pos:,}, and measurement zone {self.meas_pos_spin.value():,}. The profile is cropped just below the measurement region."
        )
        if counts_per_cm:
            caption += f" UpDownMotor1cm = {counts_per_cm:,} counts/cm."
        self.profile_caption.setText(caption)

    def _confirm_speed_if_needed(self, raw_velocity: int) -> bool:
        if self._suppress_speed_confirm:
            self._last_confirmed_speed = raw_velocity
            return True
        estimate = self._estimated_speed_cm_per_second(raw_velocity)
        if estimate is None or estimate < HIGH_SPEED_CONFIRM_CM_PER_SEC:
            self._last_confirmed_speed = raw_velocity
            return True
        previous = self._last_confirmed_speed or self._default_z_speed_raw()
        response = QtWidgets.QMessageBox.question(
            self,
            "High Speed Confirmation",
            f"{raw_velocity:,} counts/s is about {estimate:.2f} cm/s, which is unusually fast for manual Z motion. Continue?",
            QtWidgets.QMessageBox.StandardButton.Yes | QtWidgets.QMessageBox.StandardButton.No,
            QtWidgets.QMessageBox.StandardButton.No,
        )
        if response == QtWidgets.QMessageBox.StandardButton.Yes:
            self._last_confirmed_speed = raw_velocity
            return True
        self._suppress_speed_confirm = True
        self.z_velocity_spin.setValue(previous)
        self._suppress_speed_confirm = False
        return False

    def _update_motion_hints(self) -> None:
        self.z_velocity_estimate.setText(self._format_z_velocity_estimate(int(self.z_velocity_spin.value())))
        self.jog_step_estimate.setText(self._format_jog_step_estimate(int(self.jog_step_spin.value())))
        self.target_raw_hint.setText(self._format_target_hint(int(self.target_raw_spin.value())))
        self.live_target_label.setText(f"Target {self.target_raw_spin.value():,}")
        self._settings.z_velocity_raw = int(self.z_velocity_spin.value())
        self._settings.jog_step_raw = int(self.jog_step_spin.value())
        self._settings.target_raw = int(self.target_raw_spin.value())

    def _on_velocity_changed(self, value: int) -> None:
        if self._confirm_speed_if_needed(int(value)):
            self._update_motion_hints()

    def _assumed_measurement_target_raw(self) -> int | None:
        counts_per_cm = self._counts_per_cm()
        if counts_per_cm == 0:
            return None
        return self.meas_pos_spin.value() + int(round(self.sample_height_spin.value() * counts_per_cm / 2.0))

    def _update_assumed_target_text(self) -> None:
        target = self._assumed_measurement_target_raw()
        if target is None:
            self.assumed_target_label.setText(
                "Assumed measurement target is unavailable until UpDownMotor1cm is loaded from the settings INI."
            )
            self._refresh_profile_model(self._last_live_raw)
            return
        target_cm = self._raw_to_cm(target)
        self.assumed_target_label.setText(
            f"Assumed holder target for MeasPos {self.meas_pos_spin.value():,}: {target:,} raw"
            + ("" if target_cm is None else f" (~{target_cm:+.3f} cm).")
        )
        self.meas_pos_pill.setText(f"MeasPos {self.meas_pos_spin.value():,}")
        self._update_reference_positions_text()
        self._refresh_profile_model(self._last_live_raw)

    def _update_scan_window_hint(self) -> None:
        center = self._assumed_measurement_target_raw()
        counts_per_cm = self._counts_per_cm()
        if center is None or counts_per_cm == 0:
            self.scan_window_hint.setText(
                "Scan window preview becomes available once UpDownMotor1cm is loaded from the INI."
            )
            return
        half_range_counts = int(round(self.scan_half_range_spin.value() * counts_per_cm))
        step_counts = max(1, int(round(self.scan_step_spin.value() * counts_per_cm)))
        start = center - half_range_counts
        end = center + half_range_counts
        low, high = self._safe_raw_bounds()
        self.scan_window_hint.setText(
            f"Assumed center {center:,}. Scan {start:,} to {end:,} in steps of {abs(step_counts):,} raw. "
            f"Safe range [{low:,}, {high:,}]."
        )
        self._settings.sample_height_cm = float(self.sample_height_spin.value())
        self._settings.scan_half_range_cm = float(self.scan_half_range_spin.value())
        self._settings.scan_step_cm = float(self.scan_step_spin.value())
        self._settings.scan_settle_s = float(self.scan_settle_spin.value())

    def _update_safety_label(self, current_raw: int | None) -> None:
        low, high = self._safe_raw_bounds()
        message = f"Safe envelope: top switch / zero at 0, lower scan clip at {max(abs(low), abs(high)):,} raw (1.15 x |MeasPos|)."
        if current_raw is not None:
            message += f" Live Z {current_raw:,}."
        self.safety_label.setText(message)

    def _poll_live_state(self) -> None:
        if not self.controller.is_connected:
            self.current_position_pill.setText("Z disconnected")
            self.top_switch_pill.setText("Z TOP --")
            self.live_raw_label.setText("Raw --")
            self.live_cm_label.setText("Z -- cm")
            self._update_safety_label(None)
            self._update_connections_status()
            self._refresh_profile_model(None)
            return
        try:
            current_raw = self.controller.read_position()
            top_switch = self.controller.top_switch_active()
        except Exception as exc:
            self.current_position_pill.setText("Z read error")
            self.top_switch_pill.setText("Z TOP ?")
            self._append(f"Live poll error: {exc}")
            return
        current_cm = self._raw_to_cm(current_raw)
        self._last_live_raw = current_raw
        self.current_position_pill.setText(f"Z {current_raw:,}")
        self.top_switch_pill.setText("Z TOP ON" if top_switch else "Z TOP OFF")
        self.live_raw_label.setText(f"Raw {current_raw:,}")
        self.live_cm_label.setText("Z -- cm" if current_cm is None else f"Z {current_cm:+.3f} cm")
        self._update_safety_label(current_raw)
        self._update_connections_status()
        self._refresh_profile_model(current_raw)

    def _require_motor(self) -> bool:
        if self.controller.is_connected:
            return True
        QtWidgets.QMessageBox.warning(self, "Motor Not Connected", "Connect the up/down motor first.")
        return False

    def _move_checked(self, target: int, action: str, *, allow_outside_safe_bounds: bool = False) -> MoveResult | None:
        if not self._require_motor():
            return None
        velocity = int(self.z_velocity_spin.value())
        if not self._confirm_speed_if_needed(velocity):
            return None
        if not allow_outside_safe_bounds:
            low, high = self._safe_raw_bounds()
            if not (low <= target <= high):
                QtWidgets.QMessageBox.warning(
                    self,
                    "Target Outside Safe Bounds",
                    f"{action} target {target:,} is outside the enforced safe range [{low:,}, {high:,}].",
                )
                return None
        try:
            result = self.controller.move_to_raw(target, velocity)
        except Exception as exc:
            QtWidgets.QMessageBox.warning(self, f"{action} Error", str(exc))
            return None
        self._append(f"{action}: target={target:,}, final={result.final_position:,}, success={result.success}")
        self._poll_live_state()
        return result

    def _move_to_target_raw(self) -> None:
        self._move_checked(int(self.target_raw_spin.value()), "Move To Raw Target")

    def _move_to_assumed_measurement(self) -> None:
        target = self._assumed_measurement_target_raw()
        if target is None:
            QtWidgets.QMessageBox.warning(self, "Missing UpDownMotor1cm", "Load a settings INI with UpDownMotor1cm first.")
            return
        self._move_checked(target, "Move To Assumed Meas Z")

    def _home_to_top(self) -> None:
        if not self._require_motor():
            return
        try:
            result = self.controller.home_to_top()
        except Exception as exc:
            QtWidgets.QMessageBox.warning(self, "Home To Top Error", str(exc))
            return
        self._append(f"Home To Top: final={result.final_position:,}, success={result.success}")
        self._poll_live_state()

    def _move_to_preset(self, name: str, target: int, use_pickup: bool = False, use_dropoff: bool = False) -> None:
        if not self._require_motor():
            return
        try:
            if use_pickup:
                result = self.controller.sample_pickup()
            elif use_dropoff:
                result = self.controller.sample_dropoff()
            else:
                velocity = int(self.z_velocity_spin.value())
                if not self._confirm_speed_if_needed(velocity):
                    return
                result = self.controller.move_to_raw(target, velocity)
        except Exception as exc:
            QtWidgets.QMessageBox.warning(self, f"{name.title()} Error", str(exc))
            return
        self._append(f"{name.title()}: target={target:,}, final={result.final_position:,}, success={result.success}")
        self._poll_live_state()

    def _jog_relative(self, upward: bool) -> None:
        if not self._require_motor():
            return
        velocity = int(self.z_velocity_spin.value())
        if not self._confirm_speed_if_needed(velocity):
            return
        step = abs(int(self.jog_step_spin.value()))
        sign = self._measurement_sign()
        delta = -step * sign if upward else step * sign
        current = self.controller.read_position()
        target = current + delta
        low, high = self._safe_raw_bounds()
        if not (low <= target <= high):
            QtWidgets.QMessageBox.warning(
                self,
                "Jog Outside Safe Bounds",
                f"Requested jog would move to {target:,}, outside [{low:,}, {high:,}].",
            )
            return
        try:
            result = self.controller.jog_relative(delta, velocity)
        except Exception as exc:
            QtWidgets.QMessageBox.warning(self, "Jog Error", str(exc))
            return
        direction = "Jog Up" if upward else "Jog Down"
        self._append(f"{direction}: delta={delta:,}, final={result.final_position:,}, success={result.success}")
        self._poll_live_state()

    def _take_squid_baseline(self) -> None:
        if not self.squid.is_connected:
            QtWidgets.QMessageBox.warning(self, "SQUID Not Connected", "Connect the SQUID port first.")
            return
        self.take_baseline_btn.setEnabled(False)
        QtWidgets.QApplication.processEvents()
        try:
            self._baseline_raw = self.squid.take_baseline()
        except Exception as exc:
            self.take_baseline_btn.setEnabled(True)
            QtWidgets.QMessageBox.warning(self, "Baseline Error", str(exc))
            return
        self.take_baseline_btn.setEnabled(True)
        x_raw, y_raw, z_raw = self._baseline_raw
        x_emu = x_raw * self._calibration.xcal * self._calibration.range_fact
        y_emu = y_raw * self._calibration.ycal * self._calibration.range_fact
        z_emu = z_raw * self._calibration.zcal * self._calibration.range_fact
        self.baseline_label.setText(
            f"Baseline captured: X {x_emu:+.3e}, Y {y_emu:+.3e}, Z {z_emu:+.3e} emu"
        )
        self._append("Captured SQUID baseline for measurement-position scan.")

    def _scan_targets(self) -> tuple[list[int], int, int] | None:
        center = self._assumed_measurement_target_raw()
        counts_per_cm = self._counts_per_cm()
        if center is None or counts_per_cm == 0:
            QtWidgets.QMessageBox.warning(self, "Missing UpDownMotor1cm", "Load a settings INI with UpDownMotor1cm first.")
            return None
        half_range_counts = int(round(self.scan_half_range_spin.value() * counts_per_cm))
        step_counts = int(round(self.scan_step_spin.value() * counts_per_cm))
        if step_counts == 0:
            QtWidgets.QMessageBox.warning(self, "Invalid Step", "Scan step is too small after conversion to raw counts.")
            return None
        low, high = self._safe_raw_bounds()
        requested_start = center - half_range_counts
        requested_stop = center + half_range_counts
        lower = max(min(requested_start, requested_stop), low)
        upper = min(max(requested_start, requested_stop), high)
        if lower > upper:
            QtWidgets.QMessageBox.warning(self, "Unsafe Scan Window", "The requested scan window falls completely outside the enforced safe range.")
            return None
        direction = 1 if upper >= lower else -1
        step_abs = abs(step_counts)
        values: list[int] = []
        current = lower
        while current <= upper:
            values.append(current)
            current += step_abs
        if values[-1] != upper:
            values.append(upper)
        return values, low, high

    def _start_scan(self) -> None:
        if self._scan_worker is not None:
            return
        if not self._require_motor():
            return
        if not self.squid.is_connected:
            QtWidgets.QMessageBox.warning(self, "SQUID Not Connected", "Connect the SQUID port before starting the scan.")
            return
        if self._baseline_raw is None:
            QtWidgets.QMessageBox.warning(self, "Baseline Required", "Take a SQUID baseline before starting the scan.")
            return
        target_info = self._scan_targets()
        if target_info is None:
            return
        targets, safe_min_raw, safe_max_raw = target_info
        assumed_target = self._assumed_measurement_target_raw()
        if assumed_target is None:
            return
        self._scan_points = []
        self._refresh_plot()
        self.scan_start_btn.setEnabled(False)
        self.scan_stop_btn.setEnabled(True)
        self.apply_suggestion_btn.setEnabled(False)
        self.scan_result_label.setText("Running scan…")
        worker = ScanWorker(
            controller=self.controller,
            squid=self.squid,
            calibration=self._calibration,
            baseline_raw=self._baseline_raw,
            target_positions=targets,
            settle_s=float(self.scan_settle_spin.value()),
            velocity_raw=int(self.z_velocity_spin.value()),
            sample_height_cm=float(self.sample_height_spin.value()),
            counts_per_cm=self._counts_per_cm(),
            target_meas_center_raw=assumed_target,
            safe_min_raw=safe_min_raw,
            safe_max_raw=safe_max_raw,
            parent=self,
        )
        worker.point_acquired.connect(self._handle_scan_point)
        worker.scan_complete.connect(self._handle_scan_complete)
        worker.scan_failed.connect(self._handle_scan_failed)
        worker.log_message.connect(self._append)
        worker.finished.connect(self._scan_thread_finished)
        self._scan_worker = worker
        self._append(
            f"Starting measurement Z scan with {len(targets)} points around assumed target {assumed_target:,}."
        )
        worker.start()

    def _stop_scan(self) -> None:
        if self._scan_worker is None:
            return
        self._scan_worker.request_stop()
        self._append("Requested scan stop.")

    @QtCore.Slot(object)
    def _handle_scan_point(self, point: ScanPoint) -> None:
        self._scan_points.append(point)
        self._refresh_plot()

    @QtCore.Slot(object)
    def _handle_scan_complete(self, result: ScanResult) -> None:
        self._scan_points = list(result.points)
        self._refresh_plot(result.suggested_z_cm, result.suggested_target_raw)
        if result.suggested_meas_pos_raw is not None:
            self._pending_meas_pos_suggestion = result.suggested_meas_pos_raw
            self.apply_suggestion_btn.setEnabled(True)
        else:
            self._pending_meas_pos_suggestion = None
            self.apply_suggestion_btn.setEnabled(False)
        summary = "No suggestion available."
        if result.suggested_z_cm is not None and result.suggested_meas_pos_raw is not None and result.suggested_target_raw is not None:
            summary = (
                f"Best fit ({result.fit_method}) gives optimal Z {result.suggested_z_cm:+.3f} cm "
                f"at holder target {result.suggested_target_raw:,} raw, mapping to MeasPos {result.suggested_meas_pos_raw:,}."
            )
        if result.note:
            summary = summary + " " + result.note
        self.scan_result_label.setText(summary)
        self._append(summary)

    @QtCore.Slot(str)
    def _handle_scan_failed(self, message: str) -> None:
        self.scan_result_label.setText(message)
        self._append(f"Scan failed: {message}")
        QtWidgets.QMessageBox.warning(self, "Scan Failed", message)

    def _scan_thread_finished(self) -> None:
        self.scan_start_btn.setEnabled(True)
        self.scan_stop_btn.setEnabled(False)
        self._scan_worker = None
        self._poll_live_state()

    def _refresh_plot(self, suggested_z_cm: float | None = None, suggested_target_raw: int | None = None) -> None:
        self._plot_suggested_z_cm = suggested_z_cm
        self._plot_suggested_target_raw = suggested_target_raw
        center_cm = self._raw_to_cm(self.meas_pos_spin.value())
        self.profile_scene.set_scan_detail(
            self._scan_points,
            0.0 if center_cm is None else center_cm,
            float(self.scan_half_range_spin.value()),
            suggested_z_cm,
            suggested_target_raw,
        )

    def _apply_scan_suggestion(self) -> None:
        suggested = getattr(self, "_pending_meas_pos_suggestion", None)
        if suggested is None:
            return
        self.meas_pos_spin.setValue(int(suggested))
        self._append(f"Accepted suggested MeasPos {suggested:,}. Save the settings file when ready.")
        self._update_assumed_target_text()
        self.apply_suggestion_btn.setEnabled(False)

    def closeEvent(self, event: QtGui.QCloseEvent) -> None:  # noqa: N802
        if self._scan_worker is not None:
            self._scan_worker.request_stop()
            self._scan_worker.wait(2000)
        try:
            self._settings = UpDownSettings(
                motor_port=self.motor_port_combo.currentText().strip(),
                squid_port=self.squid_port_combo.currentText().strip(),
                squid_baud=self._settings.squid_baud,
                settings_path=self.settings_path_edit.text().strip(),
                min_raw_count=self._safe_raw_bounds()[0],
                max_raw_count=self._safe_raw_bounds()[1],
                pickup_raw=int(self.pickup_raw_spin.value()),
                dropoff_raw=int(self.dropoff_raw_spin.value()),
                susceptibility_meter_raw=int(self.susceptibility_raw_spin.value()),
                z_velocity_raw=int(self.z_velocity_spin.value()),
                jog_step_raw=int(self.jog_step_spin.value()),
                target_raw=int(self.target_raw_spin.value()),
                sample_height_cm=float(self.sample_height_spin.value()),
                scan_half_range_cm=float(self.scan_half_range_spin.value()),
                scan_step_cm=float(self.scan_step_spin.value()),
                scan_settle_s=float(self.scan_settle_spin.value()),
            )
            save_settings(self._settings)
        except Exception:
            pass
        try:
            self.controller.disconnect()
        except Exception:
            pass
        try:
            self.squid.disconnect()
        except Exception:
            pass
        super().closeEvent(event)


def main() -> int:
    app = QtWidgets.QApplication(sys.argv)
    apply_liquid_glass_theme(app)
    assets_dir = Path(__file__).resolve().parent.parent / "assets"
    set_app_icon(app, "updown_control_icon.png", assets_dir)
    window = MainWindow()
    window.show()
    return app.exec()
