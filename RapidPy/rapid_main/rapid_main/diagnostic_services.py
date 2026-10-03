"""Adapter-backed services used by diagnostic dialogs in rapid_main.

The VB6 application exposed SQUID, IRM/ARM and vacuum workflows through
separate forms that performed device I/O directly. In rapid_main these have
been converted into explicit backend services with deterministic no-comm
fallbacks so ownership and orchestration can be kept in-process.
"""
from __future__ import annotations

import abc
import json
import random
import time
import math
import re
from dataclasses import dataclass
from typing import Mapping, Protocol, runtime_checkable

from rapidpy_common.hardware import (
    HardwareError,
    MotorAxisConfig,
    MotorSerialClient,
    MotorTelemetry,
    convert_position_to_hole,
)
from rapid_main.config import (
    AfDemagConfig,
    IrmArmConfig,
    SquidConfig,
    SusceptibilityConfig,
    VacuumConfig,
)
from rapid_main.communication_log import CommunicationEvent, CommunicationLogger
from rapid_main.dac import DacChannelConfig, DacCommandPlanner, DacVoltageCommand
from rapid_main.susceptibility_transport import (
    SusceptibilitySerialClient,
    SusceptibilityTransportConfig,
)

try:
    from rapidpy_common.adwin_af import (
        AdwinAFController,
        AdwinBoardConfig,
        AdwinCoilLimits,
        AdwinError,
        AdwinRampRequest,
    )
except Exception:  # pragma: no cover - optional ADwin transport dependency
    AdwinAFController = None
    AdwinBoardConfig = None
    AdwinCoilLimits = None
    AdwinError = RuntimeError
    AdwinRampRequest = None

try:
    from updown_control.app import (
        read_calibration_from_ini,
        SquidCalibration,
        SquidMomentReader,
        VacuumController,
    )
except Exception:  # pragma: no cover - optional import at runtime
    read_calibration_from_ini = None
    SquidCalibration = None
    SquidMomentReader = None
    VacuumController = None


_DEFAULT_SQUID_CALIBRATION = {
    "xcal": -3.410,
    "ycal": -3.470,
    "zcal": -2.516,
    "range_fact": 1.0e-5,
}


def _to_bool_connected(value: object) -> bool:
    """Read connection flags that may be a callable or direct property."""
    if callable(value):
        try:
            return bool(value())
        except TypeError:
            return True
    return bool(value)


def _coerce_squid_calibration(config: dict[str, float], fallback: dict[str, float]) -> object:
    """Create a SquidCalibration instance if available, otherwise store a dict.

    The returned type is intentionally broad because calibration metadata is not
    yet wired into the rest of the SQUID protocol path in this module.
    """
    calibration_dict = dict(fallback)
    if not isinstance(config, dict):
        return fallback
    calibration_dict.update(config)
    if SquidCalibration is None:
        return calibration_dict
    try:
        return SquidCalibration(**calibration_dict)
    except TypeError:
        return SquidCalibration()


class _SimpleSquidCalibration:
    __slots__ = ("xcal", "ycal", "zcal", "range_fact")

    def __init__(self, payload: dict[str, float]) -> None:
        self.xcal = float(payload.get("xcal", _DEFAULT_SQUID_CALIBRATION["xcal"]))
        self.ycal = float(payload.get("ycal", _DEFAULT_SQUID_CALIBRATION["ycal"]))
        self.zcal = float(payload.get("zcal", _DEFAULT_SQUID_CALIBRATION["zcal"]))
        self.range_fact = float(payload.get("range_fact", _DEFAULT_SQUID_CALIBRATION["range_fact"]))


def _ensure_squid_calibration(calibration: object) -> object:
    """Return a calibration object compatible with ``read_moment``."""
    if isinstance(calibration, _SimpleSquidCalibration):
        return calibration
    if isinstance(calibration, SquidCalibration):
        return calibration
    if isinstance(calibration, dict):
        return (
            SquidCalibration(**{**_DEFAULT_SQUID_CALIBRATION, **calibration})
            if SquidCalibration is not None
            else _SimpleSquidCalibration(calibration)
        )
    if calibration is None:
        return (
            SquidCalibration(**_DEFAULT_SQUID_CALIBRATION)
            if SquidCalibration is not None
            else _SimpleSquidCalibration(_DEFAULT_SQUID_CALIBRATION)
        )
    return _SimpleSquidCalibration(_DEFAULT_SQUID_CALIBRATION)


def _parse_irm_axis(axis: str) -> str:
    """Map user-facing axis labels into ADwin active-coil identifiers."""
    norm = (axis or "").strip().lower()
    if norm.startswith("z") or "up" in norm:
        return "axial"
    if norm.startswith("x") or norm.startswith("y"):
        return "transverse"
    raise ValueError(f"Unknown IRM/ARM axis '{axis}'.")


def _parse_ramp_seconds(label: str) -> float:
    label_text = (label or "").strip()
    norm = label_text.lower()
    if "10" in norm:
        return 10.0
    if "30" in norm:
        return 30.0
    if "60" in norm:
        return 60.0
    return 30.0


def _plan_irm_dac_command(
    field_mT: float,
    axis: str,
    cfg: IrmArmConfig | None = None,
) -> DacVoltageCommand:
    """Convert an IRM/ARM field request into an explicit DAC command."""

    cfg = cfg or IrmArmConfig()
    coil = _parse_irm_axis(axis)
    channel = DacChannelConfig(
        channel=0 if coil == "axial" else 1,
        name=f"IRM {coil}",
        minimum_v=0.0,
        maximum_v=float(cfg.irm_max_voltage),
        safe_v=0.0,
    )
    planner = DacCommandPlanner([channel])
    voltage_v = float(cfg.irm_voltage_intercept) + float(field_mT) * float(cfg.irm_voltage_slope)
    return planner.plan_write(channel.channel, voltage_v, purpose="irm-ramp")


def _field_mT_to_volts(
    field_mT: float,
    axis: str = "Z (up-axis)",
    cfg: IrmArmConfig | None = None,
) -> float:
    """Convert mT to a validated ADwin-compatible DAC voltage."""

    command = _plan_irm_dac_command(field_mT, axis, cfg)
    if not command.allowed:
        raise HardwareError(command.command_summary)
    return command.voltage_v


def validate_irm_request(field_mT: float, axis: str, steps: int, cfg: IrmArmConfig) -> None:
    """Validate the entire ramp before configuration changes or actuator calls."""
    field = float(field_mT)
    if not math.isfinite(field) or field < 0:
        raise ValueError("IRM field must be finite and non-negative; backfield requires a polarity adapter.")
    if int(steps) != steps or int(steps) < 1:
        raise ValueError("IRM ramp steps must be a positive integer.")
    # Check both ends: a negative calibration slope can make a later ramp
    # invalid after an earlier ramp has already energized the coil.
    _field_mT_to_volts(field / int(steps), axis, cfg)
    _field_mT_to_volts(field, axis, cfg)


def validate_arm_request(peak_af_mT: float, bias_mT: float, steps: int, cfg: IrmArmConfig) -> None:
    peak, bias = float(peak_af_mT), float(bias_mT)
    if not math.isfinite(peak) or peak < 0 or not math.isfinite(bias):
        raise ValueError("ARM peak and bias must be finite, with a non-negative peak.")
    if int(steps) != steps or int(steps) < 1:
        raise ValueError("ARM steps must be a positive integer.")
    from .arm_bias import plan_arm_bias
    plan_arm_bias(bias, cfg)


def _adwin_event_payload(action: str, **values: object) -> str:
    """Return a stable, single-line ADwin request/result evidence payload."""

    return json.dumps(
        {"action": str(action), **values},
        sort_keys=True,
        separators=(",", ":"),
        default=str,
    )


def _adwin_ramp_request_payload(request: object) -> str:
    return _adwin_event_payload(
        "run_ramp",
        active_coil=str(getattr(request, "active_coil", "")),
        hold_ms=int(getattr(request, "hold_ms", 0)),
        io_rate_hz=float(getattr(request, "io_rate_hz", 0.0)),
        monitor_peak_voltage=float(getattr(request, "peak_monitor_voltage", 0.0)),
        noise_level=int(getattr(request, "noise_level", 0)),
        ramp_down_mode=int(getattr(request, "ramp_down_mode", 0)),
        ramp_mode=int(getattr(request, "ramp_mode", 0)),
        ramp_peak_voltage=float(getattr(request, "ramp_peak_voltage", 0.0)),
        sine_freq_hz=float(getattr(request, "sine_freq_hz", 0.0)),
        slope_down=float(getattr(request, "slope_down", 0.0)),
        slope_up=float(getattr(request, "slope_up", 0.0)),
    )


def _adwin_ramp_result_payload(result: object) -> str:
    return _adwin_event_payload(
        "run_ramp_result",
        down_slope_vps=float(getattr(result, "down_slope_vps")),
        down_start=int(getattr(result, "down_start")),
        in_count=int(getattr(result, "in_count")),
        monitor_peak_v=float(getattr(result, "monitor_peak_v")),
        out_count=int(getattr(result, "out_count")),
        points_per_period=float(getattr(result, "points_per_period")),
        ramp_peak_v=float(getattr(result, "ramp_peak_v")),
        timestep_s=float(getattr(result, "timestep_s")),
        up_count=int(getattr(result, "up_count")),
    )


def _verify_adwin_controller(controller: object, *, label: str) -> int:
    """Verify a live ADwin board without issuing a treatment or relay command."""

    test_version = getattr(controller, "test_version", None)
    if not callable(test_version):
        raise HardwareUnavailableError(f"{label} controller has no readiness probe.")
    version = int(test_version())
    if version == 0:
        raise HardwareUnavailableError(f"{label} ADwin board is unreachable or not booted.")
    return version


#: Marker stamped on every simulated payload, log line, UI label, and file so a
#: simulated run can never be mistaken for hardware evidence.
SIMULATION_MARKER = "SIMULATED"
SIMULATION_STATEMENT = (
    "SIMULATED RUN - synthetic values from a no-communication backend. "
    "Not hardware evidence and not valid for production measurement."
)


class DiagnosticContractError(RuntimeError):
    """Raised when a diagnostic backend fails a command."""


class HardwareUnavailableError(DiagnosticContractError):
    """Raised when a hardware-mode backend cannot be constructed.

    Hardware mode must fail closed: a missing package, driver, port,
    calibration, adapter, or connection blocks preflight instead of quietly
    substituting a simulator.
    """


class UnavailableBackend:
    """Placeholder for a hardware backend that could not be constructed.

    This is deliberately **not** a simulator. It never produces a reading, it
    reports the real construction error, and every action raises. It exists so
    a diagnostics window can still open and show why the device is missing.
    """

    simulated = False

    def __init__(self, name: str, reason: str) -> None:
        self._name = str(name)
        self._reason = str(reason)

    @property
    def name(self) -> str:
        return self._name

    @property
    def reason(self) -> str:
        return self._reason

    def is_connected(self) -> bool:
        return False

    def is_available(self) -> bool:
        return False

    def status(self) -> str:
        return f"{self._name} unavailable: {self._reason}"

    def available_axes(self) -> tuple[str, ...]:
        return ()

    def __getattr__(self, item: str):
        if item.startswith("_"):
            raise AttributeError(item)

        def _refuse(*args: object, **kwargs: object):
            del args, kwargs
            raise HardwareUnavailableError(f"{self._name} unavailable: {self._reason}")

        return _refuse


def build_backend_or_unavailable(name: str, factory, /, *args, **kwargs):
    """Build a hardware backend, or return an explicit unavailable placeholder.

    Used by long-lived UI surfaces that must open even when a device is
    missing. Queue and measurement paths call the factories directly so the
    failure blocks preflight.
    """

    try:
        return factory(*args, **kwargs)
    except HardwareUnavailableError as exc:
        return UnavailableBackend(name, str(exc))


@dataclass(frozen=True)
class DiagnosticStatusLine:
    name: str
    connected: bool
    simulated: bool
    status: str
    fault: bool = False
    fault_reason: str = ""

    @property
    def level(self) -> str:
        if self.fault:
            return "ERROR"
        if self.simulated or not self.connected:
            return "WARNING"
        return "INFO"

    def format_for_console(self) -> str:
        mode = "simulated" if self.simulated else "hardware"
        state = "connected" if self.connected else "disconnected"
        suffix = f" fault={self.fault_reason}" if self.fault_reason else ""
        return f"{self.name}: {state}, {mode}, status={self.status}{suffix}"


def read_backend_status(name: str, backend: object) -> DiagnosticStatusLine:
    connected = False
    simulated = bool(getattr(backend, "simulated", False))
    status = ""
    fault = False
    fault_reason = ""

    try:
        connected = bool(backend.is_connected())  # type: ignore[attr-defined]
    except Exception as exc:
        fault = True
        fault_reason = f"connection read failed: {exc}"

    try:
        status = str(backend.status())  # type: ignore[attr-defined]
    except Exception as exc:
        fault = True
        status = "status unavailable"
        fault_reason = f"{fault_reason}; status read failed: {exc}".strip("; ")

    return DiagnosticStatusLine(
        name=str(name),
        connected=connected,
        simulated=simulated,
        status=status or "status unavailable",
        fault=fault,
        fault_reason=fault_reason,
    )


def collect_diagnostic_status(backends: Mapping[str, object]) -> list[DiagnosticStatusLine]:
    return [read_backend_status(name, backend) for name, backend in backends.items()]


@runtime_checkable
class VacuumBackend(Protocol):
    def is_connected(self) -> bool:
        ...

    def read_pressure(self) -> float:
        """Return current chamber pressure in mTorr."""
        ...

    def set_pump(self, on: bool) -> None:
        """Enable/disable pump control."""
        ...

    def is_pump_on(self) -> bool:
        ...

    def status(self) -> str:
        """Optional status text for UI display."""
        ...


@dataclass(frozen=True)
class VacuumSnapshot:
    pressure_mtorr: float | None
    pump_on: bool
    connected: bool
    status: str
    fault: bool
    fault_reason: str


def read_vacuum_snapshot(
    backend: VacuumBackend,
    *,
    warn_threshold: float,
) -> VacuumSnapshot:
    """Read vacuum state once and classify pressure/readback faults."""
    connected = False
    pump_on = False
    status = ""
    try:
        connected = bool(backend.is_connected())
    except Exception as exc:
        status = f"Connection status unavailable: {exc}"
    try:
        pump_on = bool(backend.is_pump_on())
    except Exception as exc:
        status = f"{status}; pump status unavailable: {exc}".strip("; ")
    try:
        backend_status = str(backend.status())
        status = f"{status}; {backend_status}".strip("; ")
    except Exception as exc:
        status = f"{status}; backend status unavailable: {exc}".strip("; ")

    if not connected:
        return VacuumSnapshot(
            pressure_mtorr=None,
            pump_on=pump_on,
            connected=False,
            status=status or "Vacuum backend is disconnected",
            fault=True,
            fault_reason="Vacuum controller is not connected.",
        )

    threshold = max(0.0, float(warn_threshold))
    pressure_telemetry_available = bool(
        getattr(backend, "pressure_telemetry_available", True)
    )
    if not pressure_telemetry_available:
        if not pump_on:
            return VacuumSnapshot(
                pressure_mtorr=None,
                pump_on=False,
                connected=True,
                status=status or "Vacuum pump is off",
                fault=True,
                fault_reason="Vacuum pump is off and pressure telemetry is unavailable.",
            )
        if threshold > 0.0:
            return VacuumSnapshot(
                pressure_mtorr=None,
                pump_on=True,
                connected=True,
                status=status or "Command-only vacuum controller",
                fault=True,
                fault_reason=(
                    "The connected legacy vacuum controller has no pressure telemetry. "
                    "Configure a live pressure adapter, or explicitly set the warning "
                    "threshold to 0 mTorr for accepted command-only operation."
                ),
            )
        return VacuumSnapshot(
            pressure_mtorr=None,
            pump_on=True,
            connected=True,
            status=(status + "; command-only mode explicitly enabled").strip("; "),
            fault=False,
            fault_reason="",
        )

    try:
        pressure = float(backend.read_pressure())
    except Exception as exc:
        return VacuumSnapshot(
            pressure_mtorr=None,
            pump_on=pump_on,
            connected=connected,
            status=status or "Vacuum status unavailable",
            fault=True,
            fault_reason=f"Vacuum pressure read failed: {exc}",
        )

    if threshold > 0.0 and pressure > threshold:
        return VacuumSnapshot(
            pressure_mtorr=pressure,
            pump_on=pump_on,
            connected=connected,
            status=status or "Vacuum pressure above threshold",
            fault=True,
            fault_reason=f"Vacuum pressure {pressure:.3f} mTorr exceeds {threshold:.3f} mTorr threshold.",
        )

    return VacuumSnapshot(
        pressure_mtorr=pressure,
        pump_on=pump_on,
        connected=connected,
        status=status or "Vacuum ready",
        fault=False,
        fault_reason="",
    )


def require_vacuum_ready(
    backend: VacuumBackend,
    *,
    warn_threshold: float,
) -> VacuumSnapshot:
    """Return a snapshot or raise when vacuum state is unsafe for automation."""
    snapshot = read_vacuum_snapshot(backend, warn_threshold=warn_threshold)
    if snapshot.fault:
        raise DiagnosticContractError(snapshot.fault_reason)
    return snapshot


@runtime_checkable
class IrmArmBackend(Protocol):
    def is_connected(self) -> bool:
        ...

    def apply_irm(self, *, max_field_mT: float, axis: str, ramp_label: str, steps: int) -> str:
        """Run IRM routine and return a short status string."""
        ...

    def apply_arm(self, *, peak_af_mT: float, bias_mT: float, steps: int | None = None) -> str:
        """Run ARM routine and return a short status string."""
        ...

    def reset_field(self) -> str:
        """Reset magnetisation field command path."""
        ...

    def status(self) -> str:
        ...


@dataclass(frozen=True)
class AfDemagCommand:
    label: str
    field_mT: float
    ramp_label: str
    sine_freq_hz: float
    settle_s: float
    tumble: bool
    tumble_pause_s: float
    ramp_peak_voltage: float
    monitor_peak_voltage: float
    calibrated_ramp: object | None = None


@runtime_checkable
class AfDemagBackend(Protocol):
    def is_connected(self) -> bool:
        ...

    def apply_af(self, command: AfDemagCommand) -> str:
        """Apply an AF demagnetization command and return operator status text."""
        ...

    def reset_field(self) -> str:
        """Return AF hardware to a safe relay/field state."""
        ...

    def status(self) -> str:
        ...


def plan_af_demag_command(
    label: str,
    cfg: AfDemagConfig | None = None,
) -> AfDemagCommand:
    """Convert an AF label/config into an explicit treatment command."""
    cfg = cfg or AfDemagConfig()
    _prefix, field_mT, _bias = _parse_af_label(label)
    target_mT = float(field_mT if field_mT is not None else cfg.peak)
    if cfg.calibration_source and cfg.enabled:
        from .af_treatment import plan_calibrated_af_ramp
        ramp = plan_calibrated_af_ramp(label, target_mT, "axial", cfg)
        return AfDemagCommand(str(label), target_mT, str(cfg.ramp_speed), ramp.frequency_hz,
                              ramp.hold_ms / 1000, bool(cfg.tumble), float(cfg.tumble_pause),
                              ramp.ramp_peak_v, ramp.monitor_peak_v, ramp)
    if not math.isfinite(target_mT) or target_mT < 0.0:
        raise DiagnosticContractError(f"AF field must be finite and non-negative: {label}")
    reference_peak = float(cfg.peak)
    if not math.isfinite(reference_peak) or reference_peak <= 0:
        raise DiagnosticContractError("AF reference peak must be finite and positive.")
    if target_mT > reference_peak:
        raise DiagnosticContractError(f"AF field {target_mT:g} mT exceeds configured peak {reference_peak:g} mT.")
    if any(not math.isfinite(float(value)) or float(value) < 0 for value in (cfg.settle, cfg.tumble_pause)):
        raise DiagnosticContractError("AF settling and tumble delay must be finite and non-negative.")

    ramp_hz = _parse_af_ramp_hz(cfg.ramp_speed)
    ramp_peak_voltage = min(10.0, max(0.0, target_mT / reference_peak * 10.0))
    monitor_peak_voltage = ramp_peak_voltage * 0.5
    return AfDemagCommand(
        label=str(label),
        field_mT=target_mT,
        ramp_label=str(cfg.ramp_speed),
        sine_freq_hz=ramp_hz,
        settle_s=max(0.0, float(cfg.settle)),
        tumble=bool(cfg.tumble),
        tumble_pause_s=max(0.0, float(cfg.tumble_pause)),
        ramp_peak_voltage=ramp_peak_voltage,
        monitor_peak_voltage=monitor_peak_voltage,
    )


def _parse_af_label(label: str) -> tuple[str, float | None, None]:
    from .treatment_labels import parse_field_treatment
    try:
        request = parse_field_treatment(label)
        if request.family not in {"AFMAX", "AFZ", "AF"}: raise ValueError("Not AF.")
    except ValueError as exc:
        raise DiagnosticContractError(f"Invalid AF treatment label: {label!r}") from exc
    return request.family, request.field_mT, None


def _parse_af_ramp_hz(label: str) -> float:
    text = (label or "").strip().lower()
    match = re.search(r"(\d+(?:\.\d+)?)\s*hz", text)
    if match:
        return max(0.1, float(match.group(1)))
    if "slow" in text:
        return 3.0
    if "fast" in text:
        return 15.0
    return 8.0


@runtime_checkable
class SquidBackend(Protocol):
    def is_connected(self) -> bool:
        ...

    def test_connection(self) -> bool:
        ...

    def disconnect(self) -> None:
        ...

    def status(self) -> str:
        ...

    def read_squid(self) -> tuple[float, float, float]:
        ...

    def read_susceptibility(self) -> float:
        ...


@runtime_checkable
class SusceptibilityBackend(Protocol):
    @property
    def simulated(self) -> bool:
        ...

    def is_connected(self) -> bool:
        ...

    def test_connection(self) -> bool:
        ...

    def disconnect(self) -> None:
        ...

    def status(self) -> str:
        ...

    def zero(self) -> str:
        ...

    def measure(self) -> float:
        ...

    def communication_events(self) -> tuple[CommunicationEvent, ...]:
        ...


@dataclass(frozen=True)
class SquidSnapshot:
    connected: bool
    status: str
    fault: bool
    fault_reason: str


def read_squid_snapshot(backend: SquidBackend) -> SquidSnapshot:
    """Read SQUID communication state without taking a measurement sample."""
    connected = False
    status = ""
    try:
        connected = bool(backend.is_connected())
    except Exception as exc:
        return SquidSnapshot(
            connected=False,
            status="SQUID connection status unavailable",
            fault=True,
            fault_reason=f"SQUID connection status unavailable: {exc}",
        )

    try:
        status = str(backend.status())
    except Exception as exc:
        status = f"SQUID status unavailable: {exc}"

    if not connected:
        reason = status or "SQUID transport is not connected."
        if "not connected" not in reason.lower() and "disconnected" not in reason.lower():
            reason = f"SQUID transport is not connected: {reason}"
        return SquidSnapshot(
            connected=False,
            status=status or "SQUID transport is not connected",
            fault=True,
            fault_reason=reason,
        )

    return SquidSnapshot(
        connected=True,
        status=status or "SQUID connected",
        fault=False,
        fault_reason="",
    )


def require_squid_ready(backend: SquidBackend) -> SquidSnapshot:
    """Return SQUID communication state or raise before unsafe automation."""
    snapshot = read_squid_snapshot(backend)
    if snapshot.fault:
        raise DiagnosticContractError(snapshot.fault_reason)
    return snapshot


class _BaseBackend(abc.ABC):
    def __init__(self, *, simulated: bool = True) -> None:
        self._simulated = bool(simulated)

    @property
    def simulated(self) -> bool:
        return self._simulated


@dataclass
class _VacuumPhysicsState:
    pump_on: bool
    pressure_mtorr: float
    ambient_pressure_mtorr: float = 120.0
    last_update: float = 0.0

    def __post_init__(self) -> None:
        if self.last_update <= 0:
            self.last_update = time.monotonic()


class VacuumNoCommBackend(_BaseBackend, VacuumBackend):
    """Deterministic no-comm simulator for vacuum diagnostics."""

    def __init__(self, cfg: VacuumConfig | None = None) -> None:
        super().__init__(simulated=True)
        cfg = cfg or VacuumConfig()
        target = max(cfg.target_pressure, 0.05)
        self._cfg = cfg
        now = time.monotonic()
        self._state = _VacuumPhysicsState(
            pump_on=bool(cfg.auto_pump),
            pressure_mtorr=cfg.target_pressure if cfg.target_pressure > 0 else 5.0,
            ambient_pressure_mtorr=max(cfg.warn_threshold, 20.0),
            last_update=now,
        )
        self._rng = random.Random(0xA11)

    def is_connected(self) -> bool:
        return True

    def status(self) -> str:
        return "No-comm simulation" if self._simulated else "Vacuum connected"

    def set_pump(self, on: bool) -> None:
        self._state.pump_on = bool(on)

    def is_pump_on(self) -> bool:
        return bool(self._state.pump_on)

    def read_pressure(self) -> float:
        now = time.monotonic()
        elapsed = max(0.0, now - self._state.last_update)
        self._state.last_update = now

        target = self._cfg.target_pressure if self._state.pump_on else self._state.ambient_pressure_mtorr
        delta = target - self._state.pressure_mtorr
        if abs(delta) > 0.0001:
            # 1st-order relaxation + bounded jitter keeps pressure traces stable.
            tau = max(self._cfg.poll_interval, 1.0)
            blend = min(1.0, elapsed / (tau * 1.5))
            self._state.pressure_mtorr += delta * blend

        jitter = self._rng.uniform(-0.08, 0.08)
        self._state.pressure_mtorr = max(
            0.001,
            self._state.pressure_mtorr + jitter,
        )
        return float(self._state.pressure_mtorr)


class IrmArmNoCommBackend(_BaseBackend, IrmArmBackend):
    """No-comm IRM / ARM simulator used by rapid_main when no dedicated hardware API is attached."""

    def __init__(self, cfg: IrmArmConfig | None = None) -> None:
        super().__init__(simulated=True)
        self._cfg = cfg or IrmArmConfig()
        self._status = "No-comm IRM/ARM simulator"

    def is_connected(self) -> bool:
        return True

    def status(self) -> str:
        return self._status

    def apply_irm(self, *, max_field_mT: float, axis: str, ramp_label: str, steps: int) -> str:
        self._cfg.irm_max_field = max(float(max_field_mT), 0.0)
        self._cfg.irm_axis = axis or self._cfg.irm_axis
        self._cfg.irm_ramp = ramp_label or self._cfg.irm_ramp
        self._cfg.irm_steps = max(1, int(steps))
        self._status = (
            f"IRM requested: {self._cfg.irm_max_field:.2f} mT on {self._cfg.irm_axis}, "
            f"{self._cfg.irm_steps} steps ({self._cfg.irm_ramp})."
        )
        return self._status

    def apply_arm(self, *, peak_af_mT: float, bias_mT: float, steps: int | None = None) -> str:
        self._cfg.arm_peak_af = max(float(peak_af_mT), 0.0)
        self._cfg.arm_bias = float(bias_mT)
        if steps is not None:
            self._cfg.irm_steps = max(1, int(steps))
        self._status = (
            f"ARM requested: peak {self._cfg.arm_peak_af:.2f} mT, bias {self._cfg.arm_bias:.3f} mT."
        )
        return self._status

    def reset_field(self) -> str:
        self._status = "IRM / ARM field reset command issued."
        return self._status


class AfDemagNoCommBackend(_BaseBackend, AfDemagBackend):
    """No-comm AF demagnetizer backend that records explicit treatment commands."""

    def __init__(self, cfg: AfDemagConfig | None = None) -> None:
        super().__init__(simulated=True)
        self._cfg = cfg or AfDemagConfig()
        self.commands: list[AfDemagCommand] = []
        self._status = "No-comm AF demagnetizer simulator"

    def is_connected(self) -> bool:
        return True

    def apply_af(self, command: AfDemagCommand) -> str:
        self.commands.append(command)
        self._status = (
            f"AF requested: {command.field_mT:.3g} mT, "
            f"{command.sine_freq_hz:.3g} Hz, settle {command.settle_s:.3g} s."
        )
        return self._status

    def reset_field(self) -> str:
        self._status = "AF field reset command issued."
        return self._status

    def status(self) -> str:
        return self._status


class AfDemagBackendAdapter(_BaseBackend, AfDemagBackend):
    """ADwin-backed AF demagnetizer treatment backend."""

    def __init__(
        self,
        cfg: AfDemagConfig | None = None,
        *,
        controller: object | None = None,
    ) -> None:
        super().__init__(simulated=False)
        self._cfg = cfg or AfDemagConfig()
        self._communication_logger = CommunicationLogger(
            "ADWIN_AF", port=f"board:{int(self._cfg.board) or 1}", max_payload_chars=2048
        )
        self._controller = controller
        self._last_result = None
        self._status = "AF ADwin backend not connected"
        self._halt_check = None
        self._connect()

    def _connect(self) -> None:
        if self._controller is None and AdwinAFController is None:
            raise HardwareUnavailableError("AF ADwin classes are unavailable on this host.")
        try:
            if self._controller is None:
                board = AdwinBoardConfig(board_num=int(self._cfg.board) or 1)
                for name in ("bin_folder", "boot_file", "process_file"):
                    if getattr(self._cfg, name):
                        setattr(board, name, getattr(self._cfg, name))
                if self._cfg.calibration_source:
                    bits = (self._cfg.axial_relay_bit, self._cfg.transverse_relay_bit)
                    if any(not 0 <= bit < 32 for bit in bits) or bits[0] == bits[1]:
                        raise ValueError("AF axial/transverse relay bits must be distinct configured ADwin outputs.")
                    board.axial_relay_bit, board.trans_relay_bit = bits
                limits = AdwinCoilLimits()
                if self._cfg.calibration_source:
                    limits = AdwinCoilLimits(self._cfg.axial_ramp_max_v, self._cfg.axial_monitor_max_v,
                                            self._cfg.transverse_ramp_max_v, self._cfg.transverse_monitor_max_v)
                self._controller = AdwinAFController(board=board, limits=limits)
            version = _verify_adwin_controller(self._controller, label="AF")
        except Exception as exc:
            self._controller = None
            self._status = f"AF ADwin backend unavailable: {exc}"
            raise HardwareUnavailableError(self._status) from exc
        self._status = f"AF ADwin backend connected (version probe {version})"

    def communication_events(self) -> tuple[CommunicationEvent, ...]:
        return tuple(self._communication_logger.transcript.events)

    def is_connected(self) -> bool:
        return self._controller is not None

    def apply_af(self, command: AfDemagCommand) -> str:
        if command.calibrated_ramp is None:
            raise HardwareError("Live field-based AF requires accepted coil calibration; use AF Tuner for an explicit voltage/clip test.")
        return self.apply_calibrated_af(command.calibrated_ramp)

    def apply_calibrated_af(self, ramp) -> str:
        if self._controller is None or AdwinRampRequest is None:
            raise HardwareError("AF ADwin backend is not connected.")
        from .af_treatment import plan_calibrated_af_ramp
        verified = plan_calibrated_af_ramp(ramp.label, ramp.field_mT, ramp.coil, self._cfg)
        if verified != ramp:
            raise HardwareError("AF calibration changed after treatment planning; rebuild the plan.")
        request = AdwinRampRequest(
            slope_up=ramp.slope_up_vps,
            slope_down=ramp.slope_down_vps,
            peak_monitor_voltage=ramp.monitor_peak_v,
            sine_freq_hz=ramp.frequency_hz,
            ramp_peak_voltage=ramp.ramp_peak_v,
            active_coil=ramp.coil,
            ramp_mode=ramp.ramp_mode,
            hold_ms=ramp.hold_ms,
            ramp_down_mode=1,
            io_rate_hz=ramp.io_rate_hz,
            noise_level=5,
        )
        payload = _adwin_ramp_request_payload(request)
        self._communication_logger.sent(payload, detail=f"AF treatment {ramp.label}; coil={ramp.coil}; calibration={ramp.calibration_source}")
        try:
            options = {"should_cancel": self._halt_check} if self._halt_check is not None else {}
            self._last_result = self._controller.run_ramp(request, **options)
            result_payload = _adwin_ramp_result_payload(self._last_result)
        except Exception as exc:
            self._status = f"AF treatment failed: {exc}"
            self._communication_logger.error(
                f"AF treatment {ramp.label} failed: {exc}", payload=payload
            )
            raise
        self._communication_logger.received(
            result_payload,
            detail=f"AF treatment {ramp.label} completed",
        )
        self._status = (
            f"AF completed: {ramp.field_mT:.3g} mT on {ramp.coil}, "
            f"{ramp.frequency_hz:.3g} Hz."
        )
        return self._status

    def set_halt_check(self, check) -> None:
        self._halt_check = check

    def recover_field(self) -> str:
        if self._controller is None:
            raise HardwareError("AF ADwin backend is not connected.")
        self._controller.recover_safe_field()
        self._status = "AF processes stopped, ramp output zeroed and relay clear verified."
        self._communication_logger.info(self._status)
        return self._status

    def reset_field(self) -> str:
        if self._controller is None or not hasattr(self._controller, "set_af_relays"):
            raise HardwareError("AF ADwin backend is not connected.")
        payload = _adwin_event_payload(
            "set_af_relays", active_coil="off", one_chan_on=True
        )
        self._communication_logger.sent(payload, detail="AF safe field reset")
        try:
            relay_word = self._controller.set_af_relays("off", one_chan_on=True)
        except Exception as exc:
            self._status = f"AF safe field reset failed: {exc}"
            self._communication_logger.error(
                f"AF safe field reset failed: {exc}", payload=payload
            )
            raise
        self._communication_logger.received(
            _adwin_event_payload("set_af_relays_result", relay_word=int(relay_word)),
            detail="AF safe field reset confirmed",
        )
        self._status = "AF field reset confirmed."
        return self._status

    def status(self) -> str:
        return self._status


class IrmArmBackendAdapter(_BaseBackend, IrmArmBackend):
    """Adapter-backed IRM/ARM backend using ADwin AF controller."""

    def __init__(
        self,
        cfg: IrmArmConfig | None = None,
        *,
        controller: object | None = None,
    ) -> None:
        super().__init__(simulated=False)
        self._cfg = cfg or IrmArmConfig()
        self._status = "IRM / ARM ADwin backend not connected"
        self._communication_logger = CommunicationLogger(
            "ADWIN_IRM_ARM", port="board:1", max_payload_chars=2048
        )
        self._controller = controller
        self._last_result = None
        self._connect()

    def _connect(self) -> None:
        if self._controller is None and AdwinAFController is None:
            raise HardwareUnavailableError("IRM / ARM ADwin classes are unavailable on this host.")
        try:
            if self._controller is None:
                board = AdwinBoardConfig()  # defaults match adwin-comms defaults
                self._controller = AdwinAFController(board=board, limits=AdwinCoilLimits())
            version = _verify_adwin_controller(self._controller, label="IRM / ARM")
        except Exception as exc:
            self._controller = None
            self._status = f"IRM / ARM ADwin backend unavailable: {exc}"
            raise HardwareUnavailableError(self._status) from exc
        self._status = f"IRM / ARM ADwin backend connected (version probe {version})"

    def communication_events(self) -> tuple[CommunicationEvent, ...]:
        return tuple(self._communication_logger.transcript.events)

    def is_connected(self) -> bool:
        return self._controller is not None

    def status(self) -> str:
        if not self.is_connected():
            return self._status
        return "IRM / ARM ADwin backend connected"

    def apply_irm(self, *, max_field_mT: float, axis: str, ramp_label: str, steps: int) -> str:
        validate_irm_request(max_field_mT, axis or self._cfg.irm_axis, steps, self._cfg)
        raise HardwareError("Live IRM requires calibrated capacitor charge/readback/fire/discharge through the pulse treatment workflow.")

    def apply_arm(self, *, peak_af_mT: float, bias_mT: float, steps: int | None = None) -> str:
        validate_arm_request(peak_af_mT, bias_mT, self._cfg.irm_steps if steps is None else steps, self._cfg)
        raise HardwareError("Live ARM requires the queue's calibrated axial AF lifecycle and independent MCC bias circuit.")

    def reset_field(self) -> str:
        if self._controller is None or not hasattr(self._controller, "set_af_relays"):
            raise HardwareError("IRM / ARM ADwin backend is not connected.")
        payload = _adwin_event_payload(
            "set_af_relays", active_coil="off", one_chan_on=True
        )
        self._communication_logger.sent(payload, detail="IRM / ARM safe field reset")
        try:
            relay_word = self._controller.set_af_relays("off", one_chan_on=True)
        except Exception as exc:
            self._status = f"IRM / ARM safe field reset failed: {exc}"
            self._communication_logger.error(
                f"IRM / ARM safe field reset failed: {exc}", payload=payload
            )
            raise
        self._communication_logger.received(
            _adwin_event_payload("set_af_relays_result", relay_word=int(relay_word)),
            detail="IRM / ARM safe field reset confirmed",
        )
        self._status = "IRM / ARM field reset confirmed."
        return self._status


class SquidNoCommBackend(_BaseBackend, SquidBackend):
    """No-comm SQUID backend."""

    def __init__(self, cfg: SquidConfig | None = None) -> None:
        super().__init__(simulated=True)
        self._cfg = cfg or SquidConfig()
        self._connected = False
        self._status = "No-comm SQUID simulation"
        self._squid_step_count = 0

    def is_connected(self) -> bool:
        return self._connected

    def status(self) -> str:
        return self._status

    def test_connection(self) -> bool:
        # Deterministic success path for simulation mode; avoids hardware lock-in.
        self._connected = True
        self._status = (
            f"Connected to SQUID serial simulator ({self._cfg.port}:{self._cfg.baud})"
        )
        return True

    def disconnect(self) -> None:
        self._connected = False
        self._status = "No-comm SQUID simulation disconnected"

    def read_squid(self) -> tuple[float, float, float]:
        self._squid_step_count += 1
        # Deterministic synthetic signal with decay + small phase drift.
        angle = self._squid_step_count * 0.55
        decay = 10 ** (-self._squid_step_count / 35.0)
        x = decay * 1.5e-6 * (1.0 + 0.04 * self._squid_step_count) * math.cos(angle)
        y = decay * 1.5e-6 * (1.0 + 0.04 * self._squid_step_count) * math.sin(angle)
        z = decay * 0.8e-6 * ((-1) ** self._squid_step_count)
        return (x, y, z)

    def read_susceptibility(self) -> float:
        x, y, z = self.read_squid()
        return float((x**2 + y**2 + z**2) ** 0.5)


class SusceptibilityNoCommBackend(_BaseBackend, SusceptibilityBackend):
    """Explicit deterministic bridge simulator for training and UI checks."""

    def __init__(self, cfg: SusceptibilityConfig | None = None) -> None:
        super().__init__(simulated=True)
        self._cfg = cfg or SusceptibilityConfig()
        self._zeroed = False
        self._count = 0
        self._status = "No-comm susceptibility bridge simulator"

    def is_connected(self) -> bool:
        return True

    def test_connection(self) -> bool:
        return True

    def disconnect(self) -> None:
        self._status = "No-comm susceptibility bridge simulator"

    def status(self) -> str:
        return self._status

    def zero(self) -> str:
        self._zeroed = True
        self._status = "Susceptibility bridge zeroed (simulated)"
        return "OK"

    def measure(self) -> float:
        if not self._zeroed:
            raise DiagnosticContractError(
                "Zero the simulated susceptibility bridge before measuring."
            )
        self._count += 1
        value = float(self._cfg.scale_factor) * (0.001 / self._count)
        self._status = f"Susceptibility measured: {value:.6g} (simulated)"
        return value

    def communication_events(self) -> tuple[CommunicationEvent, ...]:
        return ()


class SquidBackendAdapter(_BaseBackend, SquidBackend):
    """Adapter-backed SQUID backend using the shared updown_control SQUID client."""

    def __init__(self, cfg: SquidConfig | None = None) -> None:
        super().__init__(simulated=False)
        self._cfg = cfg or SquidConfig()
        self._reader = None
        self._baseline_raw: tuple[float, float, float] | None = None
        self._status = "SQUID transport not connected"

        calibration = _DEFAULT_SQUID_CALIBRATION.copy()
        if read_calibration_from_ini is not None and SquidCalibration is not None:
            try:
                self._calibration = SquidCalibration(**calibration)
            except TypeError:  # pragma: no cover - forward compatibility
                self._calibration = SquidCalibration()  # type: ignore[call-arg]
        else:  # pragma: no cover - optional dependency path
            self._calibration = _coerce_squid_calibration(calibration, _DEFAULT_SQUID_CALIBRATION)
        self._calibration = _ensure_squid_calibration(self._calibration)

    def is_connected(self) -> bool:
        return self._reader is not None

    @property
    def raw_client(self):
        """Underlying atomic 2G client, or ``None`` before connection.

        ``QueueHardwareBackend`` uses this to compose bracketed acquisition
        with the changer motion axes.
        """
        return getattr(self._reader, "raw_client", None)

    def test_connection(self) -> bool:
        if SquidMomentReader is None:
            raise DiagnosticContractError(
                "SQUID transport is unavailable (updown_control dependencies missing)."
            )

        reader = SquidMomentReader()
        try:
            reader.connect(self._cfg.port, baudrate=int(self._cfg.baud or 1200))
        except Exception as exc:
            self._status = f"SQUID connect failed ({self._cfg.port}:{self._cfg.baud}): {exc}"
            raise DiagnosticContractError(self._status) from exc

        try:
            self._baseline_raw = reader.take_baseline()
        except Exception as exc:
            self._status = f"SQUID connected but baseline read failed: {exc}"
            reader.disconnect()
            raise DiagnosticContractError(self._status) from exc

        self._reader = reader
        self._status = f"Connected to SQUID serial ({self._cfg.port}:{self._cfg.baud})"
        return True

    def disconnect(self) -> None:
        reader = self._reader
        self._reader = None
        self._baseline_raw = None
        if reader is not None:
            reader.disconnect()
        self._status = "SQUID transport disconnected"

    def set_demag_step(self, label: str) -> None:
        """Record a treatment label for compatibility with measurement worker calls."""
        # Full SQUID hardware control for AF/IRM treatment is intentionally
        # provided by dedicated AF/IRM pathways. Keep queue/measurement flows
        # from failing when treatment hooks are unavailable.
        self._status = f"SQUID demag step requested: {str(label)}"

    def status(self) -> str:
        return self._status

    def _ensure_connected(self) -> None:
        if self._reader is None or not _to_bool_connected(self._reader.is_connected):
            raise DiagnosticContractError("SQUID transport is not connected.")

    def read_squid(self) -> tuple[float, float, float]:
        self._ensure_connected()
        if self._baseline_raw is None:
            raise DiagnosticContractError("SQUID baseline has not been established.")
        calibration = _ensure_squid_calibration(self._calibration)
        x_emu, y_emu, z_emu, _moment_emu = self._reader.read_moment(  # type: ignore[union-attr]
            calibration,
            self._baseline_raw,
        )
        return (float(x_emu), float(y_emu), float(z_emu))

    def read_susceptibility(self) -> float:
        raise DiagnosticContractError(
            "The SQUID adapter measures magnetic moment, not susceptibility. "
            "Use a configured Bartington susceptibility bridge adapter."
        )


class SusceptibilityBackendAdapter(_BaseBackend, SusceptibilityBackend):
    """Live diagnostic adapter for the Bartington bridge serial protocol."""

    def __init__(
        self,
        cfg: SusceptibilityConfig | None = None,
        *,
        client: SusceptibilitySerialClient | None = None,
    ) -> None:
        super().__init__(simulated=False)
        self._cfg = cfg or SusceptibilityConfig()
        if not self._cfg.enabled:
            raise HardwareUnavailableError(
                "Susceptibility bridge is disabled in application settings."
            )
        transport_cfg = SusceptibilityTransportConfig(
            port=str(self._cfg.port),
            baud=int(self._cfg.baud),
            parity=str(self._cfg.parity),
            bytesize=int(self._cfg.bytesize),
            stopbits=float(self._cfg.stopbits),
            response_timeout_s=float(self._cfg.response_timeout),
            scale_factor=float(self._cfg.scale_factor),
        )
        self._client = client or SusceptibilitySerialClient(transport_cfg)
        self._status = (
            f"Susceptibility bridge ready ({self._cfg.port}:{self._cfg.baud}); disconnected"
        )

    def is_connected(self) -> bool:
        return bool(self._client.is_connected)

    def status(self) -> str:
        return self._status

    def test_connection(self) -> bool:
        """Open the bridge only after an operator-owned workflow requests it."""

        if not self._client.is_connected:
            self._client.connect()
        if not self._client.is_connected:
            raise HardwareUnavailableError(
                f"Susceptibility bridge did not connect on {self._cfg.port}:{self._cfg.baud}."
            )
        self._status = (
            f"Connected to susceptibility bridge ({self._cfg.port}:{self._cfg.baud})"
        )
        return True

    def _require_connected(self) -> None:
        if not self.is_connected():
            raise DiagnosticContractError(
                "Susceptibility bridge is disconnected. Connect it before sending commands."
            )

    def zero(self) -> str:
        self._require_connected()
        reply = self._client.zero()
        self._status = "Susceptibility bridge zero confirmed."
        return reply

    def measure(self) -> float:
        self._require_connected()
        value = self._client.measure()
        self._status = f"Susceptibility bridge reading: {value:.9g}"
        return value

    def disconnect(self) -> None:
        self._client.close()
        self._status = "Susceptibility bridge disconnected."

    def communication_events(self) -> tuple[CommunicationEvent, ...]:
        return self._client.communication_events()


@runtime_checkable
class DCMotorBackend(Protocol):
    """Adapter contract for DC motor diagnostic operations."""

    def connect(self, port: str, baudrate: int) -> None:
        ...

    def disconnect(self) -> None:
        ...

    def is_connected(self) -> bool:
        ...

    def move_motor(self, axis: str, *, target: int, speed: int, wait_for_stop: bool = True) -> tuple[int, int, bool]:
        ...

    def spin_turning(self, *, speed_rps: float, duration_s: float = 60.0) -> tuple[int, int, bool]:
        ...

    def goto_hole(self, *, hole: float) -> float:
        ...

    def read_hole(self, axis: str) -> float:
        ...

    def home_to_top(self) -> None:
        ...

    def home_xy_to_center(self) -> tuple[int, int]:
        ...

    def move_xy_to_corner(self) -> tuple[int, int]:
        ...

    def sample_pickup(self) -> tuple[int, int, bool]:
        ...

    def sample_dropoff(self, use_xy_table: bool = True) -> tuple[int, int, bool]:
        ...

    def available_axes(self) -> tuple[str, ...]:
        """Return the axis labels supported by this backend."""
        ...

    def read_telemetry(self, axis: str) -> MotorTelemetry:
        ...

    def status(self) -> str:
        ...


@dataclass
class _DCMotorSimulationState:
    port: str = ""
    baud: int = 9600
    connected: bool = False
    x_pos: float = 0.0
    y_pos: float = 0.0
    updown_pos: float = 0.0
    turning_pos: float = 0.0
    sample_loaded: bool = False
    last_status: str = "No-comm simulation"
    target_positions: dict[str, int] | None = None


class DCMotorNoCommBackend(_BaseBackend, DCMotorBackend):
    """Deterministic no-comm simulator used by diagnostics."""

    def __init__(self, *, port: str = "COM3", baud: int = 9600) -> None:
        super().__init__(simulated=True)
        self._state = _DCMotorSimulationState(port=port, baud=int(baud))
        self._state.target_positions = {
            "Changer (X)": 0,
            "Turning": 0,
            "Up/Down": 0,
            "Changer (Y)": 0,
        }
        self._axes = {
            "Changer (X)",
            "Turning",
            "Up/Down",
            "Changer (Y)",
        }
        self._rng = random.Random(0xB15)

    def connect(self, port: str, baudrate: int) -> None:
        self._state.port = str(port)
        self._state.baud = int(baudrate)
        self._state.connected = True
        self._state.last_status = f"Connected (sim) to {self._state.port}@{self._state.baud}"

    def disconnect(self) -> None:
        self._state.connected = False
        self._state.last_status = "Disconnected (sim)"

    def is_connected(self) -> bool:
        return bool(self._state.connected)

    def _require_axis(self, axis: str) -> None:
        if axis not in self._axes:
            raise ValueError(f"Unknown axis '{axis}'.")

    def status(self) -> str:
        return self._state.last_status

    def move_motor(self, axis: str, *, target: int, speed: int, wait_for_stop: bool = True) -> tuple[int, int, bool]:
        del wait_for_stop
        self._require_axis(axis)
        jitter = self._rng.uniform(-0.2, 0.2) if speed > 0 else 0.0
        final = int(float(target) + jitter)
        self._state.target_positions[axis] = int(target)
        if axis == "Changer (X)":
            self._state.x_pos = float(target)
        elif axis == "Turning":
            self._state.turning_pos = float(target)
        elif axis == "Up/Down":
            self._state.updown_pos = float(target)
        else:
            self._state.y_pos = float(target)
        return (target, final, True)

    def spin_turning(self, *, speed_rps: float, duration_s: float = 60.0) -> tuple[int, int, bool]:
        if speed_rps == 0:
            return (int(self._state.turning_pos), int(self._state.turning_pos), True)
        self._state.turning_pos -= speed_rps * 360.0 * float(duration_s)
        final = int(self._state.turning_pos)
        return (int(self._state.turning_pos), final, True)

    def goto_hole(self, *, hole: float) -> float:
        try:
            hole_value = float(hole)
        except (TypeError, ValueError) as exc:
            raise ValueError("Hole must be numeric.") from exc
        self._state.x_pos = hole_value * 1000.0
        self._state.target_positions["Changer (X)"] = int(self._state.x_pos)
        return self._state.x_pos / 1000.0

    def read_hole(self, axis: str) -> float:
        self._require_axis(axis)
        if axis != "Changer (X)":
            raise ValueError(f"Reading hole is only supported on Changer (X), got '{axis}'.")
        return convert_position_to_hole(
            int(self._state.x_pos),
            slot_min=1,
            slot_max=101,
            one_step=-1000,
        )

    def home_to_top(self) -> None:
        self._state.updown_pos = 0.0
        self._state.last_status = "Homed to top (sim)"

    def home_xy_to_center(self) -> tuple[int, int]:
        self._state.x_pos = 0.0
        self._state.y_pos = 0.0
        return (int(self._state.x_pos), int(self._state.y_pos))

    def move_xy_to_corner(self) -> tuple[int, int]:
        self._state.x_pos = -3000.0
        self._state.y_pos = -2500.0
        return (int(self._state.x_pos), int(self._state.y_pos))

    def sample_pickup(self) -> tuple[int, int, bool]:
        self._state.sample_loaded = True
        return (1, 1, True)

    def sample_dropoff(self, use_xy_table: bool = True) -> tuple[int, int, bool]:
        del use_xy_table
        self._state.sample_loaded = False
        return (1, 1, True)

    def available_axes(self) -> tuple[str, ...]:
        return tuple(self._axes)

    def read_telemetry(self, axis: str) -> MotorTelemetry:
        self._require_axis(axis)
        if not self._state.connected:
            raise HardwareError("Motor controller is not connected (simulated).")
        if axis == "Changer (X)":
            actual = int(self._state.x_pos)
        elif axis == "Turning":
            actual = int(self._state.turning_pos)
        elif axis == "Up/Down":
            actual = int(self._state.updown_pos)
        else:
            actual = int(self._state.y_pos)

        target = int(self._state.target_positions.get(axis, actual))
        return MotorTelemetry(
            timestamp=time.monotonic(),
            axis_name=axis,
            target_position=target,
            actual_position=actual,
            position_error=int(self._rng.uniform(-6.0, 6.0)),
            velocity_1=int(self._rng.uniform(-120, 120)),
            velocity_2=int(self._rng.uniform(-160, 160)),
            actual_torque=int(self._rng.uniform(-280, 280)),
        )


class DCMotorBackendAdapter(_BaseBackend, DCMotorBackend):
    """Hardware adapter used by the DC Motor diagnostic dialog."""

    def __init__(
        self,
        *,
        port: str,
        baud: int,
        config=None,
    ) -> None:
        super().__init__(simulated=False)
        self._communication_logger = CommunicationLogger(
            "DC_MOTOR", port=str(port), max_payload_chars=2048
        )
        self._client = MotorSerialClient(trace=self._trace)
        self._axes = {
            "Changer (X)": MotorAxisConfig("ChangerX", 1, 1),
            "Turning": MotorAxisConfig("Turning", 2, 2),
            "Up/Down": MotorAxisConfig("UpDown", 3, 3),
            "Changer (Y)": MotorAxisConfig("ChangerY", 4, 4),
        }
        self.wiring_summary = ""
        self._station_baud = None
        if config is not None and config.motor_station.calibration_source:
            from .hardware_contracts import _build_motor_controller_config
            from rapidpy_common.motor_routing import RoutedMotorSerialClient
            station = config.motor_station
            for label, key in (("Changer (X)", "changer_x"), ("Turning", "turning"), ("Up/Down", "updown"), ("Changer (Y)", "changer_y")):
                axis = self._axes[label]
                axis.port = station.ports.get(key, "")
                axis.address = station.addresses.get(key, 0)
            self._client = RoutedMotorSerialClient(_build_motor_controller_config(config), list(self._axes.values()), trace=self._trace)
            self._station_baud = station.baud
            self.wiring_summary = "; ".join(f"{label}: {axis.port} @{axis.address}" for label, axis in self._axes.items())
        self._port = ""
        self._baud = 9600
        self._connected = False
        self._status = "Disconnected (manual)"
        if port:
            self._port = str(port)
            self._baud = int(baud)

    def _trace(self, direction: str, payload: str, detail: str) -> None:
        if detail.startswith("port="):
            self._communication_logger.port = detail.partition(" ")[0].removeprefix("port=")
        if direction == "TX":
            self._communication_logger.sent(payload, detail=detail)
        elif direction == "RX":
            self._communication_logger.received(payload, detail=detail)
        elif direction == "ERROR":
            self._communication_logger.error(detail, payload=payload)
        else:
            self._communication_logger.info(detail or payload)

    def communication_events(self) -> tuple[CommunicationEvent, ...]:
        return tuple(self._communication_logger.transcript.events)

    def _require_move_success(self, action: str, result: object) -> None:
        if bool(getattr(result, "success", False)):
            return
        detail = (
            f"{action} failed: target={getattr(result, 'target', 'unknown')} "
            f"final={getattr(result, 'final_position', 'unknown')}"
        )
        self._communication_logger.error(detail)
        raise HardwareError(detail)

    def connect(self, port: str, baudrate: int) -> None:
        self._port = str(port)
        self._baud = int(baudrate)
        if self._station_baud is not None:
            self._baud = self._station_baud
        self._communication_logger.port = self._port
        try:
            self._client.connect(self._port, baudrate=self._baud)
        except Exception as exc:  # pragma: no cover - hardware transport behavior
            raise HardwareError(f"Unable to connect motor controller at {self._port}:{self._baud}: {exc}") from exc
        self._connected = True

    def disconnect(self) -> None:
        try:
            self._client.disconnect()
        finally:
            self._connected = False

    def is_connected(self) -> bool:
        return bool(self._connected and _to_bool_connected(self._client.is_connected))

    def available_axes(self) -> tuple[str, ...]:
        return tuple(self._axes.keys())

    def _require_connected(self) -> None:
        if not self.is_connected():
            raise HardwareError("Motor controller is not connected.")

    def move_motor(self, axis: str, *, target: int, speed: int, wait_for_stop: bool = True) -> tuple[int, int, bool]:
        axis_cfg = self._axes.get(axis)
        if axis_cfg is None:
            raise ValueError(f"Unknown axis '{axis}'.")
        self._require_connected()
        result = self._client.move_motor(axis_cfg, target, int(speed), wait_for_stop=bool(wait_for_stop))
        self._require_move_success(f"move {axis}", result)
        return (result.target, result.final_position, result.success)

    def spin_turning(self, *, speed_rps: float, duration_s: float = 60.0) -> tuple[int, int, bool]:
        self._require_connected()
        result = self._client.turning_motor_spin(
            self._axes["Turning"],
            speed_rps=float(speed_rps),
            duration_s=float(duration_s),
        )
        self._require_move_success("spin Turning", result)
        return (result.target, result.final_position, result.success)

    def goto_hole(self, *, hole: float) -> float:
        self._require_connected()
        result = self._client.changer_motor_to_hole(self._axes["Changer (X)"], float(hole), wait_for_stop=True)
        self._require_move_success(f"move changer to hole {hole}", result)
        return float(convert_position_to_hole(
            result.final_position,
            slot_min=self._client.config.slot_min,
            slot_max=self._client.config.slot_max,
            one_step=self._client.config.one_step,
        ))

    def read_hole(self, axis: str) -> float:
        self._require_connected()
        if axis != "Changer (X)":
            raise ValueError(f"Reading hole is only supported on Changer (X), got '{axis}'.")
        pos = self._client.read_position(self._axes["Changer (X)"])
        return float(convert_position_to_hole(
            pos,
            slot_min=self._client.config.slot_min,
            slot_max=self._client.config.slot_max,
            one_step=self._client.config.one_step,
        ))

    def home_to_top(self) -> None:
        self._require_connected()
        result = self._client.home_to_top(self._axes["Up/Down"])
        self._require_move_success("home Up/Down to top", result)

    def home_xy_to_center(self) -> tuple[int, int]:
        self._require_connected()
        x_res, y_res = self._client.home_xy_to_center(
            self._axes["Changer (X)"],
            self._axes["Changer (Y)"],
            self._axes["Up/Down"],
        )
        self._require_move_success("home Changer (X) to center", x_res)
        self._require_move_success("home Changer (Y) to center", y_res)
        return (int(x_res.final_position), int(y_res.final_position))

    def move_xy_to_corner(self) -> tuple[int, int]:
        self._require_connected()
        x_res, y_res = self._client.move_xy_to_corner(
            self._axes["Changer (X)"],
            self._axes["Changer (Y)"],
            self._axes["Up/Down"],
        )
        self._require_move_success("move Changer (X) to corner", x_res)
        self._require_move_success("move Changer (Y) to corner", y_res)
        return (int(x_res.final_position), int(y_res.final_position))

    def sample_pickup(self) -> tuple[int, int, bool]:
        self._require_connected()
        result = self._client.sample_pickup(self._axes["Up/Down"])
        self._require_move_success("sample pickup", result)
        return (result.target, result.final_position, result.success)

    def sample_dropoff(self, use_xy_table: bool = True) -> tuple[int, int, bool]:
        self._require_connected()
        result = self._client.sample_dropoff(self._axes["Up/Down"], use_xy_table=bool(use_xy_table))
        self._require_move_success("sample dropoff", result)
        return (result.target, result.final_position, result.success)

    def read_telemetry(self, axis: str) -> MotorTelemetry:
        axis_cfg = self._axes.get(axis)
        if axis_cfg is None:
            raise ValueError(f"Unknown axis '{axis}'.")
        self._require_connected()
        sample = self._client.read_telemetry(axis_cfg)
        return MotorTelemetry(
            timestamp=sample.timestamp,
            axis_name=axis,
            target_position=sample.target_position,
            actual_position=sample.actual_position,
            position_error=sample.position_error,
            velocity_1=sample.velocity_1,
            velocity_2=sample.velocity_2,
            actual_torque=sample.actual_torque,
        )

    def status(self) -> str:
        if self._simulated:
            return "No-comm adapter active"
        if self._connected:
            return f"Connected to {self.wiring_summary or self._port}:{self._baud}"
        return "DC motor adapter ready (disconnected)"


class VacuumBackendAdapter(_BaseBackend, VacuumBackend):
    """Adapter-backed vacuum controller using the shared updown_control vacuum client."""

    def __init__(self, cfg: VacuumConfig | None = None) -> None:
        super().__init__(simulated=False)
        self._cfg = cfg or VacuumConfig()
        self.pressure_telemetry_available = False
        self._communication_logger = CommunicationLogger(
            "VACUUM", port=str(self._cfg.port)
        )
        self._pump_only_controller: VacuumController | None
        self._pump_only_controller = None
        self._status = "Vacuum transport not connected"
        if VacuumController is None:
            raise HardwareUnavailableError("Vacuum controller dependency is unavailable.")
        if not str(self._cfg.port).strip():
            raise HardwareUnavailableError("Vacuum serial port is not configured.")
        self._connect()
        if bool(self._cfg.auto_pump):
            self.set_pump(True)

    def _trace(self, direction: str, payload: str, detail: str) -> None:
        if direction == "TX":
            self._communication_logger.sent(payload, detail=detail)
        elif direction == "RX":
            self._communication_logger.received(payload, detail=detail)
        elif direction == "ERROR":
            self._communication_logger.error(detail, payload=payload)
        else:
            self._communication_logger.info(detail or payload)

    def _connect(self) -> None:
        from rapidpy_common.hardware_safety import HardwareSafetyStore, default_safety_path
        store = HardwareSafetyStore(default_safety_path())
        with store.operation_lease():
            if store.pending() is not None:
                raise HardwareUnavailableError('Recover the original unfinished hardware operation before connecting vacuum controls.')
            controller = VacuumController(trace=self._trace)
            controller.connect(self._cfg.port, baudrate=int(self._cfg.baud or 9600))
        if not controller.is_connected:
            raise HardwareUnavailableError(
                f"Vacuum controller did not connect on {self._cfg.port}:{self._cfg.baud}."
            )
        self._pump_only_controller = controller
        self._status = f"Connected to vacuum serial ({self._cfg.port}:{self._cfg.baud})"

    def communication_events(self) -> tuple[CommunicationEvent, ...]:
        return tuple(self._communication_logger.transcript.events)

    def _require_controller(self) -> VacuumController:
        controller = self._pump_only_controller
        if controller is None or not controller.is_connected:
            raise HardwareUnavailableError("Vacuum controller is not connected.")
        return controller

    def is_connected(self) -> bool:
        return bool(self._pump_only_controller and self._pump_only_controller.is_connected)

    def status(self) -> str:
        return self._status

    def set_pump(self, on: bool) -> None:
        from rapidpy_common.hardware_safety import HardwareSafetyStore, default_safety_path
        store = HardwareSafetyStore(default_safety_path())
        with store.operation_lease():
            if store.pending() is not None:
                raise HardwareUnavailableError('Recover the original unfinished hardware operation before changing vacuum outputs.')
            controller = self._require_controller()
            controller.set_enabled(on)
            if bool(controller.is_enabled) != bool(on):
                raise HardwareError(
                    f"Vacuum controller did not confirm pump {'on' if on else 'off'} state."
                )

    def is_pump_on(self) -> bool:
        return bool(self._require_controller().is_enabled)

    def read_pressure(self) -> float:
        self._require_controller()
        raise HardwareUnavailableError(
            "The legacy vacuum serial controller provides motor/valve acknowledgements "
            "but no pressure telemetry."
        )


def build_vacuum_backend(
    cfg: VacuumConfig,
    *,
    nocomm: bool = False,
    allow_simulation_fallback: bool = False,
) -> VacuumBackend:
    """Build a vacuum backend for the configured mode.

    Hardware mode fails closed: adapter construction failures raise
    :class:`HardwareUnavailableError` instead of returning a simulator.
    """
    if nocomm:
        return VacuumNoCommBackend(cfg)
    try:
        return VacuumBackendAdapter(cfg)
    except Exception as exc:
        if allow_simulation_fallback:
            return VacuumNoCommBackend(cfg)
        raise HardwareUnavailableError(
            f"Vacuum backend is unavailable in hardware mode: {exc}"
        ) from exc


def build_irm_arm_backend(
    cfg: IrmArmConfig,
    *,
    nocomm: bool = False,
    allow_simulation_fallback: bool = False,
) -> IrmArmBackend:
    """Build an IRM/ARM backend for the configured mode.

    Hardware mode fails closed: adapter construction failures raise
    :class:`HardwareUnavailableError` instead of returning a simulator.
    """
    if nocomm:
        return IrmArmNoCommBackend(cfg)
    try:
        return IrmArmBackendAdapter(cfg)
    except Exception as exc:
        if allow_simulation_fallback:
            return IrmArmNoCommBackend(cfg)
        raise HardwareUnavailableError(
            f"IRM/ARM backend is unavailable in hardware mode: {exc}"
        ) from exc


def build_af_demag_backend(
    cfg: AfDemagConfig,
    *,
    nocomm: bool = False,
    allow_simulation_fallback: bool = False,
) -> AfDemagBackend:
    """Build an AF demagnetizer backend for the configured mode.

    Hardware mode fails closed: adapter construction failures raise
    :class:`HardwareUnavailableError` instead of returning a simulator.
    """
    if nocomm:
        return AfDemagNoCommBackend(cfg)
    try:
        return AfDemagBackendAdapter(cfg)
    except Exception as exc:
        if allow_simulation_fallback:
            return AfDemagNoCommBackend(cfg)
        raise HardwareUnavailableError(
            f"AF demagnetizer backend is unavailable in hardware mode: {exc}"
        ) from exc


def build_squid_backend(
    cfg: SquidConfig,
    *,
    nocomm: bool = False,
    allow_simulation_fallback: bool = False,
) -> SquidBackend:
    """Build a SQUID backend for the configured mode.

    Hardware mode fails closed: adapter construction failures raise
    :class:`HardwareUnavailableError` instead of returning a simulator.
    """
    if nocomm:
        return SquidNoCommBackend(cfg)
    try:
        return SquidBackendAdapter(cfg)
    except Exception as exc:
        if allow_simulation_fallback:
            return SquidNoCommBackend(cfg)
        raise HardwareUnavailableError(
            f"SQUID backend is unavailable in hardware mode: {exc}"
        ) from exc


def build_susceptibility_backend(
    cfg: SusceptibilityConfig,
    *,
    nocomm: bool = False,
    allow_simulation_fallback: bool = False,
) -> SusceptibilityBackend:
    """Build an explicit bridge backend without silent live-mode simulation."""

    if nocomm:
        return SusceptibilityNoCommBackend(cfg)
    try:
        return SusceptibilityBackendAdapter(cfg)
    except Exception as exc:
        if allow_simulation_fallback:
            return SusceptibilityNoCommBackend(cfg)
        raise HardwareUnavailableError(
            f"Susceptibility bridge is unavailable in hardware mode: {exc}"
        ) from exc


def build_dcmotor_backend(
    *,
    port: str = "COM3",
    baud: int = 9600,
    nocomm: bool = False,
    allow_simulation_fallback: bool = False,
    config=None,
) -> DCMotorBackend:
    """Build a DC motor diagnostic backend.

    Hardware mode fails closed. ``allow_simulation_fallback`` is an explicit,
    caller-visible opt-in for training or offline development; the returned
    simulator is labelled and reports ``simulated`` as ``True``.
    """
    if nocomm:
        return DCMotorNoCommBackend(port=port, baud=baud)
    try:
        return DCMotorBackendAdapter(port=port, baud=baud, config=config)
    except Exception as exc:
        if allow_simulation_fallback:
            return DCMotorNoCommBackend(port=port, baud=baud)
        raise HardwareUnavailableError(
            f"DC motor backend is unavailable in hardware mode: {exc}"
        ) from exc

