"""Hardware service contracts used by rapid_main workflows.

This module defines the stable protocol used by the measurement engine and
serves as the core abstraction boundary for later real-backend adapters.
"""
from __future__ import annotations

from dataclasses import dataclass
import math
import re
from typing import Protocol

from rapid_main.config import AppConfig
from rapidpy_common.hardware import (
    HardwareError as MotorHardwareError,
    MotorAxisConfig,
    MotorControllerConfig,
    MotorSerialClient,
)


class HardwareError(RuntimeError):
    """Raised for hardware-contract failures."""


class QueueAutomationError(HardwareError, MotorHardwareError):
    """Raised when queue movement/automation fails."""


def _to_bool_connected(value: object) -> bool:
    """Read connection flags that may be callable methods or boolean properties."""
    if callable(value):
        try:
            return bool(value())
        except TypeError:
            return True
    return bool(value)


def _to_float(value: object, *, default: float) -> float:
    try:
        return float(value)
    except (TypeError, ValueError):
        return default


def _to_speed_percent(value: object, *, default: float) -> float:
    val = _to_float(value, default=default)
    if val <= 0:
        return default
    if val > 100:
        return 100.0
    return val


def _scale_motor_speed(base: int, scale_pct: float, *, minimum: int = 100_000) -> int:
    """Scale a motor speed counter using a legacy percent control.

    RapidPy keeps ``changer.speed_xy`` and ``changer.speed_z`` as percentage-like
    fields in the UI for familiar control semantics. To preserve existing defaults,
    a value of ``20`` for ``speed_xy`` keeps ``changer_speed`` unchanged.
    """
    if base <= 0:
        return minimum
    factor = max(0.01, min(100.0, scale_pct)) / 20.0
    scaled = int(base * factor)
    return max(minimum, scaled)


def _build_motor_controller_config(config: AppConfig) -> MotorControllerConfig:
    """Map legacy RapidMain changer settings into motor controller settings."""
    defaults = MotorControllerConfig()
    xy_pct = _to_speed_percent(config.changer.speed_xy, default=20.0)
    z_pct = _to_speed_percent(config.changer.speed_z, default=15.0)
    return MotorControllerConfig(
        changer_speed=_scale_motor_speed(defaults.changer_speed, xy_pct),
        turner_speed=defaults.turner_speed,
        lift_speed_slow=_scale_motor_speed(defaults.lift_speed_slow, z_pct),
        lift_speed_normal=_scale_motor_speed(defaults.lift_speed_normal, z_pct),
        lift_speed_fast=_scale_motor_speed(defaults.lift_speed_fast, z_pct),
    )


@dataclass(frozen=True)
class PreflightResult:
    """Result of backend preflight checks.

    Attributes
    ----------
    ok:
        True only when it is safe to begin the requested workflow.
    blockers:
        Conditions that prevent execution.
    warnings:
        Conditions that allow execution but should be surfaced to the user.
    """
    ok: bool
    blockers: tuple[str, ...]
    warnings: tuple[str, ...]

    @classmethod
    def pass_ok(cls) -> "PreflightResult":
        """Convenience constructor for success path."""
        return cls(ok=True, blockers=(), warnings=())

    @classmethod
    def blocked(cls, *blockers: str, warnings: tuple[str, ...] | None = None) -> "PreflightResult":
        """Convenience constructor for blocked flow."""
        return cls(ok=False, blockers=tuple(blockers), warnings=tuple(warnings or ()))


class MeasurementBackend(Protocol):
    """Protocol for AF/measurement hardware backends."""

    def read_squid(self) -> tuple[float, float, float]:
        """Return `(x, y, z)` in emu."""
        ...

    def set_demag_step(self, label: str) -> None:
        """Apply a treatment step command label."""
        ...

    def read_susceptibility(self) -> float:
        """Return susceptibility reading (emu/Oe)."""
        ...

    def preflight(self) -> PreflightResult:
        """Return preflight status before a measurement sequence begins."""
        ...

    def is_available(self) -> bool:
        """Return ``True`` if backend is currently usable."""
        ...


class QueueAutomationBackend(Protocol):
    """Protocol for queue-motion and changer-side automation commands."""

    def init_up(self, file_id: str) -> None:
        """Prepare or validate a sample set before first-side measurement."""
        ...

    def holder(self, hole: int) -> None:
        """Move to a safe holder position for the current queue step."""
        ...

    def goto_hole(self, hole: int) -> None:
        """Move the changer XY stage to the requested hole index."""
        ...

    def flip(self) -> None:
        """Rotate the sample in the holder for the reverse-side measurement."""
        ...


class MeasurementAutomationBackend(MeasurementBackend, QueueAutomationBackend, Protocol):
    """Combined contract for hardware backends that power both measurement and queue movement."""


# Backward-compatible alias retained for legacy imports in existing modules.
# `HardwareBackend` was used in early modernization drafts before the protocol
# was renamed to `MeasurementBackend`. Keeping both names avoids unnecessary
# churn while we continue migrating references across the UI/service layers.
HardwareBackend = MeasurementBackend


class NoCommBackend:
    """No-comm simulator backend used until full hardware integration is live."""

    def __init__(self) -> None:
        self._step_count = 0
        self._current_hole = 0
        self._flipped = False

    def read_squid(self) -> tuple[float, float, float]:
        x = 1.23e-6 * math.exp(-self._step_count * 0.15) * math.cos(math.radians(self._step_count * 12.5))
        y = 1.23e-6 * math.exp(-self._step_count * 0.15) * math.sin(math.radians(self._step_count * 12.5))
        z = 0.5e-6 * math.exp(-self._step_count * 0.15)
        return (x, y, z)

    def set_demag_step(self, label: str) -> None:
        del label
        self._step_count += 1

    def read_susceptibility(self) -> float:
        return 1e-3 * math.exp(-self._step_count * 0.1)

    def init_up(self, file_id: str) -> None:
        """No-op queue automation adapter for simulator mode."""
        del file_id
        self._flipped = False

    def holder(self, hole: int) -> None:
        """No-op queue automation adapter for simulator mode."""
        self._current_hole = int(hole)

    def goto_hole(self, hole: int) -> None:
        """No-op queue automation adapter for simulator mode."""
        self._current_hole = int(hole)

    def flip(self) -> None:
        """No-op queue automation adapter for simulator mode."""
        self._flipped = not self._flipped

    def return_to_safe_state(self) -> None:
        """No-op safe-state handler for simulator mode."""
        return

    def is_available(self) -> bool:
        return True

    def preflight(self) -> PreflightResult:
        # Simulator preflight is intentionally permissive so development and
        # training can proceed without hardware dependencies.
        return PreflightResult.pass_ok()

    # Optional timeout knobs consumed by the worker
    preflight_timeout: float | None = None
    step_timeout: float | None = None
    read_timeout: float | None = None
    susceptibility_timeout: float | None = None


_DEMAG_LABEL_RE = re.compile(r"^(AF(?:MAX|Z)?|IRM|ARM)\s*(\d+(?:\.\d+)?)?(?:_([0-9]+(?:\.[0-9]+)?))?$")
_THERMAL_LABEL_RE = re.compile(r"^(TT|TH|TEMP)\s*(\d+(?:\.\d+)?)$")


def _parse_demag_label(label: str) -> tuple[str, float | None, float | None]:
    """Parse AF/IRM/ARM treatment labels.

    Returns:
        (prefix, field_mT, bias_mT)
        where ``bias_mT`` is populated only for labels like ``ARM100_0.5``.
    """
    text = (label or "").strip().upper()
    match = _DEMAG_LABEL_RE.match(text)
    if not match:
        return text, None, None
    prefix = match.group(1)
    value = float(match.group(2)) if match.group(2) else None
    bias = float(match.group(3)) if match.group(3) else None
    return prefix, value, bias


def _parse_thermal_label(label: str) -> tuple[str, float | None]:
    """Parse TT/TH/TEMP labels into a normalized prefix and Celsius target."""

    text = (label or "").strip().upper()
    match = _THERMAL_LABEL_RE.match(text)
    if not match:
        return text, None
    return match.group(1), float(match.group(2))


class QueueHardwareBackend(MeasurementAutomationBackend):
    """Adapter that combines motor automation with measurement placeholders."""

    preflight_timeout: float | None = 12.0
    step_timeout: float | None = 45.0
    read_timeout: float | None = 8.0
    susceptibility_timeout: float | None = 2.0
    return_timeout: float | None = 20.0

    def __init__(self, config: AppConfig) -> None:
        self._config = config
        self._motor_config = _build_motor_controller_config(config)
        self._client = MotorSerialClient(self._motor_config)
        try:
            from .diagnostic_services import build_af_demag_backend, build_irm_arm_backend, build_squid_backend
            self._measurement = build_squid_backend(self._config.squid, nocomm=bool(self._config.general.nocomm))
            self._irm_arm = build_irm_arm_backend(self._config.irm_arm, nocomm=bool(self._config.general.nocomm))
            self._af_demag = build_af_demag_backend(self._config.af_demag, nocomm=bool(self._config.general.nocomm))
        except Exception:
            self._measurement = NoCommBackend()
            from .diagnostic_services import AfDemagNoCommBackend, IrmArmNoCommBackend

            self._irm_arm = IrmArmNoCommBackend()
            self._af_demag = AfDemagNoCommBackend()
        self._axes = {
            "changer_x": MotorAxisConfig("ChangerX", 1, 1),
            "turning": MotorAxisConfig("Turning", 2, 2),
            "updown": MotorAxisConfig("UpDown", 3, 3),
            "changer_y": MotorAxisConfig("ChangerY", 4, 4),
        }
        self._connected = False
        self._last_hole = 1
        self._last_flip = False
        self._sample_loaded = False

    def read_squid(self) -> tuple[float, float, float]:
        return self._measurement.read_squid()

    def set_demag_step(self, label: str) -> None:
        """Apply a demagnetization step label to the measurement backend.

        AF labels are routed through the AF demagnetizer backend; IRM/ARM
        labels are routed through the adapter-backed IRM/ARM backend.
        Other labels still use the measurement backend contract.
        """
        step_type, field_mT, bias_mT = _parse_demag_label(label)
        if step_type.startswith("AF"):
            from .diagnostic_services import plan_af_demag_command

            method = getattr(self._af_demag, "apply_af", None)
            if callable(method):
                method(plan_af_demag_command(label, self._config.af_demag))
                return
            raise RuntimeError("AF demagnetizer adapter not available for AF treatment.")

        if step_type == "IRM":
            if field_mT is None:
                raise ValueError(f"IRM step requires numeric field value: {label}")
            method = getattr(self._irm_arm, "apply_irm", None)
            if callable(method):
                method(
                    max_field_mT=float(field_mT),
                    axis=self._config.irm_arm.irm_axis,
                    ramp_label=self._config.irm_arm.irm_ramp,
                    steps=int(max(1, int(self._config.irm_arm.irm_steps))),
                )
                return
            raise RuntimeError("IRM/ARM adapter not available for IRM treatment.")

        if step_type == "ARM":
            if field_mT is None:
                field_mT = self._config.irm_arm.arm_peak_af
            method = getattr(self._irm_arm, "apply_arm", None)
            if callable(method):
                method(
                    peak_af_mT=float(field_mT),
                    bias_mT=float(bias_mT if bias_mT is not None else self._config.irm_arm.arm_bias),
                    steps=int(max(1, int(self._config.irm_arm.irm_steps))),
                )
                return
            raise RuntimeError("IRM/ARM adapter not available for ARM treatment.")

        thermal_type, temperature_c = _parse_thermal_label(label)
        if temperature_c is not None:
            method = getattr(self._measurement, "apply_thermal", None)
            if callable(method):
                method(temperature_c=float(temperature_c), label=f"{thermal_type}{temperature_c:g}")
                return

        method = getattr(self._measurement, "set_demag_step", None)
        if method is None or not callable(method):
            # Measurement hardware is present but may still operate in a read-only
            # or diagnostic-only mode. Queue-driven execution should continue and
            # treat treatment labels as acknowledged.
            return
        method(label)

    def read_susceptibility(self) -> float:
        return self._measurement.read_susceptibility()

    def preflight(self) -> PreflightResult:
        blockers: list[str] = []
        warnings: list[str] = []
        if self._config.general.nocomm:
            return PreflightResult.pass_ok()

        warnings.extend(self._collect_preflight_warnings())
        if not self._measurement.is_connected():
            try:
                self._measurement.test_connection()
            except Exception as exc:
                blockers.append(f"SQUID connection failed: {exc}")

        if self._connected and _to_bool_connected(self._measurement.is_connected):
            return PreflightResult(ok=not blockers, blockers=tuple(blockers), warnings=tuple(warnings))

        port = (self._config.changer.port or "").strip()
        if not port:
            blockers.append("Changer serial port is not configured.")
            return PreflightResult.blocked(*blockers, warnings=tuple(warnings))

        baud = int(self._config.changer.baud or 0)
        if baud <= 0:
            baud = 9600

        try:
            self._client.connect(port, baudrate=baud)
            self._connected = True
            return PreflightResult(ok=not blockers, blockers=tuple(blockers), warnings=tuple(warnings))
        except Exception as exc:
            return PreflightResult.blocked(
                f"Cannot open changer serial port '{port}': {exc}",
                warnings=tuple(warnings),
            )

    def _collect_preflight_warnings(self) -> list[str]:
        return []

    def is_available(self) -> bool:
        return self._connected and _to_bool_connected(self._measurement.is_connected)

    def holder(self, hole: int) -> None:
        """Move to the nearest holder position.

        This maps VB6 queue-holder movement into the motor changer motion domain.
        """
        self._ensure_connected()
        hole = int(hole)

        if self._sample_loaded:
            self._client.sample_dropoff(self._axes["updown"])
            self._sample_loaded = False

        if hole <= 0:
            return

        self._client.changer_motor_to_hole(self._axes["changer_x"], float(hole), wait_for_stop=True)
        self._last_hole = hole
        pickup = self._client.sample_pickup(self._axes["updown"])
        if not pickup.success:
            raise QueueAutomationError(f"sample pickup failed at hole {hole}")
        self._sample_loaded = True

    def goto_hole(self, hole: int) -> None:
        if hole <= 0:
            if self._sample_loaded:
                self._client.sample_dropoff(self._axes["updown"])
                self._sample_loaded = False
            return
        self._ensure_connected()
        hole = int(hole)
        self._client.changer_motor_to_hole(self._axes["changer_x"], float(hole), wait_for_stop=True)
        self._last_hole = hole

    def flip(self) -> None:
        self._ensure_connected()
        # A 180° rotation emulates switching to the reverse face.
        self._client.turning_motor_rotate(self._axes["turning"], 180.0, wait_for_stop=True)
        self._last_flip = not self._last_flip

    def init_up(self, file_id: str) -> None:
        del file_id
        self._ensure_connected()
        if self._sample_loaded:
            try:
                self._client.sample_dropoff(self._axes["updown"])
            finally:
                self._sample_loaded = False
        # Homing the up/down axis before a measurement file run keeps motion
        # deterministic for both orientation passes.
        self._client.home_to_top(self._axes["updown"])

    def return_to_safe_state(self) -> None:
        if not self._connected or not _to_bool_connected(self._client.is_connected):
            return
        try:
            reset_af = getattr(self._af_demag, "reset_field", None)
            if callable(reset_af):
                reset_af()
            if self._sample_loaded:
                self._client.sample_dropoff(self._axes["updown"])
                self._sample_loaded = False
            self._client.halt(self._axes["changer_x"])
            self._client.halt(self._axes["changer_y"])
            self._client.halt(self._axes["turning"])
            self._client.halt(self._axes["updown"])
        finally:
            self._connected = _to_bool_connected(self._client.is_connected)

    def _ensure_connected(self) -> None:
        if self._connected and _to_bool_connected(self._client.is_connected):
            return
        preflight = self.preflight()
        if not preflight.ok:
            raise QueueAutomationError("; ".join(preflight.blockers))


def build_measurement_backend(config: AppConfig) -> MeasurementAutomationBackend:
    """Construct the active measurement backend from configuration."""
    # Phase 2 keeps this as a hard interface boundary. Hardware backends can be
    # added here without changing the UI or worker contracts.
    if config.general.nocomm:
        return NoCommBackend()
    # Queue/changer movement + simulated SQUID/SUSC backend by default.
    return QueueHardwareBackend(config)
