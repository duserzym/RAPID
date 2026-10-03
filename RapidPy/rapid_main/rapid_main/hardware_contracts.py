"""Hardware service contracts used by rapid_main workflows.

This module defines the stable protocol used by the measurement engine and
serves as the core abstraction boundary for later real-backend adapters.
"""
from __future__ import annotations

from dataclasses import dataclass
from contextlib import contextmanager
import dataclasses
import hashlib
import json
import math
from pathlib import Path
import re
from typing import Callable, Protocol, TYPE_CHECKING
import uuid

from rapid_main import software_version
from rapid_main.communication_log import CommunicationLogger
from rapid_main.config import AppConfig

if TYPE_CHECKING:  # pragma: no cover - typing only
    from rapid_main.holder_measurement import HolderMeasurementOutcome
    from rapid_main.holder_state import HolderStateStore
    from rapid_main.squid_transport import BracketedSquidBackend
    from rapid_main.susceptibility_acquisition import SusceptibilityAcquisitionRecord

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


class SusceptibilitySafeStateError(HardwareError):
    """A susceptibility acquisition could not confirm the lift safe state.

    This is a distinct, high-severity outcome: the message keeps the original
    acquisition error (if any) and the safe-return failure side by side.
    """


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
    if config.motor_station.calibration_source:
        from dataclasses import fields
        import math
        values = config.motor_station.controller
        known = {item.name for item in fields(MotorControllerConfig)}
        if any(key not in known or not math.isfinite(float(value)) for key, value in values.items()):
            raise ValueError("Invalid native motor calibration.")
        required = {"slot_min", "slot_max", "one_step", "changer_speed", "turner_speed",
                    "turning_motor_full_rotation", "turning_motor_1rps", "lift_speed_slow",
                    "lift_speed_normal", "lift_speed_fast", "lift_acceleration", "meas_pos",
                    "sample_bottom", "sample_height", "updown_torque_factor", "updown_max_torque", "pickup_torque_throttle"}
        if not required.issubset(values):
            raise ValueError("Incomplete native motor calibration: " + ", ".join(sorted(required - values.keys())))
        if values["one_step"] == 0 or values["turning_motor_full_rotation"] == 0 or values["slot_max"] < values["slot_min"]:
            raise ValueError("Invalid native motor travel calibration.")
        positive = required - {"one_step", "turning_motor_full_rotation", "meas_pos", "sample_bottom"}
        if any(values[key] <= 0 for key in positive) or values["pickup_torque_throttle"] > 1 or values["updown_torque_factor"] > 100 or values["updown_max_torque"] > 32767:
            raise ValueError("Invalid native motor speed/torque calibration.")
        fractional = {"one_step", "pickup_torque_throttle"}
        if any(not float(value).is_integer() for key, value in values.items() if key not in fractional):
            raise ValueError("Native motor integer settings must not contain fractional units.")
        values = dict(values, meas_pos=config.motion.meas_pos, sample_bottom=config.motion.sample_bottom,
                      sample_height=config.motion.sample_height)
        return MotorControllerConfig(**{key: float(value) if key in fractional else int(value) for key, value in values.items()})
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
    """No-comm simulator backend used until full hardware integration is live.

    Every value it returns is synthetic. ``simulated`` is part of the public
    contract so the worker, UI, and output writers can label the run and keep
    it out of production output paths.
    """

    simulated = True

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


_THERMAL_LABEL_RE = re.compile(r"^(TT|TH|TEMP)\s*(\d+(?:\.\d+)?)$")
_MEASUREMENT_ONLY_LABEL_RE = re.compile(r"^(?:NRM(?:-[XYZ])?|REPEAT\d+)$")


def _parse_demag_label(label: str) -> tuple[str, float | None, float | None]:
    """Parse AF/IRM/ARM treatment labels.

    Returns:
        (prefix, field_mT, bias_mT)
        where ``bias_mT`` is populated only for labels like ``ARM100_0.5``.
    """
    text = (label or "").strip().upper()
    from .treatment_labels import parse_field_treatment
    try:
        request = parse_field_treatment(text)
    except ValueError:
        return text, None, None
    return request.family, request.field_mT, request.bias_mT


def _parse_thermal_label(label: str) -> tuple[str, float | None]:
    """Parse TT/TH/TEMP labels into a normalized prefix and Celsius target."""

    text = (label or "").strip().upper()
    match = _THERMAL_LABEL_RE.match(text)
    if not match:
        return text, None
    return match.group(1), float(match.group(2))


def _is_measurement_only_label(label: str) -> bool:
    return bool(_MEASUREMENT_ONLY_LABEL_RE.fullmatch((label or "").strip().upper()))


class QueueHardwareBackend(MeasurementAutomationBackend):
    """Adapter that combines motor automation with SQUID measurement.

    Hardware mode fails closed. A component that cannot be constructed is
    recorded as a preflight blocker instead of being replaced by a simulator,
    and every measurement path raises rather than inventing a value.
    """

    preflight_timeout: float | None = 12.0
    # Physical multi-pass ramps and motion are bounded by their adapters.
    # An outer worker timeout cannot cancel those calls or safely interrupt
    # their cleanup; use the installed cooperative halt check instead.
    step_timeout: float | None = None
    # Bracketed acquisition is bounded by serial/motion adapters and performs
    # safe whole-block recovery. A worker-thread timeout cannot cancel physical
    # I/O and eight seconds is shorter than a normal six-latch acquisition.
    read_timeout: float | None = None
    # Susceptibility acquisition homes, moves slowly to the coil, and waits for
    # a bridge reply bounded by ``susceptibility.response_timeout``. A worker
    # thread timeout cannot cancel that physical I/O, so cancellation is
    # cooperative through ``set_halt_check`` instead.
    susceptibility_timeout: float | None = None
    return_timeout: float | None = 20.0

    def __init__(
        self,
        config: AppConfig,
        *,
        holder_store: "HolderStateStore | None" = None,
        susceptibility_backend=None,
        clock=None,
        safety_store=None,
    ) -> None:
        self._config = config
        from .hardware_safety import HardwareSafetyStore, default_safety_path
        self._safety_store = safety_store if safety_store is not None else HardwareSafetyStore(default_safety_path())
        # Injected only by tests; production uses the real wall/monotonic clock
        # so the VB6 ARC and settling delays are actually observed.
        self._acquisition_clock = clock
        self._backend_errors: list[str] = []
        self._motor_communication_logger = CommunicationLogger(
            "DC_MOTOR", port=str(config.changer.port), max_payload_chars=2048
        )
        self._axes = {
            "changer_x": MotorAxisConfig("ChangerX", 1, 1),
            "turning": MotorAxisConfig("Turning", 2, 2),
            "updown": MotorAxisConfig("UpDown", 3, 3),
            "changer_y": MotorAxisConfig("ChangerY", 4, 4),
        }
        self._client = None
        self._motor_config = None
        try:
            self._motor_config = _build_motor_controller_config(config)
            if config.motor_station.calibration_source:
                from rapidpy_common.motor_routing import RoutedMotorSerialClient
                station = config.motor_station
                for key, axis in self._axes.items():
                    axis.port = station.ports.get(key, "")
                    axis.address = station.addresses.get(key, 0)
                self._client = RoutedMotorSerialClient(self._motor_config, list(self._axes.values()), trace=self._trace_motor_communication)
            else:
                self._client = MotorSerialClient(self._motor_config, trace=self._trace_motor_communication)
        except Exception as exc:
            self._backend_errors.append(f"DC motor station configuration is unavailable: {exc}")
        nocomm = bool(config.general.nocomm)

        from .diagnostic_services import (
            build_af_demag_backend,
            build_irm_arm_backend,
            build_squid_backend,
        )

        self._measurement = self._build_component("SQUID", build_squid_backend, config.squid, nocomm=nocomm)
        self._irm_arm = self._build_component("IRM/ARM", build_irm_arm_backend, config.irm_arm, nocomm=True) if nocomm else None
        self._af_demag = self._build_component(
            "AF demagnetizer", build_af_demag_backend, config.af_demag, nocomm=nocomm
        )
        self._arm_bias = None
        if not nocomm and config.irm_arm.arm_enabled:
            from .arm_bias import ArmBiasBackend
            self._arm_bias = self._build_component("ARM bias", lambda cfg, **_kw: ArmBiasBackend(cfg), config.irm_arm, nocomm=False)
        self._susceptibility = susceptibility_backend
        self._pulse_irm = None
        if not nocomm and (config.pulse_irm.axial_enabled or config.pulse_irm.transverse_enabled):
            from .pulse_treatment import PulseIrmBackend
            def pulse_builder(cfg,**_kwargs):
                return PulseIrmBackend(cfg,getattr(self._af_demag,"_controller",None))
            self._pulse_irm = self._build_component("Pulse IRM",pulse_builder,config.pulse_irm,nocomm=False)

        self._connected = False
        self._last_hole = 1
        self._last_flip = False
        self._sample_loaded = False

        self._holder_store = holder_store if holder_store is not None else _default_holder_store(config)
        self._bracketed: "BracketedSquidBackend | None" = None
        self._acquisition_error = ""
        self._measuring_holder = False
        self._direction_up = True
        self._sample_name = ""
        self._treatment_label = ""
        self._run_id = ""
        self._operator = str(config.general.operator or "")
        self._last_holder_outcome: "HolderMeasurementOutcome | None" = None
        self._susceptibility_records: list["SusceptibilityAcquisitionRecord"] = []
        self._af_treatment_records = []
        self._pulse_treatment_records = []
        self._halt_check: Callable[[], bool] | None = None

    def _trace_motor_communication(
        self, direction: str, payload: str, detail: str
    ) -> None:
        logger = self._motor_communication_logger
        if detail.startswith("port="):
            logger.port = detail.partition(" ")[0].removeprefix("port=")
        if direction == "TX":
            logger.sent(payload, detail=detail)
        elif direction == "RX":
            logger.received(payload, detail=detail)
        elif direction == "ERROR":
            logger.error(detail, payload=payload)
        else:
            logger.info(detail or payload)

    def _require_motion_success(self, action: str, result: object) -> None:
        if bool(getattr(result, "success", False)):
            return
        detail = (
            f"{action} failed: target={getattr(result, 'target', 'unknown')} "
            f"final={getattr(result, 'final_position', 'unknown')}"
        )
        self._motor_communication_logger.error(detail)
        raise QueueAutomationError(detail)

    # -- construction helpers ---------------------------------------------

    def _build_component(self, name: str, factory, cfg, *, nocomm: bool):
        try:
            return factory(cfg, nocomm=nocomm)
        except Exception as exc:
            self._backend_errors.append(f"{name} backend unavailable: {exc}")
            return None

    def __getattr__(self, name: str):
        """Expose the recovery hook only when a real recovery path exists.

        ``MeasurementWorker`` retries a rejected block only when the backend
        implements ``recover_flux_count_discontinuity``. Delegating through
        ``__getattr__`` keeps that contract honest: with no bracketed backend
        the attribute simply does not exist and the run fails instead of
        pretending it recovered.
        """
        if name in ("recover_flux_count_discontinuity", "flux_discontinuity_retries"):
            bracketed = self.__dict__.get("_bracketed")
            if bracketed is not None:
                attribute = getattr(bracketed, name, None)
                if attribute is not None:
                    return attribute
        raise AttributeError(name)

    # -- state exposed to the worker and UI --------------------------------

    @property
    def simulated(self) -> bool:
        return bool(getattr(self._measurement, "simulated", False))

    @property
    def holder_store(self) -> "HolderStateStore":
        return self._holder_store

    @property
    def susceptibility_backend(self):
        """Shared bridge instance used by diagnostics and queued acquisition."""

        return self._susceptibility

    @property
    def last_holder_outcome(self) -> "HolderMeasurementOutcome | None":
        return self._last_holder_outcome

    @property
    def susceptibility_acquisition_records(self) -> tuple["SusceptibilityAcquisitionRecord", ...]:
        """Immutable evidence for every holder and sample bridge acquisition."""

        return tuple(self._susceptibility_records)

    def set_halt_check(self, check: Callable[[], bool] | None) -> None:
        """Install the operator-halt probe used between acquisition phases."""

        self._halt_check = check

    @property
    def af_treatment_records(self):
        return tuple(self._af_treatment_records)

    @property
    def has_unresolved_rotation_fault(self):
        if self._durable_fault_family() in {"rrm", "unknown"}:
            return True
        for record in reversed(getattr(self, "_af_treatment_records", ())):
            if record.schema.startswith("rapidpy.rrm."):
                return not record.safe_state_confirmed
        return False

    @property
    def pulse_treatment_records(self):
        return tuple(self._pulse_treatment_records)

    @property
    def has_unresolved_pulse_fault(self):
        if self._durable_fault_family() in {"pulse", "unknown"}:
            return True
        records = getattr(self,"_pulse_treatment_records",())
        return bool(records and not records[-1].safe_state_confirmed)

    def _safety_profile(self):
        return {name: dataclasses.asdict(getattr(self._config, name))
                for name in ("motor_station", "changer", "motion", "af_demag", "irm_arm", "pulse_irm")}

    def _durable_fault_family(self):
        from .hardware_safety import HardwareSafetyError
        store = getattr(self, "_safety_store", None)
        if store is None or self._config.general.nocomm:
            return ""
        try:
            pending = store.pending()
            return pending["family"] if pending else ""
        except HardwareSafetyError:
            return "unknown"

    @property
    def has_unresolved_hardware_fault(self):
        return bool(self._durable_fault_family()) or self.has_unresolved_rotation_fault or self.has_unresolved_pulse_fault

    def _begin_safety_operation(self, family, plan, *, bias_mT=None):
        store = getattr(self, "_safety_store", None)
        if store is None:  # private object.__new__ fixtures have no native constructor
            return None
        snapshot = dataclasses.asdict(plan)
        if family == "arm":
            snapshot["bias_mT"] = bias_mT
        return store.begin(family, snapshot, self._safety_profile(),
                           sample_id=self._sample_name or "sample", run_id=self._run_id)

    def _finish_safety_operation(self, token, record):
        if token is not None:
            self._safety_store.finish(token, self._safety_profile(), record)

    @contextmanager
    def _safety_operation(self, family, plan, *, bias_mT=None):
        store = getattr(self, "_safety_store", None)
        if store is None:
            yield None
            return
        with store.operation_lease():
            yield self._begin_safety_operation(family, plan, bias_mT=bias_mT)

    @property
    def acquisition_error(self) -> str:
        return self._acquisition_error

    def communication_events(self):
        """Return immutable live SQUID and treatment traffic in time order."""

        events = []
        motor_logger = getattr(self, "_motor_communication_logger", None)
        if motor_logger is not None:
            events.extend(motor_logger.transcript.events)
        for source in (
            self._bracketed,
            self._af_demag,
            self._irm_arm,
            getattr(self, "_arm_bias", None),
            getattr(self, "_susceptibility", None),
        ):
            if source is None or bool(getattr(source, "simulated", False)):
                continue
            provider = getattr(source, "communication_events", None)
            if not callable(provider):
                continue
            events.extend(provider())
        return tuple(sorted(events, key=lambda event: event.timestamp))

    @property
    def transport_recovery_records(self):
        """Return immutable whole-block transport recovery evidence."""

        if self._bracketed is None:
            return ()
        return tuple(self._bracketed.transport_recovery_records)

    def holder_status(self):
        """Holder validity summary for the operator UI."""
        return self._holder_store.status(is_up=self._direction_up)

    def set_measurement_context(
        self,
        *,
        sample_name: str = "",
        treatment_label: str = "",
        run_id: str = "",
        operator: str = "",
        is_up: bool | None = None,
    ) -> None:
        """Record identity carried into each block audit record."""
        if sample_name:
            self._sample_name = str(sample_name)
        if treatment_label:
            self._treatment_label = str(treatment_label)
        if run_id:
            self._run_id = str(run_id)
        if operator:
            self._operator = str(operator)
        if is_up is not None:
            self._direction_up = bool(is_up)

    # -- measurement -------------------------------------------------------

    def read_squid(self):
        """Return one coherent bracketed block, or fail.

        A sample block requires a valid holder correction; a holder block
        subtracts nothing, matching VB6 ``blankHolder``.
        """
        self._ensure_bracketed()
        if self._bracketed is None:
            measurement = self._require_measurement()
            return measurement.read_squid()
        if not self._measuring_holder:
            self._holder_store.require_valid(is_up=self._direction_up)
        return self._bracketed.read_squid()

    def read_susceptibility(self) -> float:
        """Run the VB6 ``Susceptibility_Measure`` sequence for the current sample.

        The value is ``(scaled_sample - holder.susceptibility_raw) *
        SusceptibilityMomentFactorCGS``. The holder value comes only from the
        accepted holder record, never from a magnetic moment or SQUID read.
        """
        if self._config.general.nocomm:
            return self._require_measurement().read_susceptibility()
        blockers = self._susceptibility_blockers(require_holder=True)
        if blockers:
            raise HardwareError("; ".join(blockers))
        holder = self._holder_store.require_valid(is_up=self._direction_up)
        holder_value = holder.require_susceptibility()
        self._ensure_connected()
        record = self._acquire_susceptibility(
            sample_id=self._sample_name or "sample",
            is_holder=False,
            holder_scaled_value=holder_value,
            holder_evidence_id=holder.susceptibility_evidence_id,
        )
        if record.susceptibility is None:
            raise HardwareError("Susceptibility acquisition completed without a value.")
        return float(record.susceptibility)

    def _susceptibility_blockers(self, *, require_holder: bool) -> list[str]:
        """Fail-closed readiness for a live bridge acquisition, without I/O."""

        from .susceptibility_acquisition import SusceptibilityAcquisitionConfig

        blockers: list[str] = []
        cfg = self._config.susceptibility
        if not bool(cfg.enabled):
            blockers.append(
                "Susceptibility bridge is disabled in Settings; SUSC cannot run in hardware mode."
            )
        bridge = self._susceptibility
        if bridge is None:
            blockers.append(
                "No susceptibility bridge backend is shared with the measurement queue."
            )
        elif bool(getattr(bridge, "simulated", False)):
            blockers.append(
                "The susceptibility bridge backend is simulated and cannot provide "
                "hardware-mode evidence."
            )
        elif not all(
            callable(getattr(bridge, name, None))
            for name in ("zero", "measure", "is_connected", "test_connection")
        ):
            reason = str(getattr(bridge, "reason", "") or "")
            blockers.append(
                "Susceptibility bridge is unavailable" + (f": {reason}" if reason else ".")
            )
        scale = _to_float(cfg.scale_factor, default=float("nan"))
        if not math.isfinite(scale) or scale == 0.0:
            blockers.append("Susceptibility scale factor must be finite and nonzero.")
        positions = getattr(self._config, "motion", None)
        sample_height = int(positions.sample_height) if positions is not None else 0
        if positions is None or not positions.configured:
            blockers.append(
                positions.unconfigured_reason()
                if positions is not None
                else "Lift positions are not configured."
            )
        try:
            SusceptibilityAcquisitionConfig(
                coil_position=int(cfg.coil_position),
                sample_height=sample_height,
                moment_factor_cgs=float(cfg.moment_factor_cgs),
            ).validate()
        except (TypeError, ValueError) as exc:
            blockers.append(f"Susceptibility geometry/factor invalid: {exc}")
        if require_holder:
            from .holder_state import HolderStateError

            try:
                holder = self._holder_store.require_valid(is_up=self._direction_up)
                holder.require_susceptibility()
            except HolderStateError as exc:
                blockers.append(f"Holder susceptibility unavailable: {exc}")
        return blockers

    def _acquire_susceptibility(
        self,
        *,
        sample_id: str,
        is_holder: bool,
        holder_scaled_value: float | None = None,
        holder_evidence_id: str = "",
    ) -> "SusceptibilityAcquisitionRecord":
        """Connect the shared bridge if needed and run one audited acquisition."""

        from .squid_transport import MotorVerticalController
        from .susceptibility_acquisition import (
            SusceptibilityAcquisitionConfig,
            SusceptibilityAcquisitionError,
            SusceptibilityAcquisitionService,
        )

        bridge = self._susceptibility
        if not _to_bool_connected(getattr(bridge, "is_connected", False)):
            # Opening the bridge happens only inside an operator-started run
            # that already holds the "susceptibility" ownership lease.
            bridge.test_connection()
        cfg = self._config.susceptibility
        service = SusceptibilityAcquisitionService(
            bridge,
            MotorVerticalController(self._client, self._axes["updown"]),
            config=SusceptibilityAcquisitionConfig(
                coil_position=int(cfg.coil_position),
                sample_height=int(self._config.motion.sample_height),
                moment_factor_cgs=float(cfg.moment_factor_cgs),
            ),
            clock=getattr(self._acquisition_clock, "now", None),
            should_cancel=lambda: bool(self._halt_check is not None and self._halt_check()),
            id_factory=lambda: f"susc-{uuid.uuid4().hex}",
        )
        try:
            record = service.acquire(
                sample_id=sample_id,
                is_holder=is_holder,
                holder_scaled_value=holder_scaled_value,
                holder_evidence_id=holder_evidence_id,
            )
        except SusceptibilityAcquisitionError as exc:
            self._susceptibility_records.append(exc.record)
            if exc.record.safe_return_error:
                raise SusceptibilitySafeStateError(
                    "SAFE-STATE NOT CONFIRMED after susceptibility acquisition: "
                    f"{exc}. Inspect the lift before continuing."
                ) from exc
            raise HardwareError(f"Susceptibility acquisition failed: {exc}") from exc
        self._susceptibility_records.append(record)
        return record

    def set_demag_step(self, label: str) -> None:
        """Apply a demagnetization step label to the measurement backend.

        AF labels are routed through the AF demagnetizer backend; IRM/ARM
        labels are routed through the adapter-backed IRM/ARM backend.
        Explicit measurement-only labels use the measurement backend contract.
        """
        step_type, field_mT, bias_mT = _parse_demag_label(label)
        if label.strip().upper().startswith("RRM"):
            if not self._config.general.nocomm:
                blockers = self._rrm_blockers(label)
                if blockers: raise HardwareError("; ".join(blockers))
                self._execute_rrm_treatment(label)
            else:
                self._measurement.set_demag_step(label)
            self._treatment_label = str(label)
            return
        if not self._config.general.nocomm and (
            step_type in {"AF", "AFMAX", "AFZ", "ARM"}
            or (step_type == "IRM" and field_mT is not None)
        ):
            blockers = self._actuator_blockers(label, step_type, field_mT, bias_mT)
            if blockers:
                raise HardwareError("; ".join(blockers))
        if step_type in {"AF", "AFMAX", "AFZ"}:
            from .diagnostic_services import plan_af_demag_command

            if not self._config.general.nocomm:
                self._execute_af_treatment(label)
                self._treatment_label = str(label)
                return
            method = getattr(self._af_demag, "apply_af", None)
            if callable(method):
                method(plan_af_demag_command(label, self._config.af_demag))
                self._treatment_label = str(label)
                return
            raise HardwareError("AF demagnetizer adapter not available for AF treatment.")

        if step_type == "IRM":
            if field_mT is None:
                raise ValueError(f"IRM step requires numeric field value: {label}")
            if not self._config.general.nocomm:
                from .treatment_labels import parse_field_treatment
                self._execute_pulse_treatment(field_mT, axis=parse_field_treatment(label).axis)
                self._treatment_label = str(label)
                return
            method = getattr(self._irm_arm, "apply_irm", None)
            if callable(method):
                method(
                    max_field_mT=float(field_mT),
                    axis=self._config.irm_arm.irm_axis,
                    ramp_label=self._config.irm_arm.irm_ramp,
                    steps=int(max(1, int(self._config.irm_arm.irm_steps))),
                )
                self._treatment_label = str(label)
                return
            raise HardwareError("IRM/ARM adapter not available for IRM treatment.")

        if step_type == "ARM":
            if field_mT is None:
                field_mT = self._config.irm_arm.arm_peak_af
            if not self._config.general.nocomm:
                self._execute_af_treatment(label, peak_af_mT=field_mT,
                                           bias_mT=self._config.irm_arm.arm_bias if bias_mT is None else bias_mT)
                self._treatment_label = str(label)
                return
            method = getattr(self._irm_arm, "apply_arm", None)
            if callable(method):
                method(
                    peak_af_mT=float(field_mT),
                    bias_mT=float(bias_mT if bias_mT is not None else self._config.irm_arm.arm_bias),
                    steps=int(max(1, int(self._config.irm_arm.irm_steps))),
                )
                self._treatment_label = str(label)
                return
            raise HardwareError("IRM/ARM adapter not available for ARM treatment.")

        thermal_type, temperature_c = _parse_thermal_label(label)
        if temperature_c is not None:
            method = getattr(self._measurement, "apply_thermal", None)
            if callable(method):
                method(temperature_c=float(temperature_c), label=f"{thermal_type}{temperature_c:g}")
                self._treatment_label = str(label)
                return
            if self._config.general.nocomm:
                simulated_step = getattr(self._measurement, "set_demag_step", None)
                if callable(simulated_step):
                    simulated_step(label)
                self._treatment_label = str(label)
                return
            raise HardwareError(
                f"Thermal treatment {label} is planning-only: no production furnace/oven "
                "adapter with temperature readback and safety interlocks is configured."
            )

        if (label or "").strip().upper() == "SUSC" and not self._config.general.nocomm:
            # SUSC applies no treatment. The bridge acquisition runs through
            # ``read_susceptibility`` before this call, matching VB6 ordering.
            reasons = self._susceptibility_blockers(require_holder=True)
            if reasons:
                raise HardwareError(
                    "Susceptibility step SUSC cannot run in hardware mode: " + "; ".join(reasons)
                )
            self._treatment_label = str(label)
            return

        if not self._config.general.nocomm and not _is_measurement_only_label(label):
            raise HardwareError(
                f"Treatment label {label!r} has no production actuator route. "
                "Configure and validate a typed hardware adapter before running it live."
            )

        method = getattr(self._measurement, "set_demag_step", None)
        if method is None or not callable(method):
            # Measurement-only labels do not require an actuator command. The
            # subsequent acquisition/read path remains the source of truth.
            self._treatment_label = str(label)
            return
        method(label)
        self._treatment_label = str(label)

    def validate_treatment_plan(self, labels: tuple[str, ...]) -> PreflightResult:
        """Reject live treatment families that have no production actuator."""

        if self._config.general.nocomm:
            return PreflightResult.pass_ok()
        blockers: list[str] = []
        susceptibility_checked = False
        apply_thermal = getattr(self._measurement, "apply_thermal", None)
        for label in labels:
            if label.strip().upper().startswith("RRM"):
                blockers.extend(self._rrm_blockers(label))
                continue
            step_type, field_mT, bias_mT = _parse_demag_label(label)
            if step_type in {"AF", "AFMAX", "AFZ", "ARM"} or (step_type == "IRM" and field_mT is not None):
                blockers.extend(self._actuator_blockers(label, step_type, field_mT, bias_mT))
                continue
            _thermal_type, temperature_c = _parse_thermal_label(label)
            if temperature_c is not None:
                if not callable(apply_thermal):
                    blockers.append(
                        f"Thermal treatment {label} cannot run in hardware mode: no production "
                        "furnace/oven adapter with temperature readback and safety interlocks "
                        "is configured."
                    )
                continue
            if (label or "").strip().upper() == "SUSC":
                if not susceptibility_checked:
                    susceptibility_checked = True
                    reasons = self._susceptibility_blockers(require_holder=True)
                    if reasons:
                        blockers.append(
                            "Susceptibility step SUSC cannot run in hardware mode: "
                            + "; ".join(reasons)
                        )
                continue
            if _is_measurement_only_label(label):
                continue
            blockers.append(
                f"Treatment label {label!r} cannot run in hardware mode: no production "
                "actuator route is configured and validated."
            )
        if blockers:
            return PreflightResult.blocked(*blockers)
        return PreflightResult.pass_ok()

    def _actuator_blockers(self, label: str, kind: str, field: float | None, bias: float | None) -> list[str]:
        """Read-only request validation; no connection, ramp, or state mutation."""
        from .diagnostic_services import DiagnosticContractError, plan_af_demag_command, validate_irm_request, validate_arm_request

        is_af = kind in {"AF", "AFMAX", "AFZ"}
        is_arm = kind == "ARM"
        backend = getattr(self, "_arm_bias", None) if is_arm else (self._af_demag if is_af else getattr(self,"_pulse_irm",None))
        method = "set_bias_mT" if is_arm else ("apply_calibrated_af" if is_af else "make_circuit")
        reasons = []
        if self.has_unresolved_rotation_fault:
            reasons.append("RRM rotation/field safe state is unverified; stop recovery is required")
        if self.has_unresolved_pulse_fault:
            reasons.append("capacitor/relay safe state is unverified; run discharge recovery before another treatment")
        if backend is None or not callable(getattr(backend, method, None)):
            reasons.append("required actuator adapter is unavailable")
        elif bool(getattr(backend, "simulated", False)):
            reasons.append("simulated actuator cannot execute a live treatment")
        else:
            connected = getattr(backend, "is_connected", None)
            if callable(connected):
                try:
                    if not connected():
                        reasons.append("actuator adapter is disconnected")
                except Exception as exc:
                    reasons.append(f"actuator readiness unavailable: {exc}")
        try:
            if is_af:
                from .af_treatment import plan_af_treatment
                plan_af_treatment(label, self._config.af_demag, self._config.motion.sample_height)
            elif kind == "IRM":
                from .treatment_labels import parse_field_treatment
                self._plan_pulse_treatment(field, axis=parse_field_treatment(label).axis)
            else:
                cfg = self._config.irm_arm
                from .arm_bias import plan_arm_bias
                from .af_treatment import plan_af_treatment
                peak = cfg.arm_peak_af if field is None else field
                plan_arm_bias(cfg.arm_bias if bias is None else bias, cfg)
                plan_af_treatment(f"AFZ{peak:g}", self._config.af_demag, self._config.motion.sample_height)
                reasons.extend(self._actuator_blockers(f"AFZ{peak:g}", "AFZ", peak, None))
        except (TypeError, ValueError, MotorHardwareError, DiagnosticContractError) as exc:
            reasons.append(str(exc))
        return [f"Treatment {label!r} blocked: {reason}" for reason in reasons]

    def _plan_pulse_treatment(self,field, *, axis=None):
        from .pulse_irm import plan_pulse_irm
        from .pulse_circuit import validate_pulse_bindings
        from .pulse_treatment import pulse_motion_target
        axis = str(self._config.irm_arm.irm_axis if axis is None else axis).upper().strip()
        if axis.startswith("Z"):
            coil,angle = "axial",0
        elif axis in {"X","X AXIS","Y","Y AXIS"}:
            coil,angle = "transverse",90 if axis.startswith("Y") else 0
        else:
            raise ValueError("Pulse IRM axis must be Z, X or Y.")
        cfg = self._config.pulse_irm
        validate_pulse_bindings(cfg)
        pulse_motion_target(cfg,self._config.motion.sample_height)
        return plan_pulse_irm(field,coil,cfg),angle

    def _execute_pulse_treatment(self,field, *, axis=None):
        from .pulse_treatment import PulseTreatmentService,PulseTreatmentError
        from .squid_transport import MotorTurningController,MotorVerticalController
        plan,angle = self._plan_pulse_treatment(field, axis=axis)
        self._ensure_connected()
        options = {"should_cancel":self._halt_check}
        if self._acquisition_clock is not None:
            options.update(sleep=self._acquisition_clock.sleep,monotonic=self._acquisition_clock.monotonic,clock=self._acquisition_clock.now)
        circuit = self._pulse_irm.make_circuit(**options)
        service = PulseTreatmentService(circuit,MotorVerticalController(self._client,self._axes["updown"]),
                                       MotorTurningController(self._client,self._axes["turning"]),clock=options.get("clock"))
        with self._safety_operation("pulse", plan) as token:
            try:
                record = service.execute(plan,sample_id=self._sample_name or "sample",run_id=self._run_id,
                                         sample_height=self._config.motion.sample_height,orientation_deg=angle)
            except PulseTreatmentError as exc:
                self._pulse_treatment_records.append(exc.record)
                self._finish_safety_operation(token, exc.record)
                raise HardwareError(f"Pulse treatment failed: {exc}") from exc
            self._pulse_treatment_records.append(record)
            self._finish_safety_operation(token, record)

    def _execute_af_treatment(self, label: str, *, peak_af_mT=None, bias_mT=None) -> None:
        from .af_treatment import AfTreatmentError, AfTreatmentService, AfTreatmentPlan, plan_af_treatment
        from .squid_transport import MotorTurningController, MotorVerticalController

        af_label = label if peak_af_mT is None else f"AFZ{peak_af_mT:g}"
        plan = plan_af_treatment(af_label, self._config.af_demag, self._config.motion.sample_height)
        if peak_af_mT is not None:
            plan = AfTreatmentPlan(label, plan.target_position, plan.passes)
        self._ensure_connected()
        service = AfTreatmentService(
            self._af_demag,
            MotorVerticalController(self._client, self._axes["updown"]),
            MotorTurningController(self._client, self._axes["turning"]),
            should_cancel=self._halt_check,
            bias_adapter=self._arm_bias if peak_af_mT is not None else None,
            bias_mT=bias_mT,
            **({"sleep": self._acquisition_clock.sleep, "clock": self._acquisition_clock.now}
               if self._acquisition_clock is not None else {}),
        )
        with self._safety_operation("arm" if peak_af_mT is not None else "af", plan, bias_mT=bias_mT) as token:
            try:
                record = service.execute(plan, sample_id=self._sample_name or "sample", run_id=self._run_id)
            except AfTreatmentError as exc:
                self._af_treatment_records.append(exc.record)
                self._finish_safety_operation(token, exc.record)
                raise HardwareError(f"AF treatment failed: {exc}") from exc
            self._af_treatment_records.append(record)
            self._finish_safety_operation(token, record)

    def _rrm_blockers(self, label):
        reasons = []
        try:
            from .rrm_treatment import plan_rrm_treatment
            plan = plan_rrm_treatment(label, self._config)
            if plan.bias_mT is not None:
                bias = getattr(self, "_arm_bias", None)
                if bias is None or not callable(getattr(bias, "set_bias_mT", None)) or bias.simulated or not bias.is_connected():
                    reasons.append("RRM bias requires a connected physical ARM bias adapter")
            if self.has_unresolved_rotation_fault or self.has_unresolved_pulse_fault:
                reasons.append("instrument safe state is unverified; recovery is required")
            if self._af_demag is None or not callable(getattr(self._af_demag, "apply_calibrated_af", None)):
                reasons.append("calibrated AF adapter unavailable")
            elif self._af_demag.simulated or not self._af_demag.is_connected():
                reasons.append("a connected physical AF adapter is required")
            if not callable(getattr(self._af_demag, "set_halt_check", None)):
                reasons.append("RRM AF adapter requires cooperative rotation monitoring")
        except (ValueError, TypeError, MotorHardwareError) as exc:
            reasons.append(str(exc))
        return [f"RRM {label!r} blocked: {reason}" for reason in reasons]

    def _execute_rrm_treatment(self, label):
        from .rrm_treatment import RrmTreatmentService, plan_rrm_treatment
        from .af_treatment import AfTreatmentError
        from .squid_transport import MotorTurningController, MotorVerticalController
        plan = plan_rrm_treatment(label, self._config)
        self._ensure_connected()
        options = {"should_cancel": self._halt_check, "bias_adapter": getattr(self, "_arm_bias", None) if plan.bias_mT is not None else None}
        if self._acquisition_clock is not None:
            options.update(sleep=self._acquisition_clock.sleep, monotonic=self._acquisition_clock.monotonic, clock=self._acquisition_clock.now)
        service = RrmTreatmentService(self._af_demag, MotorVerticalController(self._client, self._axes["updown"]),
                                      MotorTurningController(self._client, self._axes["turning"]), **options)
        with self._safety_operation("rrm", plan) as token:
            try:
                record = service.execute(plan, sample_id=self._sample_name or "sample", run_id=self._run_id)
            except AfTreatmentError as exc:
                self._af_treatment_records.append(exc.record)
                self._finish_safety_operation(token, exc.record)
                raise HardwareError(f"RRM treatment failed: {exc}") from exc
            self._af_treatment_records.append(record)
            self._finish_safety_operation(token, record)

    # -- preflight ---------------------------------------------------------

    def preflight(self) -> PreflightResult:
        blockers: list[str] = list(self._backend_errors)
        warnings: list[str] = []
        if self._durable_fault_family():
            blockers.append("A persisted unfinished hardware operation requires verified recovery before motion.")
        if self.has_unresolved_rotation_fault:
            blockers.append("RRM stationary rotation/field is unverified; stop recovery is required before motion.")
        if self.has_unresolved_pulse_fault:
            blockers.append("Pulse capacitor/relay safe state is unverified; discharge recovery is required before motion.")
        if self._config.general.nocomm:
            return PreflightResult.pass_ok()

        warnings.extend(self._collect_preflight_warnings())

        positions = getattr(self._config, "motion", None)
        if positions is not None and not positions.configured:
            blockers.append(positions.unconfigured_reason())

        if self._measurement is None:
            blockers.append("SQUID measurement backend is not available in hardware mode.")
        elif not _to_bool_connected(self._measurement.is_connected):
            try:
                self._measurement.test_connection()
            except Exception as exc:
                blockers.append(f"SQUID connection failed: {exc}")

        if blockers:
            return PreflightResult.blocked(*blockers, warnings=tuple(warnings))

        if self._connected and self._client is not None and _to_bool_connected(self._client.is_connected) and _to_bool_connected(self._measurement.is_connected):
            self._ensure_bracketed()
            return PreflightResult(
                ok=not blockers, blockers=tuple(blockers), warnings=tuple(warnings + self._acquisition_warnings())
            )

        port = (self._config.changer.port or "").strip()
        if not port:
            blockers.append("Changer serial port is not configured.")
            return PreflightResult.blocked(*blockers, warnings=tuple(warnings))

        baud = int(self._config.motor_station.baud if self._config.motor_station.calibration_source else self._config.changer.baud or 0)
        if baud <= 0:
            baud = 9600

        try:
            self._client.connect(port, baudrate=baud)
            self._connected = True
        except Exception as exc:
            return PreflightResult.blocked(
                f"Cannot open changer serial port '{port}': {exc}",
                warnings=tuple(warnings),
            )

        self._ensure_bracketed()
        return PreflightResult(
            ok=True, blockers=(), warnings=tuple(warnings + self._acquisition_warnings())
        )

    def _acquisition_warnings(self) -> list[str]:
        if self._bracketed is None and self._acquisition_error:
            return [f"Bracketed SQUID acquisition unavailable: {self._acquisition_error}"]
        return []

    def _collect_preflight_warnings(self) -> list[str]:
        return []

    def is_available(self) -> bool:
        return (
            self._connected
            and self._measurement is not None
            and _to_bool_connected(self._measurement.is_connected)
        )

    def release_squid_for_external_tool(self) -> None:
        """Close the retained SQUID port before a standalone owner is launched.

        The next hardware preflight reconnects and rebuilds bracketed acquisition.
        No automatic reconnect is attempted here because opening a laboratory
        device is an operator-visible action.
        """

        measurement = self._measurement
        if measurement is None:
            self._bracketed = None
            return
        connected = _to_bool_connected(measurement.is_connected)
        disconnect = getattr(measurement, "disconnect", None)
        if connected and not callable(disconnect):
            raise HardwareError(
                "The active SQUID backend is connected but cannot release its serial port."
            )
        if connected:
            disconnect()
        self._bracketed = None
        self._acquisition_error = ""

    def _require_measurement(self):
        if self._measurement is None:
            raise HardwareError(
                "; ".join(self._backend_errors)
                or "SQUID measurement backend is not available in hardware mode."
            )
        return self._measurement

    def _ensure_bracketed(self) -> None:
        """Compose the bracketed acquisition once transports exist."""
        if self._bracketed is not None or self._config.general.nocomm:
            return
        from .acquisition import BracketedAcquisitionService
        from .squid_transport import (
            BracketedSquidBackend,
            MotorTurningController,
            MotorVerticalController,
            RawSquidTransport,
            SquidTransportConfig,
            acquisition_config_from_app_config,
        )

        positions = getattr(self._config, "motion", None)
        if positions is None or not positions.configured:
            self._acquisition_error = (
                positions.unconfigured_reason() if positions is not None else "motion positions missing"
            )
            return
        raw_client = getattr(self._measurement, "raw_client", None)
        if raw_client is None:
            self._acquisition_error = (
                "the active SQUID backend does not expose a raw 2G client"
            )
            return
        try:
            transport = RawSquidTransport(
                raw_client,
                config=SquidTransportConfig(
                    port=str(self._config.squid.port),
                    baud=int(self._config.squid.baud or 1200),
                    settle_delay_s=float(self._config.squid.settle_time),
                ),
            )
            service = BracketedAcquisitionService(
                transport,
                MotorVerticalController(self._client, self._axes["updown"]),
                MotorTurningController(self._client, self._axes["turning"]),
                config=acquisition_config_from_app_config(
                    self._config,
                    zero_position=positions.zero_position(),
                    measurement_position=positions.measurement_position(),
                ),
                clock=self._acquisition_clock,
            )
        except Exception as exc:
            self._acquisition_error = str(exc)
            return
        self._bracketed = BracketedSquidBackend(
            service,
            holder_provider=self._holder_positions,
            direction_provider=lambda: self._direction_up,
            context_provider=self._block_context,
            simulated=self.simulated,
            communication_events_provider=transport.communication_events,
        )
        self._acquisition_error = ""

    def _holder_positions(self):
        if self._measuring_holder:
            return None
        return self._holder_store.positions_for(self._direction_up)

    def _block_context(self):
        from .acquisition import BlockContext

        current = self._holder_store.current
        return BlockContext(
            sample_name=self._sample_name,
            treatment_label=self._treatment_label,
            run_id=self._run_id,
            operator=self._operator,
            software_version=software_version(),
            config_hash=config_fingerprint(self._config),
            holder_record_id=current.record_version if current is not None else "",
            holder_recorded_iso=current.measured_at_iso if current is not None else "",
            is_holder_block=self._measuring_holder,
            simulated=self.simulated,
        )

    # -- queue automation --------------------------------------------------

    def holder(self, hole: int) -> None:
        """Move to the holder position and measure a replacement correction.

        VB6 ``SampleCommand`` "Holder" is a measurement, not just motion. The
        previous correction stays active unless the new block passes every
        check.
        """
        self._ensure_connected()
        hole = int(hole)
        measure_susceptibility = bool(self._config.susceptibility.enabled)
        if measure_susceptibility:
            blockers = self._susceptibility_blockers(require_holder=False)
            if blockers:
                raise QueueAutomationError(
                    "Holder susceptibility cannot be measured: " + "; ".join(blockers)
                )

        if self._sample_loaded:
            dropoff = self._client.sample_dropoff(self._axes["updown"])
            self._require_motion_success("holder sample dropoff", dropoff)
            self._sample_loaded = False

        if hole > 0:
            changer = self._client.changer_motor_to_hole(
                self._axes["changer_x"], float(hole), wait_for_stop=True
            )
            self._require_motion_success(f"holder changer move to hole {hole}", changer)
            self._last_hole = hole
            pickup = self._client.sample_pickup(self._axes["updown"])
            self._require_motion_success(f"holder sample pickup at hole {hole}", pickup)
            self._sample_loaded = True

        self._measure_holder(hole, measure_susceptibility=measure_susceptibility)

    def _measure_holder(self, hole: int, *, measure_susceptibility: bool = False) -> None:
        from .holder_measurement import HolderMeasurementService

        self._ensure_bracketed()
        if self._bracketed is None:
            raise QueueAutomationError(
                "Holder measurement is unavailable: "
                + (self._acquisition_error or "no bracketed SQUID acquisition is configured")
            )
        previous_sample = self._sample_name
        self._sample_name = "Holder"
        self._measuring_holder = True
        susceptibility_record = None
        try:
            if measure_susceptibility:
                # VB6 Measure: the holder bridge read precedes Measure_Read.
                try:
                    susceptibility_record = self._acquire_susceptibility(
                        sample_id=_holder_identity(hole), is_holder=True
                    )
                except Exception:
                    self._persist_holder_susceptibility_evidence(strict=False)
                    raise
                self._persist_holder_susceptibility_evidence(strict=True)
            service = HolderMeasurementService(
                self._bracketed.read_squid,
                self._holder_store,
                recover=self._bracketed.recover_flux_count_discontinuity,
                averaging_cycles=max(1, int(self._config.squid.samples_per_pos or 1)),
                flux_discontinuity_retries=int(self._bracketed.flux_discontinuity_retries),
                clock=getattr(self._acquisition_clock, "now", None),
            )
            outcome = service.measure(
                holder_id=_holder_identity(hole),
                hole=hole,
                susceptibility=susceptibility_record,
            )
        finally:
            self._measuring_holder = False
            self._sample_name = previous_sample
        self._last_holder_outcome = outcome
        if not outcome.installed:
            raise QueueAutomationError(outcome.rejection_reason)

    def _persist_holder_susceptibility_evidence(self, *, strict: bool) -> None:
        """Publish the latest holder bridge record beside the holder store.

        The accepted holder names this file by ``susceptibility_evidence_id``.
        A write failure blocks installation so the identity always resolves.
        """
        from .susceptibility_acquisition import write_susceptibility_acquisition

        if not self._susceptibility_records:
            return
        record = self._susceptibility_records[-1]
        if not record.is_holder:
            return
        store_path = self._holder_store.path
        if store_path is None:
            return
        target = store_path.parent / "holder_susceptibility" / f"{record.acquisition_id}.json"
        if target.exists():
            return
        try:
            write_susceptibility_acquisition(target, record)
        except Exception as exc:
            if not strict:
                # Never hide the acquisition failure behind an evidence write.
                self._motor_communication_logger.error(
                    f"holder susceptibility evidence write failed: {exc}"
                )
                return
            raise QueueAutomationError(
                f"Holder susceptibility evidence could not be written to {target}: {exc}"
            ) from exc

    def goto_hole(self, hole: int) -> None:
        if hole <= 0:
            if self._sample_loaded:
                dropoff = self._client.sample_dropoff(self._axes["updown"])
                self._require_motion_success("sample dropoff", dropoff)
                self._sample_loaded = False
            return
        self._ensure_connected()
        hole = int(hole)
        changer = self._client.changer_motor_to_hole(
            self._axes["changer_x"], float(hole), wait_for_stop=True
        )
        self._require_motion_success(f"changer move to hole {hole}", changer)
        self._last_hole = hole

    def flip(self) -> None:
        self._ensure_connected()
        # A 180° rotation emulates switching to the reverse face.
        turn = self._client.turning_motor_rotate(
            self._axes["turning"], 180.0, wait_for_stop=True
        )
        self._require_motion_success("sample flip 180 degrees", turn)
        self._last_flip = not self._last_flip
        self._direction_up = not self._direction_up

    def init_up(self, file_id: str) -> None:
        del file_id
        self._ensure_connected()
        if self._sample_loaded:
            dropoff = self._client.sample_dropoff(self._axes["updown"])
            self._require_motion_success("initial sample dropoff", dropoff)
            self._sample_loaded = False
        # Homing the up/down axis before a measurement file run keeps motion
        # deterministic for both orientation passes.
        homed = self._client.home_to_top(self._axes["updown"])
        self._require_motion_success("home Up/Down to top", homed)
        self._direction_up = True

    def return_to_safe_state(self) -> None:
        store = getattr(self, "_safety_store", None)
        pending = store.pending() if store is not None and not self._config.general.nocomm else None
        if pending:
            with store.operation_lease():
                pending = store.pending()
                if pending:
                    if pending["family"] not in {"af_diagnostic", "motion_diagnostic", "station_diagnostic"}:
                        pending = store.pending(self._safety_profile())
                    self._recover_durable_treatment(pending)
        records = getattr(self,"_pulse_treatment_records",())
        if records and not records[-1].safe_state_confirmed:
            from .pulse_circuit import PulseCircuitError
            pulse = getattr(self,"_pulse_irm",None)
            if pulse is None:
                raise HardwareError("Pulse safe state is unverified; motor return and relay reset are inhibited.")
            options = {}
            if self._acquisition_clock is not None:
                options.update(sleep=self._acquisition_clock.sleep,monotonic=self._acquisition_clock.monotonic,clock=self._acquisition_clock.now)
            try:
                recovery = pulse.make_circuit(**options).recover_safe_state(sample_id=self._sample_name or "sample",run_id=self._run_id)
            except PulseCircuitError as exc:
                self._pulse_treatment_records.append(exc.record)
                raise HardwareError(f"Pulse discharge recovery failed; motor return and relay reset inhibited: {exc}") from exc
            self._pulse_treatment_records.append(recovery)
        if self.has_unresolved_rotation_fault:
            from .rrm_treatment import RrmTreatmentService
            from .af_treatment import AfTreatmentError
            from .squid_transport import MotorTurningController, MotorVerticalController
            prior = next(record for record in reversed(self._af_treatment_records) if record.schema.startswith("rapidpy.rrm."))
            service = RrmTreatmentService(self._af_demag, MotorVerticalController(self._client, self._axes["updown"]),
                                          MotorTurningController(self._client, self._axes["turning"]), bias_adapter=getattr(self, "_arm_bias", None) if prior.plan.bias_mT is not None else None)
            try:
                recovery = service.recover(prior.plan, sample_id=self._sample_name or prior.sample_id, run_id=self._run_id or prior.run_id)
            except AfTreatmentError as exc:
                self._af_treatment_records.append(exc.record)
                raise HardwareError(f"RRM stop recovery failed; lift return inhibited: {exc}") from exc
            self._af_treatment_records.append(recovery)
        bias = getattr(self, "_arm_bias", None)
        bias_error = ""
        if bias is not None and callable(getattr(bias, "clear_bias", None)):
            try:
                bias.clear_bias()
            except Exception as exc:
                bias_error = f"ARM bias reset failed: {exc}"
        if not self._connected or not _to_bool_connected(self._client.is_connected):
            if bias_error:
                raise HardwareError(bias_error)
            return
        errors: list[str] = [bias_error] if bias_error else []
        try:
            reset_af = getattr(self._af_demag, "reset_field", None)
            if callable(reset_af):
                try:
                    reset_af()
                except Exception as exc:
                    errors.append(f"AF reset failed: {exc}")
            if self._sample_loaded:
                try:
                    dropoff = self._client.sample_dropoff(self._axes["updown"])
                    self._require_motion_success("safe-state sample dropoff", dropoff)
                except Exception as exc:
                    errors.append(str(exc))
                else:
                    self._sample_loaded = False
            for axis_name in ("changer_x", "changer_y", "turning", "updown"):
                try:
                    self._client.halt(self._axes[axis_name])
                except Exception as exc:
                    errors.append(f"halt {axis_name} failed: {exc}")
        finally:
            self._connected = _to_bool_connected(self._client.is_connected)
        if errors:
            detail = "Safe-state return incomplete: " + "; ".join(errors)
            self._motor_communication_logger.error(detail)
            raise QueueAutomationError(detail)

    def _recover_durable_treatment(self, pending):
        """Recover only the persisted station; never replay its treatment."""
        if pending["family"] == "af_diagnostic":
            raise HardwareError("An unfinished ADwin helper diagnostic requires recovery in its original helper using its original board settings.")
        if pending["family"] == "motion_diagnostic":
            raise HardwareError("An unfinished motor helper diagnostic requires verified stop recovery in its original helper using its original port and axis settings.")
        if pending["family"] == "station_diagnostic":
            raise HardwareError("A held vacuum/lift diagnostic requires recovery in Up/Down Control with its original participating ports before main-app motion or treatment.")
        sample_id, run_id = pending["sample_id"], pending["run_id"]
        token = pending["token"]
        if pending["family"] == "pulse":
            from .pulse_circuit import PulseCircuitError
            if self._pulse_irm is None:
                raise HardwareError("The latched pulse station is unavailable for discharge recovery.")
            options = {}
            if self._acquisition_clock is not None:
                options.update(sleep=self._acquisition_clock.sleep, monotonic=self._acquisition_clock.monotonic, clock=self._acquisition_clock.now)
            try:
                record = self._pulse_irm.make_circuit(**options).recover_safe_state(sample_id=sample_id, run_id=run_id)
            except PulseCircuitError as exc:
                self._pulse_treatment_records.append(exc.record)
                self._finish_safety_operation(token, exc.record)
                raise HardwareError(f"Persistent pulse discharge recovery failed: {exc}") from exc
            record = self._recover_pulse_motion(record, pending)
            self._pulse_treatment_records.append(record)
            self._finish_safety_operation(token, record)
            if not record.safe_state_confirmed:
                raise HardwareError("Pulse discharge succeeded but specimen motion recovery is unverified: " + "; ".join(record.cleanup_errors))
            return
        if pending["family"] in {"af", "arm"}:
            self._recover_durable_af(pending)
            return
        from .rrm_treatment import RrmTreatmentPlan, RrmTreatmentService
        from .af_treatment import CalibratedAfRamp, AfTreatmentError
        from .squid_transport import MotorTurningController, MotorVerticalController
        snapshot = dict(pending["plan"])
        snapshot["ramp"] = CalibratedAfRamp(**snapshot["ramp"])
        plan = RrmTreatmentPlan(**snapshot)
        if self._client is None or self._af_demag is None:
            raise HardwareError("The latched RRM station is unavailable for stop recovery.")
        # Reset on the existing ADwin board: booting could switch relays while
        # an unknown process is still running. No treatment command is sent.
        self._af_demag.recover_field()
        if not _to_bool_connected(self._client.is_connected):
            station = self._config.motor_station
            self._client.connect(self._config.changer.port, baudrate=station.baud)
            self._connected = True
        # The recovery service independently verifies stop before returning
        # reference/lift. Route its reset to the non-booting recovery method.
        class RecoveryAdapter:
            def reset_field(_self):
                return self._af_demag.recover_field()
        service = RrmTreatmentService(RecoveryAdapter(), MotorVerticalController(self._client, self._axes["updown"]),
                                      MotorTurningController(self._client, self._axes["turning"]),
                                      bias_adapter=self._arm_bias if plan.bias_mT is not None else None)
        if plan.bias_mT is not None and self._arm_bias is None:
            raise HardwareError("Latched RRM bias adapter is unavailable; recovery remains pending.")
        try:
            record = service.recover(plan, sample_id=sample_id, run_id=run_id)
        except AfTreatmentError as exc:
            self._af_treatment_records.append(exc.record)
            self._finish_safety_operation(token, exc.record)
            raise HardwareError(f"Persistent RRM stop recovery failed: {exc}") from exc
        self._af_treatment_records.append(record)
        self._finish_safety_operation(token, record)

    def _connect_recovery_transport(self):
        if self._client is None:
            raise HardwareError("Motor station unavailable for recovery.")
        if not _to_bool_connected(self._client.is_connected):
            cfg = self._config
            baud = cfg.motor_station.baud if cfg.motor_station.calibration_source else cfg.changer.baud
            self._client.connect(cfg.changer.port, baudrate=baud)
        self._connected = True

    def _recover_pulse_motion(self, circuit_record, pending):
        from datetime import datetime, timezone
        from .pulse_treatment import PulseTreatmentRecord
        from .pulse_circuit import PulsePhase
        from .squid_transport import MotorTurningController, MotorVerticalController
        phases, errors = [], []
        def action(name, operation, check=lambda value: True):
            try:
                value = operation()
                if not check(value):
                    raise HardwareError("Recovery motion was not verified.")
                phases.append(PulsePhase(name, datetime.now(timezone.utc).isoformat(), True, ""))
                return True
            except Exception as exc:
                phases.append(PulsePhase(name, datetime.now(timezone.utc).isoformat(), False, str(exc)))
                errors.append(f"{name}: {exc}")
                return False
        if circuit_record.safe_state_confirmed is not True:
            errors.append("Capacitor/relay discharge is unverified; motor recovery withheld.")
        elif action("connect_recovery_transport", self._connect_recovery_transport):
            turning = MotorTurningController(self._client, self._axes["turning"])
            vertical = MotorVerticalController(self._client, self._axes["updown"])
            if action("stationary_turning", turning.stop_spin, lambda result: result.success is True):
                if action("restore_reference", turning.restore_spin_reference, lambda result: result.ok is True and math.isfinite(result.actual)):
                    action("home_after", vertical.home_to_top, lambda result: result.ok is True and math.isfinite(result.actual))
        return PulseTreatmentRecord("irm-" + uuid.uuid4().hex, pending["sample_id"], pending["run_id"], 0, 0,
                                    circuit_record, tuple(phases), not errors, "", tuple(errors),
                                    schema="rapidpy.irm.recovery_treatment.v1")

    def _recover_durable_af(self, pending):
        from datetime import datetime, timezone
        from .af_treatment import AfTreatmentPlan, AfPass, CalibratedAfRamp, AfTreatmentPhase, AfTreatmentRecord, AfTreatmentError
        from .squid_transport import MotorTurningController, MotorVerticalController
        snapshot = dict(pending["plan"])
        bias_mT = snapshot.pop("bias_mT", None)
        snapshot["passes"] = tuple(AfPass(item["angle_deg"], CalibratedAfRamp(**item["ramp"]), item["pause_before_s"]) for item in snapshot["passes"])
        plan = AfTreatmentPlan(**snapshot)
        phases, errors = [], []
        def action(name, operation, check=lambda value: True):
            try:
                value = operation()
                if not check(value):
                    raise HardwareError("Hardware completion was not verified.")
                phases.append(AfTreatmentPhase(name, True, datetime.now(timezone.utc).isoformat()))
            except Exception as exc:
                phases.append(AfTreatmentPhase(name, False, datetime.now(timezone.utc).isoformat(), str(exc)))
                errors.append(f"{name}: {exc}")
        if pending["family"] == "arm":
            action("bias_clear", lambda: self._arm_bias.clear_bias())
        action("field_recovery", lambda: self._af_demag.recover_field())
        # Transport setup does not move an axis. Every subsequent movement is
        # withheld until both independent field cleanup actions succeed.
        if not errors:
            action("connect_recovery_transport", self._connect_recovery_transport)
        if not errors:
            turning = MotorTurningController(self._client, self._axes["turning"])
            vertical = MotorVerticalController(self._client, self._axes["updown"])
            action("stationary_turning", turning.stop_spin, lambda result: result.success is True)
            if not errors:
                action("restore_reference", turning.restore_spin_reference, lambda result: result.ok is True and math.isfinite(result.actual))
            if not errors:
                action("home_after", vertical.home_to_top, lambda result: result.ok is True and math.isfinite(result.actual))
        record = AfTreatmentRecord("af-" + uuid.uuid4().hex, pending["sample_id"], pending["run_id"], plan, 0,
                                   tuple(phases), "", tuple(errors), not errors, False,
                                   schema=f"rapidpy.{pending['family']}.recovery.v1", bias_mT=bias_mT)
        self._af_treatment_records.append(record)
        self._finish_safety_operation(pending["token"], record)
        if errors:
            raise AfTreatmentError(record)

    def _ensure_connected(self) -> None:
        if self._durable_fault_family():
            raise HardwareError("An unfinished hardware operation persists; motor motion is inhibited until verified recovery.")
        if self.has_unresolved_rotation_fault:
            raise HardwareError("RRM stationary rotation/field is unverified; motor motion is inhibited until stop recovery.")
        if self.has_unresolved_pulse_fault:
            raise HardwareError("Pulse capacitor/relay safe state is unverified; motor motion is inhibited until discharge recovery.")
        if self._connected and _to_bool_connected(self._client.is_connected):
            return
        preflight = self.preflight()
        if not preflight.ok:
            raise QueueAutomationError("; ".join(preflight.blockers))


def _holder_identity(hole: int) -> str:
    return f"holder-{int(hole):03d}" if int(hole) > 0 else "holder"


def _default_holder_store(config: AppConfig) -> "HolderStateStore":
    from .holder_state import HolderStateStore

    data_dir = (config.general.data_dir or "").strip()
    path = Path(data_dir) / "holder_correction.json" if data_dir else None
    return HolderStateStore(path, allow_simulated=bool(config.general.nocomm))


def config_fingerprint(config: AppConfig) -> str:
    """Stable short hash of the active configuration for audit records."""

    try:
        payload = json.dumps(dataclasses.asdict(config), sort_keys=True, default=str)
    except Exception:
        payload = repr(config)
    return hashlib.sha256(payload.encode("utf-8")).hexdigest()[:16]


def build_measurement_backend(
    config: AppConfig,
    *,
    susceptibility_backend=None,
) -> MeasurementAutomationBackend:
    """Construct the active measurement backend from configuration.

    ``NO_COMM`` mode returns the labelled simulator. Hardware mode returns the
    queue backend, which records construction failures as preflight blockers
    instead of substituting a simulator.
    """
    if config.general.nocomm:
        return NoCommBackend()
    return QueueHardwareBackend(
        config,
        susceptibility_backend=susceptibility_backend,
    )
