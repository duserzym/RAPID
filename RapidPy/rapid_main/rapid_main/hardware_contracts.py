"""Hardware service contracts used by rapid_main workflows.

This module defines the stable protocol used by the measurement engine and
serves as the core abstraction boundary for later real-backend adapters.
"""
from __future__ import annotations

from dataclasses import dataclass
import dataclasses
import hashlib
import json
import math
from pathlib import Path
import re
from typing import Protocol, TYPE_CHECKING

from rapid_main import software_version
from rapid_main.communication_log import CommunicationLogger
from rapid_main.config import AppConfig

if TYPE_CHECKING:  # pragma: no cover - typing only
    from rapid_main.holder_measurement import HolderMeasurementOutcome
    from rapid_main.holder_state import HolderStateStore
    from rapid_main.squid_transport import BracketedSquidBackend

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


_DEMAG_LABEL_RE = re.compile(r"^(AF(?:MAX|Z)?|IRM|ARM)\s*(\d+(?:\.\d+)?)?(?:_([0-9]+(?:\.[0-9]+)?))?$")
_THERMAL_LABEL_RE = re.compile(r"^(TT|TH|TEMP)\s*(\d+(?:\.\d+)?)$")
_MEASUREMENT_ONLY_LABEL_RE = re.compile(r"^(?:NRM(?:-[XYZ])?|REPEAT\d+)$")


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


def _is_measurement_only_label(label: str) -> bool:
    return bool(_MEASUREMENT_ONLY_LABEL_RE.fullmatch((label or "").strip().upper()))


class QueueHardwareBackend(MeasurementAutomationBackend):
    """Adapter that combines motor automation with SQUID measurement.

    Hardware mode fails closed. A component that cannot be constructed is
    recorded as a preflight blocker instead of being replaced by a simulator,
    and every measurement path raises rather than inventing a value.
    """

    preflight_timeout: float | None = 12.0
    step_timeout: float | None = 45.0
    # Bracketed acquisition is bounded by serial/motion adapters and performs
    # safe whole-block recovery. A worker-thread timeout cannot cancel physical
    # I/O and eight seconds is shorter than a normal six-latch acquisition.
    read_timeout: float | None = None
    susceptibility_timeout: float | None = 2.0
    return_timeout: float | None = 20.0

    def __init__(
        self,
        config: AppConfig,
        *,
        holder_store: "HolderStateStore | None" = None,
        clock=None,
    ) -> None:
        self._config = config
        # Injected only by tests; production uses the real wall/monotonic clock
        # so the VB6 ARC and settling delays are actually observed.
        self._acquisition_clock = clock
        self._motor_config = _build_motor_controller_config(config)
        self._motor_communication_logger = CommunicationLogger(
            "DC_MOTOR", port=str(config.changer.port), max_payload_chars=2048
        )
        self._client = MotorSerialClient(
            self._motor_config, trace=self._trace_motor_communication
        )
        self._axes = {
            "changer_x": MotorAxisConfig("ChangerX", 1, 1),
            "turning": MotorAxisConfig("Turning", 2, 2),
            "updown": MotorAxisConfig("UpDown", 3, 3),
            "changer_y": MotorAxisConfig("ChangerY", 4, 4),
        }
        self._backend_errors: list[str] = []
        nocomm = bool(config.general.nocomm)

        from .diagnostic_services import (
            build_af_demag_backend,
            build_irm_arm_backend,
            build_squid_backend,
        )

        self._measurement = self._build_component("SQUID", build_squid_backend, config.squid, nocomm=nocomm)
        self._irm_arm = self._build_component("IRM/ARM", build_irm_arm_backend, config.irm_arm, nocomm=nocomm)
        self._af_demag = self._build_component(
            "AF demagnetizer", build_af_demag_backend, config.af_demag, nocomm=nocomm
        )

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

    def _trace_motor_communication(
        self, direction: str, payload: str, detail: str
    ) -> None:
        logger = self._motor_communication_logger
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
    def last_holder_outcome(self) -> "HolderMeasurementOutcome | None":
        return self._last_holder_outcome

    @property
    def acquisition_error(self) -> str:
        return self._acquisition_error

    def communication_events(self):
        """Return immutable live SQUID and treatment traffic in time order."""

        events = []
        motor_logger = getattr(self, "_motor_communication_logger", None)
        if motor_logger is not None:
            events.extend(motor_logger.transcript.events)
        for source in (self._bracketed, self._af_demag, self._irm_arm):
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
        if self._config.general.nocomm:
            return self._require_measurement().read_susceptibility()
        raise HardwareError(
            "Live susceptibility acquisition is unavailable: the SQUID magnetic-moment "
            "transport is not a Bartington susceptibility bridge. Configure a typed bridge "
            "adapter with zero/measure, coil motion, and holder correction before running SUSC."
        )

    def set_demag_step(self, label: str) -> None:
        """Apply a demagnetization step label to the measurement backend.

        AF labels are routed through the AF demagnetizer backend; IRM/ARM
        labels are routed through the adapter-backed IRM/ARM backend.
        Explicit measurement-only labels use the measurement backend contract.
        """
        step_type, field_mT, bias_mT = _parse_demag_label(label)
        if step_type in {"AF", "AFMAX", "AFZ"}:
            from .diagnostic_services import plan_af_demag_command

            method = getattr(self._af_demag, "apply_af", None)
            if callable(method):
                method(plan_af_demag_command(label, self._config.af_demag))
                self._treatment_label = str(label)
                return
            raise HardwareError("AF demagnetizer adapter not available for AF treatment.")

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
                self._treatment_label = str(label)
                return
            raise HardwareError("IRM/ARM adapter not available for IRM treatment.")

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
            raise HardwareError(
                "SUSC has no production susceptibility bridge route. Configure and validate "
                "bridge zero/measure, coil motion, and holder correction before running it live."
            )

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
        apply_thermal = getattr(self._measurement, "apply_thermal", None)
        for label in labels:
            step_type, field_mT, _bias_mT = _parse_demag_label(label)
            if step_type in {"AF", "AFMAX", "AFZ", "ARM"}:
                continue
            if step_type == "IRM" and field_mT is not None:
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
                blockers.append(
                    "Susceptibility step SUSC cannot run in hardware mode: no production "
                    "Bartington bridge adapter with zero/measure, coil motion, and holder "
                    "correction is configured."
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

    # -- preflight ---------------------------------------------------------

    def preflight(self) -> PreflightResult:
        blockers: list[str] = list(self._backend_errors)
        warnings: list[str] = []
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

        if self._connected and _to_bool_connected(self._measurement.is_connected):
            self._ensure_bracketed()
            return PreflightResult(
                ok=not blockers, blockers=tuple(blockers), warnings=tuple(warnings + self._acquisition_warnings())
            )

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

        self._measure_holder(hole)

    def _measure_holder(self, hole: int) -> None:
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
        try:
            service = HolderMeasurementService(
                self._bracketed.read_squid,
                self._holder_store,
                recover=self._bracketed.recover_flux_count_discontinuity,
                averaging_cycles=max(1, int(self._config.squid.samples_per_pos or 1)),
                flux_discontinuity_retries=int(self._bracketed.flux_discontinuity_retries),
                clock=getattr(self._acquisition_clock, "now", None),
            )
            outcome = service.measure(holder_id=_holder_identity(hole), hole=hole)
        finally:
            self._measuring_holder = False
            self._sample_name = previous_sample
        self._last_holder_outcome = outcome
        if not outcome.installed:
            raise QueueAutomationError(outcome.rejection_reason)

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
        if not self._connected or not _to_bool_connected(self._client.is_connected):
            return
        errors: list[str] = []
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

    def _ensure_connected(self) -> None:
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


def build_measurement_backend(config: AppConfig) -> MeasurementAutomationBackend:
    """Construct the active measurement backend from configuration.

    ``NO_COMM`` mode returns the labelled simulator. Hardware mode returns the
    queue backend, which records construction failures as preflight blockers
    instead of substituting a simulator.
    """
    if config.general.nocomm:
        return NoCommBackend()
    return QueueHardwareBackend(config)
