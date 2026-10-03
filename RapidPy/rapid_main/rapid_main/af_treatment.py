"""Source-backed AF calibration and verified multi-pass treatment lifecycle.

References: RockmagStep.PerformStep, frmAF_2G.FindXCalibValue,
frmADWIN_AF.UpdateFields/GetUpSlope/GetDownSlope/ExecuteRamp.
All physical interfaces are injected; this module never opens hardware.
"""
from __future__ import annotations

from dataclasses import asdict, dataclass
from datetime import datetime, timezone
import math
import json
import os
from pathlib import Path
import re
import time
import uuid
from typing import Callable

from rapid_main.config import AfDemagConfig


def _positive(name: str, value: float) -> float:
    number = float(value)
    if not math.isfinite(number) or number <= 0:
        raise ValueError(f"AF {name} must be finite and positive.")
    return number


@dataclass(frozen=True)
class CalibratedAfRamp:
    label: str
    coil: str
    field_mT: float
    monitor_peak_v: float
    ramp_peak_v: float
    frequency_hz: float
    slope_up_vps: float
    slope_down_vps: float
    hold_ms: int
    io_rate_hz: float
    calibration_source: str
    ramp_mode: int = 2  # legacy CalRamp=True; mode 3 is CLIPTEST


def plan_calibrated_af_ramp(label: str, field_mT: float, coil: str, cfg: AfDemagConfig) -> CalibratedAfRamp:
    if coil not in {"axial", "transverse"}:
        raise ValueError("AF coil must be axial or transverse.")
    if not cfg.enabled or cfg.system.upper() != "ADWIN":
        raise ValueError("An enabled ADWIN AF system is required.")
    if not cfg.calibration_source.strip() or not getattr(cfg, f"{coil}_calibrated"):
        raise ValueError(f"Accepted {coil} AF calibration is required.")
    bits = (cfg.axial_relay_bit, cfg.transverse_relay_bit)
    if any(not isinstance(bit, int) or isinstance(bit, bool) or not 0 <= bit <= 5 for bit in bits) or bits[0] == bits[1]:
        raise ValueError("AF coil relay bits must be distinct ADwin-light-16 outputs in 0..5.")
    requested = float(field_mT)
    maximum = _positive(f"{coil} maximum field", getattr(cfg, f"{coil}_max_mT"))
    minimum = float(getattr(cfg, f"{coil}_min_mT"))
    if not math.isfinite(minimum) or not 0 <= minimum <= maximum:
        raise ValueError(f"Invalid {coil} AF field limits.")
    if not math.isfinite(requested) or requested < 0 or requested > maximum or (requested != 0 and requested < minimum):
        raise ValueError(f"AF field {requested:g} mT is outside the {coil} calibration limits {minimum:g}..{maximum:g} mT.")
    try:
        points = [(float(p[0]), float(p[1])) for p in getattr(cfg, f"{coil}_calibration") if len(p) == 2]
        if len(points) != len(getattr(cfg, f"{coil}_calibration")):
            raise ValueError("Each calibration point requires voltage and field.")
    except (TypeError, ValueError, IndexError) as exc:
        raise ValueError(f"Invalid {coil} AF calibration table.") from exc
    if len(points) < 2 or any(not math.isfinite(x) or not math.isfinite(y) or x < 0 or y <= 0 for x, y in points):
        raise ValueError(f"Invalid {coil} AF calibration table.")
    if any(b[1] <= a[1] or b[0] < a[0] for a, b in zip(points, points[1:])):
        raise ValueError(f"{coil} AF calibration fields must increase and voltages must not decrease.")
    # VB6 uses its zero-initialized element 0 for the first interval, and the
    # final two points for extrapolation bounded by the configured field max.
    points = [(0.0, 0.0), *points]
    low, high = points[-2:]
    for low_candidate, high_candidate in zip(points, points[1:]):
        if requested <= high_candidate[1]:
            low, high = low_candidate, high_candidate
            break
    monitor = low[0] + (requested - low[1]) * (high[0] - low[0]) / (high[1] - low[1])
    monitor_max = _positive(f"{coil} monitor limit", getattr(cfg, f"{coil}_monitor_max_v"))
    ramp_max = _positive(f"{coil} ramp limit", getattr(cfg, f"{coil}_ramp_max_v"))
    if monitor_max > 10 or ramp_max > 10 or not 0 <= monitor <= monitor_max:
        raise ValueError("AF calibration exceeds ADwin monitor/ramp voltage limits.")
    # The /2 transverse factor is explicitly present in UpdateFields.
    ramp = monitor / monitor_max * ramp_max / (2 if coil == "transverse" else 1)
    frequency = _positive(f"{coil} resonance frequency", getattr(cfg, f"{coil}_frequency_hz"))
    rate = _positive("IO rate", cfg.io_rate_hz)
    if rate > 50000:
        raise ValueError("AF IO rate exceeds the shipped sineout process's 50 kHz limit.")
    if rate / frequency < 4:
        raise ValueError("AF IO rate must provide at least four samples per period.")
    min_ms = _positive("minimum ramp-up duration", cfg.ramp_up_min_ms)
    max_ms = _positive("maximum ramp-up duration", cfg.ramp_up_max_ms)
    if max_ms < min_ms:
        raise ValueError("AF ramp-up duration bounds are reversed.")
    up_duration_ms = round(ramp / _positive(f"{coil} ramp-up rate", getattr(cfg, f"{coil}_ramp_up_vps")) * 1000)
    up_duration_ms = min(max_ms, max(min_ms, up_duration_ms))
    min_periods = _positive("minimum ramp-down periods", cfg.ramp_down_min_periods)
    max_periods = _positive("maximum ramp-down periods", cfg.ramp_down_max_periods)
    if max_periods < min_periods:
        raise ValueError("AF ramp-down period bounds are reversed.")
    periods = min(max_periods, max(min_periods, round(_positive("ramp-down periods per volt", cfg.ramp_down_periods_per_v) * ramp)))
    hold_ms = max(100, int(_positive("peak hold periods", cfg.hold_peak_periods) / frequency * 1000))
    return CalibratedAfRamp(label, coil, requested, monitor, ramp, frequency,
                            ramp / (up_duration_ms / 1000), ramp / (periods / frequency),
                            hold_ms, rate, cfg.calibration_source)


@dataclass(frozen=True)
class AfPass:
    angle_deg: float
    ramp: CalibratedAfRamp
    pause_before_s: float = 0.0


@dataclass(frozen=True)
class AfTreatmentPlan:
    label: str
    target_position: int
    passes: tuple[AfPass, ...]


def plan_af_treatment(label: str, cfg: AfDemagConfig, sample_height: int) -> AfTreatmentPlan:
    from .treatment_labels import parse_field_treatment
    request = parse_field_treatment(label)
    if request.family not in {"AFMAX", "AFZ", "AF"}:
        raise ValueError(f"Unsupported AF treatment label: {label!r}")
    if int(cfg.coil_position) == 0 or int(sample_height) <= 0:
        raise ValueError("AF coil position and specimen height must be configured.")
    target = math.floor(cfg.coil_position + int(sample_height) / 2)
    if target == 0 or (target > 0) != (cfg.coil_position > 0):
        raise ValueError("Specimen centre would cross the configured AF coil side.")
    family = request.family
    field = request.field_mT if request.field_mT is not None else float(cfg.peak)
    if family == "AFMAX":
        # AFmax uses independent coil maxima in the legacy PerformStep.
        passes = (
            AfPass(0, plan_calibrated_af_ramp(label, cfg.transverse_max_mT, "transverse", cfg)),
            AfPass(90, plan_calibrated_af_ramp(label, cfg.transverse_max_mT, "transverse", cfg)),
            AfPass(360, plan_calibrated_af_ramp(label, cfg.axial_max_mT, "axial", cfg)),
        )
    else:
        passes = (AfPass(0, plan_calibrated_af_ramp(label, field, "axial", cfg)),)
        if family == "AF":
            passes += (
                AfPass(0, plan_calibrated_af_ramp(label, field, "transverse", cfg), 3.0),
                AfPass(90, plan_calibrated_af_ramp(label, field, "transverse", cfg)),
            )
    return AfTreatmentPlan(label, target, passes)


@dataclass(frozen=True)
class AfTreatmentPhase:
    name: str
    ok: bool
    timestamp_iso: str
    detail: str = ""


@dataclass(frozen=True)
class AfTreatmentRecord:
    treatment_id: str
    sample_id: str
    run_id: str
    plan: AfTreatmentPlan
    completed_passes: int
    phases: tuple[AfTreatmentPhase, ...]
    error: str
    safe_return_errors: tuple[str, ...]
    safe_state_confirmed: bool
    simulated: bool
    schema: str = "rapidpy.af.treatment.v1"
    bias_mT: float | None = None

    def to_dict(self):
        return asdict(self)


class AfTreatmentError(RuntimeError):
    def __init__(self, record: AfTreatmentRecord):
        super().__init__("; ".join([record.error, *record.safe_return_errors]).strip("; "))
        self.record = record


def write_af_treatment_record(path: str | Path, record: AfTreatmentRecord) -> Path:
    target = Path(path)
    target.parent.mkdir(parents=True, exist_ok=True)
    temporary = target.with_name(target.name + ".tmp-" + uuid.uuid4().hex)
    try:
        with temporary.open("x", encoding="utf-8") as handle:
            json.dump(record.to_dict(), handle, indent=2, sort_keys=True, allow_nan=False)
            handle.write("\n")
            handle.flush()
            os.fsync(handle.fileno())
        # Hard-link publication is atomic and cannot overwrite an existing
        # record. A filesystem without this support fails closed.
        os.link(temporary, target)
    finally:
        if temporary.exists():
            temporary.unlink()
    return target


class AfTreatmentService:
    def __init__(self, adapter, vertical, turning, *, should_cancel: Callable[[], bool] | None = None,
                 sleep: Callable[[float], None] = time.sleep, clock=None, bias_adapter=None, bias_mT=None):
        self.adapter, self.vertical, self.turning = adapter, vertical, turning
        self.should_cancel = should_cancel or (lambda: False)
        set_halt_check = getattr(adapter, "set_halt_check", None)
        if callable(set_halt_check):
            set_halt_check(self.should_cancel)
        self.sleep = sleep
        self.clock = clock or (lambda: datetime.now(timezone.utc))
        self.bias_adapter, self.bias_mT = bias_adapter, bias_mT
        if bias_adapter is not None:
            bias_adapter.should_cancel = self.should_cancel
            bias_adapter.sleep = self.sleep

    def execute(self, plan: AfTreatmentPlan, *, sample_id: str, run_id: str) -> AfTreatmentRecord:
        phases, safe_errors = [], []
        completed, error, motion_started = 0, "", False
        def phase(name, operation, *, motion=False):
            try:
                result = operation()
                if motion and (not result.ok or not math.isfinite(float(result.actual))):
                    raise RuntimeError(result.detail or f"Unverified motion: target={result.target}, actual={result.actual}")
                phases.append(AfTreatmentPhase(name, True, self.clock().isoformat()))
                return result
            except Exception as exc:
                phases.append(AfTreatmentPhase(name, False, self.clock().isoformat(), str(exc)))
                raise
        def cancelled():
            if self.should_cancel():
                raise InterruptedError("AF treatment cancelled.")
        try:
            if not sample_id.strip() or not run_id.strip():
                raise ValueError("AF treatment requires sample and run identities.")
            if not self.adapter.is_connected():
                raise RuntimeError("AF adapter is disconnected.")
            cancelled()
            motion_started = True
            if self.bias_adapter is not None:
                phase("bias_set", lambda: self.bias_adapter.set_bias_mT(self.bias_mT))
            phase("home_before", self.vertical.home_to_top, motion=True)
            cancelled()
            phase("move_to_coil", lambda: self.vertical.move_to(plan.target_position, speed_index=1), motion=True)
            for index, treatment_pass in enumerate(plan.passes):
                cancelled()
                phase(f"rotate_{index}", lambda p=treatment_pass: self.turning.rotate_to(p.angle_deg), motion=True)
                # Cooperative pause, rather than an uncancellable three-second sleep.
                remaining = treatment_pass.pause_before_s
                while remaining > 0:
                    cancelled()
                    interval = min(0.1, remaining)
                    self.sleep(interval)
                    remaining -= interval
                cancelled()
                if treatment_pass.ramp.field_mT > 0:
                    phase(f"ramp_{index}", lambda p=treatment_pass: self.adapter.apply_calibrated_af(p.ramp))
                completed += 1
            cancelled()
            phase("rotate_reference", lambda: self.turning.rotate_to(360), motion=True)
        except Exception as exc:
            error = f"{type(exc).__name__}: {exc}"
        finally:
            if motion_started:
                # Always attempt both cleanup operations; a relay error must not
                # conceal a failed home or skip the attempt to clear the coil.
                bias_cleanup = (("bias_clear", self.bias_adapter.clear_bias, False),) if self.bias_adapter is not None else ()
                for name, action, motion in (*bias_cleanup,
                    ("field_reset", self.adapter.reset_field, False),
                    ("home_after", self.vertical.home_to_top, True),
                ):
                    try:
                        phase(name, action, motion=motion)
                    except Exception as exc:
                        safe_errors.append(f"{name}: {exc}")
        record = AfTreatmentRecord("af-" + uuid.uuid4().hex, sample_id, run_id, plan, completed,
                                   tuple(phases), error, tuple(safe_errors),
                                   motion_started and not safe_errors, bool(getattr(self.adapter, "simulated", False)), bias_mT=self.bias_mT)
        if error or safe_errors:
            raise AfTreatmentError(record)
        return record
