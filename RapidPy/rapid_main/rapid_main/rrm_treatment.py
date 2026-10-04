"""Rotational remanence: verified continuous turning during calibrated AF.

VB6 RockmagStep uses transverse RRM / axial RRMz and clamps spin to +/-40
rps. RapidPy rejects requests outside those bounds instead of changing them.
Labels retain both parameters: RRM100/5 (100 mT at 5 rps).
"""
from dataclasses import dataclass
from datetime import datetime, timezone
import math
import re
import time
import uuid

from .af_treatment import (AfTreatmentRecord, AfTreatmentPhase, AfTreatmentError,
                           CalibratedAfRamp, plan_calibrated_af_ramp)


@dataclass(frozen=True)
class RrmTreatmentPlan:
    label: str
    target_position: int
    ramp: CalibratedAfRamp
    speed_rps: float
    duration_s: float
    motor_calibration_source: str
    full_rotation: int
    native_1rps: int
    bias_mT: float | None = None


def plan_rrm_treatment(label, config, *, sample_height=None):
    match = re.fullmatch(r"(RRMZ|RRM)\s*(\d+(?:\.\d+)?)\s*/\s*(-?\d+(?:\.\d+)?)(?:\s*RPS)?(?:\s*@\s*(\d+(?:\.\d+)?))?", label.strip().upper())
    if not match:
        raise ValueError("RRM requires field and signed spin speed: RRM100/5 or RRMZ100/-5 (mT/rps).")
    field, speed = float(match[2]), float(match[3])
    bias = float(match[4]) if match[4] is not None else None
    if bias is not None:
        from .arm_bias import plan_arm_bias
        plan_arm_bias(bias, config.irm_arm)
    if field <= 0 or not 0 < abs(speed) <= 40:
        raise ValueError("RRM requires a positive AF field and nonzero spin within +/-40 rps.")
    from .hardware_contracts import _build_motor_controller_config
    station = config.motor_station
    if not station.calibration_source or set(station.ports) != {"changer_x", "changer_y", "turning", "updown"}:
        raise ValueError("RRM requires imported native motor calibration and all axis ports.")
    motor = _build_motor_controller_config(config)
    from rapidpy_common.hardware import MotorAxisConfig
    from rapidpy_common.motor_routing import RoutedMotorSerialClient
    # Constructor validation is read-only, including duplicate port/address checks.
    RoutedMotorSerialClient(motor, [MotorAxisConfig(axis, index, station.addresses.get(axis, 0), port)
                                   for index, (axis, port) in enumerate(station.ports.items(), 1)])
    cfg = config.af_demag
    height = config.motion.sample_height if sample_height is None else sample_height
    if isinstance(height, bool) or not isinstance(height, (int, float)) or not math.isfinite(height) or int(height) != height:
        raise ValueError('RRM specimen height must be an integer count.')
    if cfg.coil_position == 0 or height <= 0:
        raise ValueError("RRM coil position and specimen height must be configured.")
    target = math.floor(cfg.coil_position + height / 2)
    if target == 0 or (target > 0) != (cfg.coil_position > 0):
        raise ValueError("RRM specimen centre crosses the configured coil side.")
    ramp = plan_calibrated_af_ramp(label, field, "axial" if match[1] == "RRMZ" else "transverse", cfg)
    ramp_seconds = ramp.ramp_peak_v / ramp.slope_up_vps + ramp.hold_ms / 1000 + ramp.ramp_peak_v / ramp.slope_down_vps
    duration = max(300., ramp_seconds + 30.)
    velocity = abs(motor.turning_motor_1rps * speed)
    travel = abs(motor.turning_motor_full_rotation * speed * duration)
    if velocity > 2**31 - 1 or travel > 2**31 - 1:
        raise ValueError("RRM spin exceeds signed 32-bit native controller limits.")
    return RrmTreatmentPlan(label, target, ramp, speed, duration, station.calibration_source,
                            motor.turning_motor_full_rotation, motor.turning_motor_1rps, bias)


class RrmTreatmentService:
    def __init__(self, adapter, vertical, turning, *, should_cancel=None, sleep=time.sleep,
                 monotonic=time.monotonic, clock=None, bias_adapter=None):
        self.adapter, self.vertical, self.turning = adapter, vertical, turning
        self.should_cancel = should_cancel or (lambda: False)
        self.sleep, self.monotonic = sleep, monotonic
        self.clock = clock or (lambda: datetime.now(timezone.utc))
        self.bias_adapter = bias_adapter
        if bias_adapter is not None:
            bias_adapter.should_cancel = self.should_cancel
            bias_adapter.sleep = sleep

    def execute(self, plan, *, sample_id, run_id):
        phases, errors = [], []
        error, started, spinning, completed = "", False, False, 0
        previous_position, previous_time = None, 0

        def phase(name, action, *, motion=False, native=False):
            try:
                value = action()
                if motion and (not value.ok or not math.isfinite(value.actual)):
                    raise RuntimeError(value.detail or "Motion was not verified.")
                if native and not value.success:
                    raise RuntimeError("Turning controller rejected the spin operation.")
                phases.append(AfTreatmentPhase(name, True, self.clock().isoformat()))
                return value
            except Exception as exc:
                phases.append(AfTreatmentPhase(name, False, self.clock().isoformat(), str(exc)))
                raise

        def check():
            nonlocal previous_position, previous_time
            if self.should_cancel():
                return True
            now = self.monotonic()
            if spinning and now - previous_time >= .25:
                position = self.turning.position()
                delta = position - previous_position
                expected_sign = -plan.full_rotation * plan.speed_rps
                if delta == 0 or delta * expected_sign <= 0:
                    raise RuntimeError("RRM rotation stalled or reversed during AF.")
                phases.append(AfTreatmentPhase("rotation_readback", True, self.clock().isoformat(),
                                               f"position={position}; delta={delta}; elapsed_s={now - previous_time:g}"))
                previous_position, previous_time = position, now
            return False

        try:
            if not sample_id.strip() or not run_id.strip():
                raise ValueError("RRM requires sample and run identities.")
            if not self.adapter.is_connected() or getattr(self.adapter, "simulated", False):
                raise ValueError("RRM requires a connected physical AF adapter.")
            if plan.bias_mT is not None and self.bias_adapter is None:
                raise ValueError("RRM bias requires a physical ARM bias adapter.")
            if check(): raise InterruptedError("RRM cancelled.")
            started = True
            if plan.bias_mT is not None:
                phase("bias_set", lambda: self.bias_adapter.set_bias_mT(plan.bias_mT))
            phase("home_before", self.vertical.home_to_top, motion=True)
            if check(): raise InterruptedError("RRM cancelled.")
            phase("move_to_coil", lambda: self.vertical.move_to(plan.target_position, speed_index=1), motion=True)
            if check(): raise InterruptedError("RRM cancelled.")
            phase("orient_before_spin", lambda: self.turning.rotate_to(0), motion=True)
            if check(): raise InterruptedError("RRM cancelled.")
            previous_position, previous_time = self.turning.position(), self.monotonic()
            # Mark intent before transport: an ACK failure can still leave motion.
            spinning = True
            phase("spin_start", lambda: self.turning.spin(plan.speed_rps, plan.duration_s), native=True)
            self.sleep(.25)
            if check(): raise InterruptedError("RRM cancelled.")
            self.adapter.set_halt_check(check)
            phase("ramp_while_spinning", lambda: self.adapter.apply_calibrated_af(plan.ramp))
            if check(): raise InterruptedError("RRM cancelled.")
            completed = 1
        except Exception as exc:
            error = f"{type(exc).__name__}: {exc}"
        finally:
            if started:
                bias_cleanup = (("bias_clear", self.bias_adapter.clear_bias, False),) if self.bias_adapter is not None else ()
                for name, action, native in (*bias_cleanup, ("field_reset", self.adapter.reset_field, False),
                                            ("spin_stop_verified", self.turning.stop_spin, True)):
                    try: phase(name, action, native=native)
                    except Exception as exc: errors.append(f"{name}: {exc}")
                spinning = False
                if not errors:
                    try:
                        phase("restore_reference", self.turning.restore_spin_reference, motion=True)
                        phase("home_after", self.vertical.home_to_top, motion=True)
                    except Exception as exc: errors.append(str(exc))
                else:
                    phases.append(AfTreatmentPhase("lift_return_withheld", False, self.clock().isoformat(), "Field or stationary rotation is unverified."))
            self.adapter.set_halt_check(self.should_cancel)
        record = AfTreatmentRecord("af-" + uuid.uuid4().hex, sample_id, run_id, plan, completed,
                                   tuple(phases), error, tuple(errors), started and not errors, False,
                                   schema="rapidpy.rrm.treatment.v1", bias_mT=plan.bias_mT)
        if error or errors: raise AfTreatmentError(record)
        return record

    def recover(self, plan, *, sample_id, run_id):
        """Independently clear field and stop turning before any lift motion."""
        phases, errors = [], []
        bias_cleanup = (("bias_clear", self.bias_adapter.clear_bias),) if self.bias_adapter is not None else ()
        for name, action in (*bias_cleanup, ("field_reset", self.adapter.reset_field), ("spin_stop_verified", self.turning.stop_spin)):
            try:
                value = action()
                if name == "spin_stop_verified" and not value.success:
                    raise RuntimeError("Stationary turning was not verified.")
                phases.append(AfTreatmentPhase(name, True, self.clock().isoformat()))
            except Exception as exc:
                phases.append(AfTreatmentPhase(name, False, self.clock().isoformat(), str(exc)))
                errors.append(f"{name}: {exc}")
        if not errors:
            for name, action in (("restore_reference", self.turning.restore_spin_reference), ("home_after", self.vertical.home_to_top)):
                try:
                    value = action()
                    if not value.ok or not math.isfinite(value.actual): raise RuntimeError(value.detail or "Motion was not verified.")
                    phases.append(AfTreatmentPhase(name, True, self.clock().isoformat()))
                except Exception as exc:
                    phases.append(AfTreatmentPhase(name, False, self.clock().isoformat(), str(exc)))
                    errors.append(f"{name}: {exc}")
                    break
        record = AfTreatmentRecord("af-" + uuid.uuid4().hex, sample_id, run_id, plan, 0, tuple(phases), "",
                                   tuple(errors), not errors, False, schema="rapidpy.rrm.recovery.v1", bias_mT=plan.bias_mT)
        if errors: raise AfTreatmentError(record)
        return record
