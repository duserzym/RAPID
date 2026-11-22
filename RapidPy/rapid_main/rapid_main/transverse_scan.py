"""Native and simulated execution of the transverse probe scan."""
from __future__ import annotations

from dataclasses import asdict, dataclass
from datetime import datetime, timezone
import json
import math
import os
import threading
import time
import uuid

from rapidpy_common.af_diagnostic_safety import publish_diagnostic_record, station_profile
from rapidpy_common.gaussmeter import GaussmeterClient
from rapidpy_common.hardware import HardwareError, convert_pos_to_angle
from rapidpy_common.hardware_safety import HardwareSafetyError
from rapidpy_common.motor_diagnostic_safety import verify_stopped_in_place

from .af_treatment import plan_calibrated_af_ramp
from .diagnostic_services import AfDemagBackendAdapter, DCMotorBackendAdapter
from .transverse import TransverseSample, TransverseScanPlan, best_transverse_sample


HELPER = "rapid_main_transverse_probe"
ANGLE_TOLERANCE_DEG = .35
_FIELD_UNIT_TO_MT = {"T": 1000., "mT": 1., "kG": 100., "G": .1}


@dataclass(frozen=True)
class TransverseScanResult:
    plan: TransverseScanPlan
    samples: tuple[TransverseSample, ...]
    best: TransverseSample
    record_id: str
    evidence_path: str


@dataclass(frozen=True)
class TransverseEvidence:
    treatment_id: str
    operation: dict
    station_profile: dict
    timestamp_iso: str
    observations: tuple
    error: str
    cleanup_error: str
    safe_state_confirmed: bool
    outputs_held: bool = False
    simulated: bool = False
    schema: str = "rapidpy.transverse.probe_evidence.v1"

    def to_dict(self):
        return asdict(self)


def _check_cancel(cancelled):
    if cancelled():
        raise HardwareError("Transverse probe operation stopped.")


class TransverseScanPause:
    def __init__(self):
        self._resume = threading.Event()
        self._resume.set()

    @property
    def requested(self):
        return not self._resume.is_set()

    def request(self):
        self._resume.clear()

    def resume(self):
        self._resume.set()

    def wait_between_points(self, cancelled, progress, *, verify=None):
        _check_cancel(cancelled)
        if not self.requested:
            return
        if verify is not None:
            verify()
        _check_cancel(cancelled)
        progress("Paused between readings." if verify is None else "Paused. AF outputs are off and both motors are stopped.")
        while not self._resume.wait(.05):
            _check_cancel(cancelled)
        _check_cancel(cancelled)
        if verify is not None:
            verify()


def _append_reading(handle, record):
    handle.write(json.dumps(record, allow_nan=False) + "\n")
    handle.flush()
    os.fsync(handle.fileno())


def _wait(seconds, cancelled):
    deadline = time.monotonic() + seconds
    while (remaining := deadline - time.monotonic()) > 0:
        _check_cancel(cancelled)
        time.sleep(min(remaining, .05))
    _check_cancel(cancelled)


def _read_peak(meter):
    reading = meter.read(wait_for_new_data=True, sample_timeout_s=2.)
    factor = _FIELD_UNIT_TO_MT.get(reading.units_label)
    if (reading.mode_index != 1 or factor is None or not 0 <= reading.range_index <= 3
            or not math.isfinite(reading.raw_value) or not math.isfinite(reading.value * factor)):
        raise HardwareError("The 908A must provide a fresh, valid DC Peak reading in Tesla or Gauss.")
    return {**asdict(reading), "timestamp": reading.timestamp.isoformat() if reading.timestamp else None,
            "peak_mT": reading.value * factor}


class NativeTransverseScan:
    def __init__(self, config, motors, af, *, meter_factory=GaussmeterClient, wait=_wait):
        if not isinstance(motors, DCMotorBackendAdapter) or not isinstance(af, AfDemagBackendAdapter):
            raise HardwareError("The native DC motors and ADwin AF adapters are required.")
        if motors._station_baud is None or not af.is_connected():
            raise HardwareError("Connect the accepted motor station and ADwin AF configuration.")
        scale = config.motor_station.lift_counts_per_cm
        height = config.af_demag.coil_position
        if scale is None or not math.isfinite(scale) or scale <= 0 or height >= 0:
            raise ValueError("Import the calibrated AF lift position and lift counts per centimetre.")
        self.config = config
        self.motors = motors
        self.af = af
        self.store = motors._safety_store
        self.client = motors._client
        self.lift = motors._axes["Up/Down"]
        self.turning = motors._axes["Turning"]
        self.rotation_counts = self.client.config.turning_motor_full_rotation
        if not 1000 <= abs(self.rotation_counts) < 2**31:
            raise ValueError("The accepted turning encoder scale does not support this scan.")
        self.height = height
        self.scale = scale
        self.meter_factory = meter_factory
        self.wait = wait

    def _profile(self):
        return {"helper": HELPER, "resources": {
            "motors": self.motors._safety_profile(port=self.turning.port, baud=self.motors._station_baud),
            "af": station_profile(self.af._controller),
        }, "af_height": self.height, "lift_counts_per_cm": self.scale}

    def pending_recovery(self):
        pending = self.store.pending()
        return bool(pending and pending["family"] == "station_diagnostic"
                    and pending["profile"].get("helper") == HELPER)

    def _field_off(self, observations):
        self.af.recover_field()
        proof = self.af._controller.last_safe_field_recovery
        if (proof.get("safe_state_confirmed") is not True or proof.get("outputs_zero_confirmed") is not True
                or proof.get("simulated") is not False):
            raise HardwareSafetyError("ADwin field-off readback is unverified.")
        observations.append({"af_outputs_off": proof})
        return proof

    def _stationary(self, observations, *, require_height=True, ready=None):
        positions, failures, axis_samples = {}, [], {}
        for name, axis in (("lift", self.lift), ("turning", self.turning)):
            if ready is not None and axis.port.upper() not in ready:
                failures.append(f"{name} port unavailable")
                continue
            samples, error = verify_stopped_in_place(self.client, axis)
            axis_samples[name] = samples
            if error:
                failures.append(f"{name}: {error}")
            else:
                positions[name] = samples[-1]["position_raw"]
        observations.append({"stationary_axes": axis_samples})
        if failures:
            raise HardwareSafetyError("Stopped motor readback failed. " + ". ".join(failures))
        if require_height and abs(positions["lift"] - self.height) > self.scale:
            raise HardwareError("Place the manually mounted probe at the calibrated AF height (within 1 cm).")
        return positions["lift"], convert_pos_to_angle(positions["turning"], self.rotation_counts)

    @staticmethod
    def _check_travel(plan, angle):
        low, high = plan.bounds
        if not low - ANGLE_TOLERANCE_DEG <= angle <= high + ANGLE_TOLERANCE_DEG:
            raise HardwareError("Turning readback is outside the operator-cleared encoder interval.")

    def _move(self, plan, target, lift_position, observations, cancelled):
        _check_cancel(cancelled)
        self._check_travel(plan, target)
        self.client._motion_cancel_check = cancelled
        try:
            movement = self.client.turning_motor_rotate(self.turning, target)
        finally:
            self.client._motion_cancel_check = None
        current_lift, actual = self._stationary(observations)
        if current_lift != lift_position:
            raise HardwareError("Lift position changed during turning. Verify the probe height.")
        self._check_travel(plan, actual)
        if movement.success is not True or abs(actual - target) > ANGLE_TOLERANCE_DEG:
            raise HardwareError(f"Turning readback {actual:.3f}° missed {target:.3f}°.")
        return actual

    def _require_original_scan(self, result, profile):
        current = self.store.read() or {}
        record = current.get("record") or {}
        expected_plan = {**asdict(result.plan), "angles_deg": list(result.plan.angles_deg)}
        if (current.get("status") != "verified" or current.get("profile") != profile
                or record.get("treatment_id") != result.record_id
                or record.get("error") or record.get("operation", {}).get("plan") != expected_plan
                or {"best_measured_sample": asdict(result.best)} not in record.get("observations", ())):
            raise HardwareSafetyError("The verified scan no longer matches this result. Scan again before applying an angle.")
        peaks = [item for item in record["observations"] if "dc_peak" in item]
        if not peaks or not result.samples or peaks[-1]["actual_deg"] != result.samples[-1].actual_deg:
            raise HardwareSafetyError("The final encoder position does not match the verified scan.")

    def _readings_path(self, token):
        return self.store.path.parent / "hardware_diagnostics" / f"transverse-{token}.jsonl"

    def _transaction(self, operation, *, run=None, previous_scan=None, recovery=False):
        with self.store.operation_lease():
            profile, ready = self._profile(), None
            if recovery:
                pending = self.store.pending(profile)
                if not pending or pending["family"] != "station_diagnostic":
                    raise HardwareSafetyError("There is no matching transverse scan to recover.")
                token, operation = pending["token"], pending["plan"]
            else:
                if previous_scan is not None:
                    self._require_original_scan(previous_scan, profile)
                token = self.store.begin("station_diagnostic", operation, profile)
            observations, failures = [], []
            error = ""
            opened_here = not self.motors.is_connected()
            result = None
            try:
                if not recovery:
                    if opened_here:
                        self.client.connect(baudrate=self.motors._station_baud)
                    self._field_off(observations)
                    result = run(observations, token)
            except BaseException as exc:
                error = f"{type(exc).__name__}: {exc}"
                raise
            finally:
                try:
                    self._field_off(observations)
                except Exception as exc:
                    failures.append(f"AF outputs: {exc}")
                try:
                    if opened_here and (recovery or not self.client.is_connected):
                        handles = {port: (client, client._serial) for port, client in self.client._connections.items()}
                        ready, errors = self.client.connect_for_stop(
                            HELPER, baudrate=self.motors._station_baud, original_handles=handles,
                        )
                        observations.append({"stop_connection_errors": errors})
                    self._stationary(observations, require_height=False, ready=ready)
                except Exception as exc:
                    failures.append(f"Motor stop: {exc}")
                finally:
                    if opened_here:
                        try:
                            self.client.disconnect()
                        except Exception as exc:
                            failures.append(f"Motor connection release: {exc}")
                if operation.get("action") == "transverse_probe_scan":
                    path = self._readings_path(token)
                    if path.is_file():
                        observations.append({"readings_checkpoint": str(path)})
                record = TransverseEvidence(
                    "station-" + uuid.uuid4().hex, operation, profile, datetime.now(timezone.utc).isoformat(),
                    tuple(observations), error, ". ".join(failures), not failures,
                )
                try:
                    folder = publish_diagnostic_record(self.store, record, family="station")
                    self.store.finish(token, profile, record)
                except Exception as exc:
                    raise HardwareSafetyError(f"Transverse scan recovery evidence is unverified. {exc}") from exc
                if failures:
                    raise HardwareSafetyError("Transverse station recovery required. " + record.cleanup_error)
        return result, record.treatment_id, str(folder / "record.json")

    def scan(self, plan, *, cancelled, progress, pause=None, sampled=None):
        ramp = plan_calibrated_af_ramp("Transverse probe scan", plan.field_mT, "transverse", self.config.af_demag)
        if ramp.monitor_peak_v <= 0 or ramp.ramp_peak_v <= 0:
            raise ValueError("The requested field does not produce a nonzero calibrated ramp.")
        reserved_ports = set(self.config.motor_station.ports.values())
        reserved_ports.update(getattr(device, "port", "") for device in (
            self.config.squid, self.config.susceptibility, self.config.vacuum,
        ))
        if f"COM{plan.meter_port}" in {str(port).strip().upper() for port in reserved_ports}:
            raise ValueError("Select a 908A port distinct from other configured station devices.")
        meter = self.meter_factory(port=plan.meter_port, mode=1)

        def run(observations, token):
            lift_position, actual = self._stationary(observations)
            self._check_travel(plan, actual)
            _check_cancel(cancelled)
            meter.connect(timeout_s=5.)
            meter.set_mode(1)
            path = self._readings_path(token)
            path.parent.mkdir(parents=True, exist_ok=True)
            samples = []

            def verify_pause_position():
                self._field_off(observations)
                current_lift, current = self._stationary(observations)
                self._check_travel(plan, current)
                if current_lift != lift_position:
                    raise HardwareError("Lift position changed at the pause boundary. Start a new scan.")
                if abs(current - actual) > ANGLE_TOLERANCE_DEG:
                    raise HardwareError("Turning position changed at the pause boundary. Start a new scan.")

            with path.open("x", encoding="utf-8", newline="\n") as readings:
                _append_reading(readings, {
                    "schema": "rapidpy.transverse.readings.v1", "operation_token": token,
                    "plan": asdict(plan), "ramp": asdict(ramp), "simulated": False,
                })
                for index, angle in enumerate(plan.angles_deg, 1):
                    if pause is not None:
                        pause.wait_between_points(cancelled, progress, verify=verify_pause_position)
                    progress(f"Measuring position {index} of {len(plan.angles_deg)} at {angle:g}°")
                    actual = self._move(plan, angle, lift_position, observations, cancelled)
                    meter.reset_peak()
                    self.wait(1., cancelled)
                    baseline = _read_peak(meter)
                    observations.append({"requested_deg": angle, "actual_deg": actual, "baseline": baseline})
                    self.af.set_halt_check(cancelled)
                    try:
                        _check_cancel(cancelled)
                        self.af.apply_calibrated_af(ramp)
                        ramp_result = asdict(self.af._last_result)
                        observations.append({"af_ramp": ramp_result})
                        self.wait(.5, cancelled)
                        peak = _read_peak(meter)
                        observations.append({"requested_deg": angle, "actual_deg": actual, "dc_peak": peak})
                        if peak["range_index"] != baseline["range_index"] or peak["units_index"] != baseline["units_index"]:
                            raise HardwareError("The 908A range or units changed during a peak acquisition.")
                    finally:
                        self.af.set_halt_check(None)
                        field_off = self._field_off(observations)
                    sample = TransverseSample(angle, actual, peak["peak_mT"])
                    _append_reading(readings, {
                        **asdict(sample), "baseline": baseline, "dc_peak": peak, "af_ramp": ramp_result,
                        "af_outputs_off": field_off,
                    })
                    samples.append(sample)
                    if sampled is not None:
                        sampled(sample, str(path))
            _check_cancel(cancelled)
            best = best_transverse_sample(tuple(samples))
            observations.append({"best_measured_sample": asdict(best)})
            return tuple(samples), best

        try:
            result, record_id, path = self._transaction(
                {"action": "transverse_probe_scan", "plan": asdict(plan), "ramp": asdict(ramp),
                 "operator": self.config.general.operator}, run=run,
            )
        finally:
            meter.disconnect()
        return TransverseScanResult(plan, *result, record_id, path)

    def apply(self, result, *, cancelled, progress):
        target = result.best.actual_deg

        def run(observations, _token):
            lift_position, current = self._stationary(observations)
            self._check_travel(result.plan, current)
            if abs(current - result.samples[-1].actual_deg) > ANGLE_TOLERANCE_DEG:
                raise HardwareError("The turning position changed after the scan. Scan again before applying an angle.")
            progress(f"Turning to the measured peak at {target:.3f}°")
            actual = self._move(result.plan, target, lift_position, observations, cancelled)
            observations.append({"requested_deg": target, "actual_deg": actual})
            return actual

        actual, _, _ = self._transaction(
            {"action": "transverse_probe_apply", "scan_record": result.record_id, "target_deg": target,
             "operator": self.config.general.operator}, run=run, previous_scan=result,
        )
        return actual

    def recover(self, *, cancelled, progress):
        progress("Verifying ADwin outputs off and both motors stopped in place")
        self._transaction(None, recovery=True)


class SimulatedTransverseScan:
    def pending_recovery(self):
        return False

    def scan(self, plan, *, cancelled, progress, pause=None, sampled=None):
        low, high = plan.bounds
        peak_angle = low + (high - low) * .4
        samples = []
        for index, angle in enumerate(plan.angles_deg, 1):
            if pause is not None:
                pause.wait_between_points(cancelled, progress)
            progress(f"Simulating position {index} of {len(plan.angles_deg)} at {angle:g}°")
            _wait(.04, cancelled)
            peak = -plan.field_mT * math.exp(-16 * ((angle - peak_angle) / (high - low))**2)
            sample = TransverseSample(angle, angle, peak)
            samples.append(sample)
            if sampled is not None:
                sampled(sample, "")
        return TransverseScanResult(plan, tuple(samples), best_transverse_sample(tuple(samples)), "", "")

    def apply(self, result, *, cancelled, progress):
        _check_cancel(cancelled)
        return result.best.actual_deg
