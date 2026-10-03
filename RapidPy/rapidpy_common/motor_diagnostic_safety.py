"""Durable ownership and stop-in-place recovery for native motor diagnostics."""
from contextlib import contextmanager
import copy
from dataclasses import asdict, dataclass
from datetime import datetime, timezone
import time
import uuid

from .hardware_safety import HardwareSafetyError, HardwareSafetyStore, default_safety_path
from .af_diagnostic_safety import publish_diagnostic_record


@dataclass(frozen=True)
class MotorDiagnosticRecord:
    treatment_id: str
    operation: dict
    timestamp_iso: str
    observations: tuple
    cleanup_observations: tuple
    error: str
    cleanup_error: str
    safe_state_confirmed: bool
    simulated: bool = False
    schema: str = "rapidpy.motion.diagnostic.v1"

    def to_dict(self):
        return asdict(self)


def verify_stopped_in_place(motor, axis, *, sleep=time.sleep):
    """Stop without homing/zeroing; verify register 7 and stable register 1.

    Quicksilver register 7 contains two signed velocity words. A complete zero
    register proves both words zero; stationary position alone is insufficient.
    Preserve a failed acknowledgement even if a fallback halt appears to work.
    """
    observations, errors = [], []
    try:
        motor.stop(axis)
    except Exception as exc:
        errors.append(f"Stop acknowledgement: {exc}")
        try:
            motor.halt(axis)
        except Exception as fallback:
            errors.append(f"Fallback halt acknowledgement: {fallback}")
    try:
        first = motor.read_registers(axis, (1, 7))
        sleep(.05)
        second = motor.read_registers(axis, (1, 7))
        for sample in (first, second):
            if len(sample) != 2 or any(isinstance(v, bool) or not isinstance(v, int) or not -(2**31) <= v < 2**31 for v in sample):
                raise HardwareSafetyError("Motor stop telemetry is incomplete or invalid.")
            observations.append({"position_raw": sample[0], "velocity_register": sample[1]})
        if first[1] != 0 or second[1] != 0 or first[0] != second[0]:
            raise HardwareSafetyError("Lift stop readback still shows motion.")
    except Exception as exc:
        errors.append(f"Stopped readback: {exc}")
    return observations, "; ".join(errors)


def _finish(motor, axis, store, token, profile, operation, observations, error, *, recovery=False):
    cleanup, cleanup_error = verify_stopped_in_place(motor, axis)
    record = MotorDiagnosticRecord("motion-" + uuid.uuid4().hex, operation,
        datetime.now(timezone.utc).isoformat(), tuple(observations), tuple(cleanup),
        error, cleanup_error, not cleanup_error,
        schema="rapidpy.motion.diagnostic_recovery.v1" if recovery else "rapidpy.motion.diagnostic.v1")
    publish_diagnostic_record(store, record, family="motion")
    if token is not None:
        store.finish(token, profile, record)
    if cleanup_error:
        raise HardwareSafetyError(f"Motor diagnostic stop recovery remains unverified: {cleanup_error}")
    return record


@contextmanager
def motor_diagnostic_operation(motor, axis, operation, profile, *, store=None):
    store = store if store is not None else HardwareSafetyStore(default_safety_path())
    axis = copy.deepcopy(axis)
    with store.operation_lease():
        token = store.begin("motion_diagnostic", operation, profile)
        # Use the persisted strict JSON snapshots for immutable evidence.
        state = store.pending(profile)
        operation, profile = state["plan"], state["profile"]
        observations, error = [], ""
        try:
            yield observations
        except BaseException as exc:
            error = str(exc)
            raise
        finally:
            _finish(motor, axis, store, token, profile, operation, observations, error)


def recover_motor_diagnostic(motor, axis, profile, *, store=None):
    store = store if store is not None else HardwareSafetyStore(default_safety_path())
    with store.operation_lease():
        pending = store.pending()
        if pending is not None and pending['family'] != 'motion_diagnostic':
            raise HardwareSafetyError("Recover the unfinished operation in its original main app or diagnostic helper.")
        if pending is not None:
            pending = store.pending(profile)
        else:
            store.begin('motion_diagnostic', {'action': 'verify_stopped_in_place'}, profile)
            pending = store.pending(profile)
        return _finish(motor, copy.deepcopy(axis), store, pending['token'],
                       profile, pending['plan'],
                       [], '', recovery=True)
