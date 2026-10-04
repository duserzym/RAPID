"""Move a calibrated XY table under the original claimed queue lifetime."""
import copy
import json
from dataclasses import asdict, dataclass, field
from datetime import datetime, timezone
import time
import uuid

from rapidpy_common.hardware_safety import HardwareSafetyError
from rapidpy_common.motor_diagnostic_safety import verify_stopped_in_place
from rapidpy_common.motor_routing import RoutedMotorSerialClient


@dataclass(frozen=True)
class QueueTableMoveRecord:
    operation: dict
    station_profile: dict
    observations: tuple
    cleanup_observations: tuple
    error: str
    cleanup_error: str
    safe_state_confirmed: bool
    simulated: bool = False
    record_id: str = field(default_factory=lambda: uuid.uuid4().hex)
    timestamp_iso: str = field(default_factory=lambda: datetime.now(timezone.utc).isoformat())
    schema: str = 'rapidpy.queue_xy_motion.v1'
    transfer_context: dict | None = None

    def to_dict(self):
        return asdict(self)


class QueueXYTableMotion:
    """No homing/replay: the caller must establish the live station reference."""

    def __init__(self, motor, axes, geometry, *, sleep=time.sleep, should_cancel=lambda: False,
                 transfer_settings=None):
        if not isinstance(motor, RoutedMotorSerialClient) or geometry.use_xy_table is not True:
            raise HardwareSafetyError('Queue XY transfers require the native routed station and XY geometry.')
        if set(axes) != {'changer_x', 'changer_y', 'updown', 'turning'}:
            raise HardwareSafetyError('Queue table transfers require all four bound motor axes.')
        if len(geometry.xy_home) != 2:
            raise HardwareSafetyError('Queue XY motion requires the original accepted home-coordinate pair.')
        if ((geometry.slot_min, geometry.slot_max, geometry.one_step)
                != (motor.config.slot_min, motor.config.slot_max, motor.config.one_step)
                or motor.config.sample_bottom == 0):
            raise HardwareSafetyError('XY geometry and lift clearance must match the accepted motor calibration.')
        self.motor, self.axes, self.geometry = motor, copy.deepcopy(axes), geometry
        self.sleep, self.should_cancel = sleep, should_cancel
        self.transfer_settings = copy.deepcopy(transfer_settings or {})
        self.profile = self._profile()
        self.last_record = None

    def _profile(self):
        return json.loads(json.dumps(dict(helper='rapid_main_queue_xy', geometry=asdict(self.geometry),
                    axes={key: asdict(axis) for key, axis in self.axes.items()},
                    controller=asdict(self.motor.config), transfer_settings=self.transfer_settings)))

    def _validate(self, session):
        session.child_store._owned()
        root = session.store._queue(session.token)
        if (root['profile'].get('helper') != 'rapid_main_queue'
                or root['profile']['stage_profiles'].get('motion') != self.profile
                or self._profile() != self.profile):
            raise HardwareSafetyError('Restore the original queue XY station calibration before motion.')
        for axis in self.axes.values():
            if self.motor._bindings.get(axis.motor_id) != (axis.name, axis.address, axis.port):
                raise HardwareSafetyError('The actual XY transport differs from the original axis binding.')
        if session.is_recovery:
            raise HardwareSafetyError('Queue recovery must stop in place, never replay table transfer motion.')
        if not self.motor.is_connected:
            raise HardwareSafetyError('All original station motor connections must be open.')

    def _check_cancel(self):
        if self.should_cancel():
            raise InterruptedError('Queue table transfer cancelled.')

    def _stopped(self):
        observations, errors = {}, []
        for key, axis in self.axes.items():
            samples, error = verify_stopped_in_place(self.motor, axis, sleep=self.sleep)
            observations[key] = samples
            if error:
                errors.append(f'{key}: {error}')
        return observations, '; '.join(errors)

    def _clearance(self, observations):
        samples = observations.get('updown', [])
        if len(samples) != 2 or any(abs(sample['position_raw']) > abs(self.motor.config.sample_bottom) * .1 for sample in samples):
            raise HardwareSafetyError('Lift clearance is outside the accepted transfer envelope.')
        if self.motor.check_internal_status(self.axes['updown'], 4) != 1:
            raise HardwareSafetyError('The live top-reference switch must verify lift clearance before table transfer.')

    def _reference_proof(self, session, reference):
        if reference is True:
            return None  # Existing explicit operator-attestation callers.
        from .queue_xy_reference import QueueXYReferenceState
        original = QueueXYReferenceState.read(session.store.latest_reference_context(session.token), session, self)
        if not isinstance(reference, QueueXYReferenceState) or reference != original or original.phase != 'referenced':
            raise HardwareSafetyError('The original verified live XY reference is required for transfer.')
        return original.to_dict()

    def move_to_slot(self, session, slot, *, reference_verified=False, field_outputs_off_verified=False,
                     sample_id='', run_id=''):
        self._validate(session)
        self._check_cancel()
        if field_outputs_off_verified is not True:
            raise HardwareSafetyError('Verify the live XY reference and field outputs off before table transfer.')
        reference_proof = self._reference_proof(session, reference_verified)
        target_x, target_y = self.geometry.xy_target(slot)
        operation = dict(action='xy_slot_move', slot=slot, target_x=target_x, target_y=target_y,
                         reference_verified=True, reference_proof=reference_proof, field_outputs_off_verified=True)
        child = session.child_store
        token = child.begin('motion', operation, self.profile, sample_id=sample_id, run_id=run_id)
        observations, error, target_verified = [], '', False
        try:
            initial, stop_error = self._stopped()
            observations.append({'phase': 'before_move', 'axes': initial})
            if stop_error:
                raise HardwareSafetyError(stop_error)
            self._clearance(initial)
            self._check_cancel()
            for key, target in (('changer_x', target_x), ('changer_y', target_y)):
                self.motor.move_motor(self.axes[key], target, self.motor.config.changer_speed, wait_for_stop=False)
                self._check_cancel()
            for key in ('changer_x', 'changer_y'):
                self.motor.wait_for_motor_stop(self.axes[key])
                self._check_cancel()
        except Exception as exc:
            error = str(exc) or type(exc).__name__
        finally:
            cleanup, cleanup_error = self._stopped()
        if not error and not cleanup_error:
            try:
                self._clearance(cleanup)
                x = cleanup['changer_x'][-1]['position_raw']
                y = cleanup['changer_y'][-1]['position_raw']
                if self.geometry.slot_from_xy_counts(x, y) != slot:
                    raise HardwareSafetyError('Both XY readbacks must match the requested registered slot.')
                self._check_cancel()
                target_verified = True
            except Exception as exc:
                error = str(exc) or type(exc).__name__
        record = QueueTableMoveRecord(copy.deepcopy(operation), copy.deepcopy(self.profile),
            tuple(observations), tuple({'axis': key, 'samples': values} for key, values in cleanup.items()),
            error, cleanup_error, target_verified and not error and not cleanup_error)
        self.last_record = record
        child.finish(token, self.profile, record)
        if not record.safe_state_confirmed:
            raise HardwareSafetyError('Queue XY transfer remains unverified: ' + '; '.join(filter(None, (error, cleanup_error))))
        return record
