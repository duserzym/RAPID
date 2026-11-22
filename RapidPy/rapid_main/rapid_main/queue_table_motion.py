"""Move a calibrated changer under the original claimed queue lifetime."""
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


class _QueueTableMotion:
    def __init__(self, motor, axes, geometry, *, sleep=time.sleep, should_cancel=lambda: False,
                 transfer_settings=None):
        if not isinstance(motor, RoutedMotorSerialClient) or geometry.use_xy_table is not self.use_xy_table:
            raise HardwareSafetyError('Queue transfers require the native routed motor and matching changer mode.')
        if set(axes) != self.axis_keys:
            raise HardwareSafetyError('Queue transfers require exactly the original station motor axes.')
        if ((geometry.slot_min, geometry.slot_max, geometry.one_step)
                != (motor.config.slot_min, motor.config.slot_max, motor.config.one_step)
                or motor.config.sample_bottom == 0):
            raise HardwareSafetyError('Changer geometry and lift clearance must match the accepted motor calibration.')
        self.motor, self.axes, self.geometry = motor, copy.deepcopy(axes), geometry
        self.sleep, self.should_cancel = sleep, should_cancel
        self.transfer_settings = copy.deepcopy(transfer_settings or {})
        self.profile = self._profile()
        self.last_record = None

    def _profile(self):
        return json.loads(json.dumps(dict(helper=self.helper, geometry=asdict(self.geometry),
                    axes={key: asdict(axis) for key, axis in self.axes.items()},
                    controller=asdict(self.motor.config), transfer_settings=self.transfer_settings)))

    def _validate(self, session):
        session.child_store._owned()
        root = session.store._queue(session.token)
        if (root['profile'].get('helper') != 'rapid_main_queue'
                or root['profile']['stage_profiles'].get('motion') != self.profile
                or self._profile() != self.profile):
            raise HardwareSafetyError('Restore the original queue station calibration before motion.')
        for axis in self.axes.values():
            if self.motor._bindings.get(axis.motor_id) != (axis.name, axis.address, axis.port):
                raise HardwareSafetyError('The actual transport differs from the original axis binding.')
        if session.is_recovery:
            raise HardwareSafetyError('Queue recovery must stop in place, never replay table transfer motion.')
        if not self.motor.is_connected:
            raise HardwareSafetyError('All original station motor connections must be open.')

    def _check_cancel(self):
        if self.should_cancel():
            raise InterruptedError('Queue table transfer cancelled.')

    def _stopped(self, *, available_ports=None):
        observations, errors = {}, []
        for key, axis in self.axes.items():
            if available_ports is not None and axis.port.upper() not in available_ports:
                samples, error = [], 'Original controller port is unavailable.'
            else:
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

    def slot_from_samples(self, samples):
        positions = {key: samples[key][-1]['position_raw'] for key in self.axis_keys if key.startswith('changer_')}
        return self.geometry.slot_from_positions(
            positions, alignment_offset=self.motor.config.sample_hole_alignment_offset)

    def move_to_slot(self, session, slot, *, reference_verified=False, field_outputs_off_verified=False,
                     sample_id='', run_id=''):
        self._validate(session)
        self._check_cancel()
        from .queue_field_outputs import field_outputs_proof
        field_proof = field_outputs_proof(session, field_outputs_off_verified)
        reference_proof = self._reference_proof(session, reference_verified)
        operation, targets = self._plan_slot_move(slot)
        operation.update(reference_verified=True, reference_proof=reference_proof, field_outputs_off_verified=True,
                         field_outputs_proof=field_proof)
        child = session.child_store
        token = child.begin('motion', operation, self.profile, sample_id=sample_id, run_id=run_id)
        observations, error, target_verified = [], '', False
        try:
            initial, stop_error = self._stopped()
            observations.append({'phase': 'before_move', 'axes': initial})
            if stop_error:
                raise HardwareSafetyError(stop_error)
            self._clearance(initial)
            self._verify_start(operation, initial)
            self._check_cancel()
            for key, target in targets.items():
                self.motor.move_motor(self.axes[key], target, self.motor.config.changer_speed, wait_for_stop=False)
                self._check_cancel()
            for key in targets:
                self.motor.wait_for_motor_stop(self.axes[key])
                self._check_cancel()
        except Exception as exc:
            error = str(exc) or type(exc).__name__
        finally:
            cleanup, cleanup_error = self._stopped()
        if not error and not cleanup_error:
            try:
                self._clearance(cleanup)
                self._verify_target(operation, cleanup)
                self._check_cancel()
                target_verified = True
            except Exception as exc:
                error = str(exc) or type(exc).__name__
        record = QueueTableMoveRecord(copy.deepcopy(operation), copy.deepcopy(self.profile),
            tuple(observations), tuple({'axis': key, 'samples': values} for key, values in cleanup.items()),
            error, cleanup_error, target_verified and not error and not cleanup_error, schema=self.move_schema)
        self.last_record = record
        child.finish(token, self.profile, record)
        if not record.safe_state_confirmed:
            raise HardwareSafetyError('Queue changer transfer remains unverified: ' + '; '.join(filter(None, (error, cleanup_error))))
        return record

    def _verify_start(self, operation, initial):
        pass

    def _verify_target(self, operation, cleanup):
        if self.slot_from_samples(cleanup) != operation['slot']:
            raise HardwareSafetyError('Stationary changer readbacks must match the requested registered slot.')


class QueueXYTableMotion(_QueueTableMotion):
    """Move the referenced XY table without replaying reference or transfer."""

    use_xy_table = True
    axis_keys = frozenset({'changer_x', 'changer_y', 'updown', 'turning'})
    helper = 'rapid_main_queue_xy'
    move_schema = 'rapidpy.queue_xy_motion.v1'

    def __init__(self, motor, axes, geometry, **kwargs):
        if len(geometry.xy_home) != 2:
            raise HardwareSafetyError('Queue XY motion requires the original accepted home-coordinate pair.')
        super().__init__(motor, axes, geometry, **kwargs)

    def _reference_proof(self, session, reference):
        if reference is True:
            return None  # Existing explicit operator-attestation callers.
        from .queue_xy_reference import QueueXYReferenceState
        original = QueueXYReferenceState.read(session.store.latest_reference_context(session.token), session, self)
        if not isinstance(reference, QueueXYReferenceState) or reference != original or original.phase != 'referenced':
            raise HardwareSafetyError('The original verified live XY reference is required for transfer.')
        return original.to_dict()

    def _plan_slot_move(self, slot):
        x, y = self.geometry.xy_target(slot)
        return dict(action='xy_slot_move', slot=slot, target_x=x, target_y=y), {'changer_x': x, 'changer_y': y}


class QueueChainTableMotion(_QueueTableMotion):
    """Move the single chain axis after the original operator establishes its origin."""

    use_xy_table = False
    axis_keys = frozenset({'changer_x', 'updown', 'turning'})
    helper = 'rapid_main_queue_chain'
    move_schema = 'rapidpy.queue_chain_motion.v1'

    def __init__(self, motor, axes, geometry, **kwargs):
        super().__init__(motor, axes, geometry, **kwargs)
        offset = motor.config.sample_hole_alignment_offset
        if not abs(offset) < .5 or abs(geometry.one_step) * geometry.slot_max >= 2**31:
            raise HardwareSafetyError('Chain alignment or full-loop travel exceeds the calibrated count envelope.')

    def _reference_proof(self, session, reference):
        from .queue_chain_reference import QueueChainReferenceState
        original = QueueChainReferenceState.read(session.store.latest_reference_context(session.token), session, self)
        if not isinstance(reference, QueueChainReferenceState) or reference != original or original.phase != 'referenced':
            raise HardwareSafetyError('The original verified live chain reference is required for transfer.')
        return original.to_dict()

    def _plan_slot_move(self, slot):
        offset = self.motor.config.sample_hole_alignment_offset
        if type(slot) is not int:
            raise HardwareSafetyError('Use an integer chain slot or the nearest empty-hole marker.')
        if slot != 0:
            self.geometry.chain_reference_counts(slot, offset)
        current = self.motor.read_position(self.axes['changer_x'])
        actual_slot = self.geometry.slot_from_positions({'changer_x': current}, alignment_offset=offset)
        if slot == 0:
            slot = self.geometry.nearest_empty(actual_slot)
        target = self.geometry.chain_target(slot, current, offset)
        if self.geometry.slot_from_positions({'changer_x': target}, alignment_offset=offset) != slot:
            raise HardwareSafetyError('The calculated chain target cannot identify the requested slot.')
        return (dict(action='chain_slot_move', slot=slot, from_slot=actual_slot, starting_counts=current,
                     target_counts=target, connection_id=self.motor._connection_id),
                {'changer_x': target})

    def slot_from_samples(self, samples):
        offset = self.motor.config.sample_hole_alignment_offset
        slots = {self.geometry.slot_from_positions({'changer_x': sample['position_raw']}, alignment_offset=offset)
                 for sample in samples['changer_x']}
        if len(slots) != 1:
            raise HardwareSafetyError('Stopped chain readbacks must identify the same calibrated slot.')
        return slots.pop()

    def _verify_start(self, operation, initial):
        samples = initial['changer_x']
        if (any(sample['position_raw'] != operation['starting_counts'] for sample in samples)
                or self.motor._connection_id != operation['connection_id']):
            raise HardwareSafetyError('The live chain position changed after planning the original transfer.')

    def _verify_target(self, operation, cleanup):
        super()._verify_target(operation, cleanup)
        if (any(abs(sample['position_raw'] - operation['target_counts']) > abs(self.geometry.one_step) * .02
                for sample in cleanup['changer_x']) or self.motor._connection_id != operation['connection_id']):
            raise HardwareSafetyError('Chain readbacks exceed the journaled target or original connection.')
