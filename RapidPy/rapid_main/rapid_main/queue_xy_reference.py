"""Bounded native two-edge XY referencing under the original queue lifetime."""
import copy
from dataclasses import asdict, dataclass, replace
import math
import time
import uuid

from rapidpy_common.hardware_safety import HardwareSafetyError
from .queue_lift_transfer import QueueSpecimenState
from .queue_holder_geometry import QueueHolderState
from .queue_table_motion import QueueTableMoveRecord


@dataclass(frozen=True)
class QueueXYReferenceState:
    queue_token: str
    reference_id: str
    connection_id: str
    home_counts: tuple
    phase: str = 'referenced'
    schema: str = 'rapidpy.queue_xy_reference_state.v1'

    def to_dict(self):
        return asdict(self)

    @classmethod
    def read(cls, value, session, table):
        try:
            state = cls(**value)
            state = replace(state, home_counts=tuple(state.home_counts))
            if (state.schema != 'rapidpy.queue_xy_reference_state.v1' or state.queue_token != session.token
                    or state.phase not in {'referenced', 'unverified'}
                    or state.connection_id != table.motor._connection_id or not state.connection_id
                    or len(state.home_counts) != 2
                    or any(type(count) is not int or not -(2**31) <= count < 2**31 for count in state.home_counts)
                    or any(abs(actual - expected) > abs(table.geometry.one_step) * .02
                           for actual, expected in zip(state.home_counts, table.geometry.xy_home))):
                raise ValueError('reference identity, transport or accepted home counts changed')
            if uuid.UUID(hex=state.reference_id).hex != state.reference_id:
                raise ValueError('invalid reference ID')
            return state
        except (TypeError, ValueError, AttributeError) as exc:
            raise HardwareSafetyError('Original XY reference is invalid: ' + str(exc)) from exc


@dataclass(frozen=True)
class QueueXYReferenceRecord(QueueTableMoveRecord):
    reference_context: dict | None = None
    schema: str = 'rapidpy.queue_xy_reference.v1'


class QueueXYReference:
    """Reference an empty station; recovery never replays a homing sweep."""
    def __init__(self, table, vacuum, *, monotonic=time.monotonic, timeout_s=120):
        if isinstance(timeout_s, bool) or not isinstance(timeout_s, (int, float)) or not math.isfinite(timeout_s) or not 0 < timeout_s <= 360:
            raise HardwareSafetyError('XY reference timeout must be finite and within 360 seconds.')
        self.table, self.vacuum, self.monotonic, self.timeout_s = table, vacuum, monotonic, timeout_s
        self.last_record = None
        self._active_session = None
        self._active_connection_id = None

    def require(self, session):
        self.table._validate(session)
        state = QueueXYReferenceState.read(session.store.latest_reference_context(session.token), session, self.table)
        if state.phase != 'referenced':
            raise HardwareSafetyError('Recover the unfinished stage before using the original XY reference.')
        return state

    def _check(self, deadline):
        if self._active_session is not None:
            self.table._validate(self._active_session)
            if self.table.motor._connection_id != self._active_connection_id:
                raise HardwareSafetyError('Original motor connection changed during XY referencing.')
        self.table._check_cancel()
        if self.monotonic() >= deadline:
            raise TimeoutError('Native XY reference deadline expired before both switch edges were verified.')

    def _switches(self, phase):
        result = {}
        for key, neg, pos in (('changer_x', 4, 5), ('changer_y', 5, 6)):
            axis = self.table.axes[key]
            negative = self.table.motor.check_internal_status(axis, neg)
            positive = self.table.motor.check_internal_status(axis, pos)
            if type(negative) is not int or type(positive) is not int or negative not in (0, 1) or positive not in (0, 1):
                raise HardwareSafetyError('XY switch readbacks must be complete binary values.')
            if negative == positive == 0:
                raise HardwareSafetyError('Both opposing XY limit switches appear active; reference is contradictory.')
            result[key] = dict(negative=negative, positive=positive)
        if self.table.motor.check_internal_status(self.table.axes['updown'], 4) != 1:
            raise HardwareSafetyError('Live lift top switch lost clearance during XY referencing.')
        return dict(phase=phase, switches=result)

    def _phase(self, phase, deadline, observations, *, distance_override=None):
        direction = -1 if phase == 'negative' else 1
        distance = self.table.motor.config.xy_neg_homing_distance if direction < 0 else self.table.motor.config.xy_pos_homing_distance
        if distance_override is not None:
            if (type(distance_override) is not int or distance_override * direction <= 0
                    or abs(distance_override) > abs(distance)):
                raise HardwareSafetyError('Bounded XY travel must stay within the accepted homing envelope.')
            distance = distance_override
        if type(distance) is not int or not -(2**31) <= distance < 2**31 or distance * direction <= 0:
            raise HardwareSafetyError('Accepted XY homing travel has the wrong sign or count range.')
        active = 'negative' if direction < 0 else 'positive'
        issued = set()
        while True:
            self._check(deadline)
            switches = self._switches(phase)
            observations.append(switches)
            if all(value[active] == 0 for value in switches['switches'].values()):
                break
            for key, value in switches['switches'].items():
                if value[active] == 0 or key in issued:
                    continue
                self._check(deadline)
                axis = self.table.axes[key]
                position = self.table.motor.read_position(axis)
                if type(position) is not int or not -(2**31) <= position + distance < 2**31:
                    raise HardwareSafetyError('Relative XY homing travel would overflow controller counts.')
                self.table.motor.move_motor(axis, distance, self.table.motor.config.changer_speed,
                    wait_for_stop=False, stop_enable=(-1 if key == 'changer_x' else -2) + (0 if direction < 0 else -1),
                    stop_condition=0, relative_mode=True)
                issued.add(key)
            self.table.sleep(.05)
        stopped, error = self.table._stopped()
        observations.append(dict(phase=phase + '_stop', axes=stopped))
        if error:
            raise HardwareSafetyError(error)
        self.table._clearance(stopped)
        self._check(deadline)
        final = self._switches(phase + '_verified')
        observations.append(final)
        if any(value[active] != 0 for value in final['switches'].values()):
            raise HardwareSafetyError('XY switch edge changed before independent stopped verification.')

    def home(self, session, *, empty_rod_confirmed=False, field_outputs_off_verified=False):
        self.table._validate(session)
        self.table._check_cancel()
        if empty_rod_confirmed is not True:
            raise HardwareSafetyError('Confirm the empty rod and verified field outputs off before native XY referencing.')
        from .queue_field_outputs import field_outputs_proof
        field_proof = field_outputs_proof(session, field_outputs_off_verified)
        binding = self.vacuum._queue_binding
        if binding is None or binding.session is not session:
            raise HardwareSafetyError('XY referencing requires the original queue vacuum owner.')
        binding._validate_owner()
        if self.vacuum.output_state_known is not True or self.vacuum.is_valve_connected() is not False:
            raise HardwareSafetyError('XY referencing requires acknowledged valve OFF.')
        specimen = session.store.latest_transfer_context(session.token)
        if specimen is not None and QueueSpecimenState.read(specimen, session, self.table.geometry).phase != 'clear':
            raise HardwareSafetyError('Return the original specimen before XY referencing.')
        holder = session.store.latest_holder_context(session.token)
        if holder is not None and QueueHolderState.read(holder, session, self.table.geometry).phase != 'clear':
            raise HardwareSafetyError('Return and clear the original blank holder before XY referencing.')
        state = QueueXYReferenceState(session.token, uuid.uuid4().hex, self.table.motor._connection_id,
                                      tuple(self.table.geometry.xy_home))
        operation = dict(action='xy_home_reference', empty_rod_confirmed=True, field_outputs_off_verified=True,
                         timeout_s=self.timeout_s, reference_context=state.to_dict(), field_outputs_proof=field_proof)
        profile, child = self.table.profile, session.child_store
        token = child.begin('motion', operation, profile, run_id=session.store._queue(session.token)['run_id'])
        deadline = self.monotonic() + self.timeout_s
        self._active_session, self._active_connection_id = session, state.connection_id
        observations, error, verified = [], '', False
        try:
            initial, failure = self.table._stopped()
            observations.append(dict(phase='before_reference', axes=initial))
            if failure:
                raise HardwareSafetyError(failure)
            self.table._clearance(initial)
            self._phase('negative', deadline, observations)
            self._phase('positive', deadline, observations)
            for key in ('changer_x', 'changer_y'):
                self._check(deadline)
                self.table.motor.zero_target_pos(self.table.axes[key])
        except Exception as exc:
            error = str(exc) or type(exc).__name__
        finally:
            cleanup, cleanup_error = self.table._stopped()
        if not error and not cleanup_error:
            try:
                self._check(deadline)
                self.table._clearance(cleanup)
                final = self._switches('reference_verified')
                observations.append(final)
                if any(value['positive'] != 0 for value in final['switches'].values()):
                    raise HardwareSafetyError('Both final positive XY edges must remain active.')
                counts = tuple(cleanup[key][-1]['position_raw'] for key in ('changer_x', 'changer_y'))
                state = replace(state, home_counts=counts)
                QueueXYReferenceState.read(state.to_dict(), session, self.table)
                binding._validate_owner()
                if self.vacuum.output_state_known is not True or self.vacuum.is_valve_connected() is not False:
                    raise HardwareSafetyError('Valve evidence changed during XY referencing.')
                verified = True
            except Exception as exc:
                error = str(exc) or type(exc).__name__
        if not verified:
            state = replace(state, phase='unverified')
        record = QueueXYReferenceRecord(copy.deepcopy(operation), copy.deepcopy(profile), tuple(observations),
            tuple(dict(axis=key, samples=values) for key, values in cleanup.items()), error, cleanup_error,
            verified, reference_context=state.to_dict())
        self.last_record = record
        try:
            child.finish(token, profile, record)
        finally:
            self._active_session, self._active_connection_id = None, None
        if not verified:
            raise HardwareSafetyError('Native XY reference remains pending: ' + '; '.join(filter(None, (error, cleanup_error))))
        return state
