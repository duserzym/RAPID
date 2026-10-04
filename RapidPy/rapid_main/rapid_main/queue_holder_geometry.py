"""Blank-holder pose evidence separate from a positive-height loaded specimen."""
import copy
from dataclasses import asdict, dataclass, replace

from rapidpy_common.hardware_safety import HardwareSafetyError
from .queue_specimen_geometry import QueueSpecimenGeometry
from .queue_table_motion import QueueTableMoveRecord
from .queue_lift_transfer import QueueSpecimenState


@dataclass(frozen=True)
class QueueHolderState:
    queue_token: str
    sample_id: str
    hole: int
    phase: str = 'ready'
    sample_height: int = 0
    schema: str = 'rapidpy.queue_blank_holder_state.v1'

    def to_dict(self):
        return asdict(self)

    @classmethod
    def read(cls, value, session, station):
        try:
            state = cls(**value)
            if (state.schema != 'rapidpy.queue_blank_holder_state.v1'
                    or state.queue_token != session.token or type(state.hole) is not int
                    or not station.is_empty(state.hole) or state.sample_id != f'holder-{state.hole:03d}'
                    or type(state.sample_height) is not int or state.sample_height != 0
                    or state.phase not in {'ready', 'clear', 'unverified'}):
                raise ValueError('invalid blank-holder identity, hole, height or phase')
            return state
        except (TypeError, ValueError) as exc:
            raise HardwareSafetyError('Original blank-holder context is invalid: ' + str(exc)) from exc


class QueueHolderGeometry(QueueSpecimenGeometry):
    is_holder = True

    def __init__(self, session, config, motor, axes, store, vacuum):
        self.vacuum = vacuum
        super().__init__(session, config, motor, axes, store)

    def _validate_station(self):
        root = super()._validate_station()
        binding = self.vacuum._queue_binding
        if binding is None or binding.session is not self.session:
            raise HardwareSafetyError('Blank-holder geometry requires the original queue vacuum owner.')
        binding._validate_owner()
        if (self.vacuum.output_state_known is not True or self.vacuum.is_pump_on() is not True
                or self.vacuum.is_valve_connected() is not False):
            raise HardwareSafetyError('Blank-holder geometry requires acknowledged pump ON and valve OFF.')
        return root

    def _read_context(self):
        value = self.session.store.latest_transfer_context(self.session.token)
        if value is not None and QueueSpecimenState.read(value, self.session, self.station).phase != 'clear':
            raise HardwareSafetyError('Return the original specimen before binding blank-holder geometry.')
        state = QueueHolderState.read(self.session.store.latest_holder_context(self.session.token),
                                     self.session, self.station)
        if state.phase != 'ready':
            raise HardwareSafetyError('Verify the original blank-holder pose before acquisition.')
        return state


@dataclass(frozen=True)
class QueueHolderPoseRecord(QueueTableMoveRecord):
    holder_context: dict | None = None
    schema: str = 'rapidpy.queue_blank_holder_pose.v1'


class QueueBlankHolderMotion:
    """Attest an empty-hole pose or raise an acquired blank rod to clearance."""
    def __init__(self, table, vacuum):
        self.table, self.vacuum = table, vacuum
        self.last_record = None

    def _validate(self, session, field_outputs_off_verified):
        self.table._validate(session)
        self.table._check_cancel()
        from .queue_field_outputs import field_outputs_proof
        field_proof = field_outputs_proof(session, field_outputs_off_verified)
        binding = self.vacuum._queue_binding
        if binding is None or binding.session is not session:
            raise HardwareSafetyError('Blank-holder motion requires the original queue vacuum owner.')
        binding._validate_owner()
        value = session.store.latest_transfer_context(session.token)
        if value is not None:
            specimen = QueueSpecimenState.read(value, session, self.table.geometry)
            if specimen.phase != 'clear':
                raise HardwareSafetyError('Return the original specimen before blank-holder acquisition.')
        if (self.vacuum.output_state_known is not True or self.vacuum.is_pump_on() is not True
                or self.vacuum.is_valve_connected() is not False):
            raise HardwareSafetyError('A blank holder requires acknowledged pump ON and valve OFF.')
        return field_proof

    def _at_hole(self, samples, hole):
        x, y = (samples[key][-1]['position_raw'] for key in ('changer_x', 'changer_y'))
        if self.table.geometry.slot_from_xy_counts(x, y) != hole:
            raise HardwareSafetyError('Both XY readbacks must verify the calibrated empty hole.')

    def prepare(self, session, hole, *, reference_verified=False, field_outputs_off_verified=False):
        field_proof = self._validate(session, field_outputs_off_verified)
        reference_proof = self.table._reference_proof(session, reference_verified)
        state = QueueHolderState(session.token, f'holder-{hole:03d}', hole)
        QueueHolderState.read(state.to_dict(), session, self.table.geometry)
        return self._execute(session, state, 'verify_blank_holder_pose', move=False, reference_proof=reference_proof,
                             field_proof=field_proof)

    def return_to_clearance(self, session, *, field_outputs_off_verified=False):
        field_proof = self._validate(session, field_outputs_off_verified)
        state = QueueHolderState.read(session.store.latest_holder_context(session.token), session, self.table.geometry)
        if state.phase != 'ready':
            raise HardwareSafetyError('Recover the original blank-holder acquisition before returning the rod.')
        return self._execute(session, state, 'return_blank_holder_clearance', move=True, field_proof=field_proof)

    def _execute(self, session, state, action, *, move, reference_proof=None, field_proof=None):
        operation = dict(action=action, holder_context=state.to_dict(), field_outputs_off_verified=True,
                         reference_proof=reference_proof, field_outputs_proof=field_proof)
        child, profile = session.child_store, self.table.profile
        token = child.begin('motion', operation, profile, sample_id=state.sample_id,
                            run_id=session.store._queue(session.token)['run_id'])
        observations, error, cleanup_error, verified = [], '', '', False
        try:
            initial, failure = self.table._stopped()
            observations.append(dict(phase='before', axes=initial))
            if failure:
                raise HardwareSafetyError(failure)
            self._at_hole(initial, state.hole)
            if not move:
                self.table._clearance(initial)
            else:
                result = self.table.motor.updown_move(self.table.axes['updown'], 0, 2)
                if result.success is not True:
                    raise HardwareSafetyError('Blank-holder return motion did not reach clearance.')
            self.table._check_cancel()
        except Exception as exc:
            error = str(exc) or type(exc).__name__
        finally:
            cleanup, cleanup_error = self.table._stopped()
        if not error and not cleanup_error:
            try:
                self._at_hole(cleanup, state.hole)
                self.table._clearance(cleanup)
                self.table._check_cancel()
                if (self.vacuum.output_state_known is not True or self.vacuum.is_valve_connected() is not False
                        or self.vacuum.is_pump_on() is not True):
                    raise HardwareSafetyError('Blank-holder valve state changed during motion.')
                verified = True
            except Exception as exc:
                error = str(exc) or type(exc).__name__
        context = replace(state, phase='clear' if move else 'ready') if verified else replace(state, phase='unverified')
        record = QueueHolderPoseRecord(copy.deepcopy(operation), copy.deepcopy(profile), tuple(observations),
            tuple(dict(axis=key, samples=values) for key, values in cleanup.items()), error, cleanup_error,
            verified, holder_context=context.to_dict())
        self.last_record = record
        child.finish(token, profile, record)
        if not verified:
            raise HardwareSafetyError('Blank-holder pose remains pending: ' + '; '.join(filter(None, (error, cleanup_error))))
        return record
