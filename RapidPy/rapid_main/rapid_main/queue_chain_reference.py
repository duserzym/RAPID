"""Establish the chain origin using the operator's observed physical slot."""
from dataclasses import asdict, dataclass
import uuid

from rapidpy_common.hardware_safety import HardwareSafetyError
from .queue_field_outputs import field_outputs_proof
from .queue_lift_transfer import QueueSpecimenState
from .queue_holder_geometry import QueueHolderState
from .queue_table_motion import QueueTableMoveRecord


@dataclass(frozen=True)
class QueueChainReferenceState:
    queue_token: str
    reference_id: str
    connection_id: str
    slot: int
    position_raw: int
    phase: str = 'referenced'
    schema: str = 'rapidpy.queue_chain_reference_state.v1'

    def to_dict(self):
        return asdict(self)

    @classmethod
    def read(cls, value, session, table):
        try:
            state = cls(**value)
            offset = table.motor.config.sample_hole_alignment_offset
            if (state.schema != 'rapidpy.queue_chain_reference_state.v1'
                    or state.queue_token != session.token or state.phase not in {'referenced', 'unverified'}
                    or not state.connection_id or state.connection_id != table.motor._connection_id
                    or uuid.UUID(hex=state.reference_id).hex != state.reference_id
                    or type(state.slot) is not int
                    or table.geometry.slot_from_positions({'changer_x': state.position_raw}, alignment_offset=offset) != state.slot):
                raise ValueError('reference identity, transport or calibrated slot changed')
            return state
        except (TypeError, ValueError, AttributeError) as exc:
            raise HardwareSafetyError('Original chain reference is invalid: ' + str(exc)) from exc


@dataclass(frozen=True)
class QueueChainReferenceRecord(QueueTableMoveRecord):
    reference_context: dict | None = None
    schema: str = 'rapidpy.queue_chain_reference.v1'


class QueueChainReference:
    def __init__(self, table, vacuum):
        self.table, self.vacuum = table, vacuum

    def require(self, session):
        self.table._validate(session)
        state = QueueChainReferenceState.read(session.store.latest_reference_context(session.token), session, self.table)
        if state.phase != 'referenced':
            raise HardwareSafetyError('Recover the unfinished stage before using the original chain reference.')
        return state

    def home(self, session, *, empty_rod_confirmed=False, field_outputs_off_verified=False):
        self.table._validate(session)
        self.table._check_cancel()
        root = session.store._queue(session.token)
        attestation = root['plan'].get('startup', {})
        slot = attestation.get('chain_slot')
        operator = attestation.get('operator')
        if (empty_rod_confirmed is not True or attestation.get('empty_rod_confirmed') is not True
                or not isinstance(operator, str) or not operator.strip() or type(slot) is not int
                or session.store.latest_reference_context(session.token) is not None):
            raise HardwareSafetyError('Initial chain referencing requires the original empty rod and confirmed physical slot.')
        proof = field_outputs_proof(session, field_outputs_off_verified)
        offset = self.table.motor.config.sample_hole_alignment_offset
        target = self.table.geometry.chain_reference_counts(slot, offset)
        state = QueueChainReferenceState(session.token, uuid.uuid4().hex, self.table.motor._connection_id, slot, target)
        QueueChainReferenceState.read(state.to_dict(), session, self.table)
        binding = self.vacuum._queue_binding

        def verify_vacuum():
            if (binding is None or binding is not self.vacuum._queue_binding or binding.session is not session
                    or self.vacuum.output_state_known is not True or self.vacuum.is_pump_on() is not True
                    or self.vacuum.is_valve_connected() is not False):
                raise HardwareSafetyError('Chain referencing requires original pump ON and valve OFF evidence.')
            binding._validate_owner()

        verify_vacuum()
        specimen = session.store.latest_transfer_context(session.token)
        holder = session.store.latest_holder_context(session.token)
        if (specimen is not None and QueueSpecimenState.read(specimen, session, self.table.geometry).phase != 'clear'
                or holder is not None and QueueHolderState.read(holder, session, self.table.geometry).phase != 'clear'):
            raise HardwareSafetyError('Clear the original specimen and blank holder before chain referencing.')

        operation = dict(action='reference_confirmed_chain_slot', attestation=attestation, target_counts=target,
                         field_outputs_proof=proof, reference_context=state.to_dict())
        child = session.child_store
        token = child.begin('motion', operation, self.table.profile, run_id=root['run_id'])
        observations, error = [], ''
        try:
            initial, failure = self.table._stopped()
            observations.append(dict(phase='before_reference', axes=initial))
            if failure:
                raise HardwareSafetyError(failure)
            self.table._clearance(initial)
            self.table._check_cancel()
            self.table.motor.relabel_pos(self.table.axes['changer_x'], target,
                                         tolerance=min(10, abs(self.table.geometry.one_step) * .02))
        except Exception as exc:
            error = str(exc) or type(exc).__name__
        finally:
            cleanup, cleanup_error = self.table._stopped()
        if not error and not cleanup_error:
            try:
                self.table._validate(session)
                self.table._check_cancel()
                self.table._clearance(cleanup)
                self.table._verify_target(dict(slot=slot, target_counts=target, connection_id=state.connection_id), cleanup)
                verify_vacuum()
            except Exception as exc:
                error = str(exc) or type(exc).__name__

        verified = not bool(error or cleanup_error)
        context = state.to_dict()
        if verified:
            context['position_raw'] = cleanup['changer_x'][-1]['position_raw']
        else:
            context['phase'] = 'unverified'
        record = QueueChainReferenceRecord(operation, self.table.profile, tuple(observations),
            tuple(dict(axis=axis, samples=values) for axis, values in cleanup.items()),
            error, cleanup_error, verified, reference_context=context)
        child.finish(token, self.table.profile, record)
        if not verified:
            raise HardwareSafetyError('Native chain reference remains pending: ' + '; '.join(filter(None, (error, cleanup_error))))
        return self.require(session)
