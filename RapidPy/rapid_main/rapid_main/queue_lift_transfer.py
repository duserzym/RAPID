"""Native pickup, height reference and supported dropoff under a queue claim."""
import copy
from dataclasses import asdict, dataclass, replace
import math
import time

from rapidpy_common.hardware_safety import HardwareSafetyError
from .queue_station import _integer
from .queue_table_motion import QueueTableMoveRecord


@dataclass(frozen=True)
class QueueSpecimenState:
    queue_token: str
    sample_id: str
    original_slot: int
    phase: str
    pickup_position_raw: int | None = None
    sample_height: int | None = None
    file_id: str = ''
    schema: str = 'rapidpy.queue_specimen_state.v1'

    def to_dict(self):
        return asdict(self)

    @classmethod
    def read(cls, value, session, geometry):
        try:
            state = cls(**value)
            if (state.schema != 'rapidpy.queue_specimen_state.v1' or state.queue_token != session.token
                    or not isinstance(state.sample_id, str) or not state.sample_id.strip()
                    or not isinstance(state.file_id, str)
                    or state.phase not in {'picked', 'lifted', 'supported', 'clear', 'unverified'}):
                raise ValueError('invalid specimen identity or phase')
            geometry.specimen_slot(state.original_slot)
            if type(state.original_slot) is not int:
                raise ValueError('invalid original slot type')
            for item in (state.pickup_position_raw, state.sample_height):
                if item is not None and (type(item) is not int or not -(2**31) <= item < 2**31):
                    raise ValueError('invalid specimen count')
            if state.phase in {'lifted', 'supported', 'clear'} and (state.sample_height is None or state.sample_height <= 0):
                raise ValueError('missing referenced sample height')
            if state.phase == 'picked' and state.pickup_position_raw is None:
                raise ValueError('missing pickup position')
            return state
        except (TypeError, ValueError) as exc:
            raise HardwareSafetyError('Original specimen transfer context is invalid: ' + str(exc)) from exc


class QueueLiftTransfer:
    def __init__(self, table, vacuum, *, monotonic=time.monotonic):
        self.table, self.vacuum, self.monotonic = table, vacuum, monotonic
        settings = table.transfer_settings
        delay = settings.get('dropoff_delay_s')
        if (isinstance(delay, bool) or not isinstance(delay, (int, float)) or not math.isfinite(delay)
                or delay < 0 or settings.get('grip_settle_s') != .3):
            raise HardwareSafetyError('Bind the accepted dropoff delay and legacy grip settling time into the queue motion profile.')
        self.dropoff_delay_s = float(delay)

    def context(self, session):
        session.child_store._owned()
        value = session.store.latest_transfer_context(session.token)
        if value is None:
            return None
        state = QueueSpecimenState.read(value, session, self.table.geometry)
        bottom = self.table.motor.config.sample_bottom
        if ((state.sample_height is not None and not 0 < state.sample_height <= abs(bottom))
                or (state.pickup_position_raw is not None and not 0 < state.pickup_position_raw - bottom <= abs(bottom))):
            raise HardwareSafetyError('Original specimen geometry exceeds the accepted lift travel envelope.')
        return state

    def _validate(self, session, field_outputs_off_verified):
        self.table._validate(session)
        if self.dropoff_delay_s != self.table.transfer_settings['dropoff_delay_s']:
            raise HardwareSafetyError('Restore the original queue dropoff settling delay before transfer.')
        self.table._check_cancel()
        from .queue_field_outputs import field_outputs_proof
        field_proof = field_outputs_proof(session, field_outputs_off_verified)
        binding = self.vacuum._queue_binding
        if binding is None or binding.session is not session:
            raise HardwareSafetyError('Specimen transfer requires the original queue vacuum owner.')
        binding._validate_owner()
        if self.vacuum.is_pump_on() is not True:
            raise HardwareSafetyError('Specimen transfer requires acknowledged pump power.')
        return field_proof

    def _state(self, session, phase):
        state = self.context(session)
        if state is None or state.phase != phase:
            raise HardwareSafetyError('Recover the original specimen transfer phase before another lift action.')
        return state

    def _at_slot(self, observations, slot):
        x = observations['changer_x'][-1]['position_raw']
        y = observations['changer_y'][-1]['position_raw']
        if self.table.geometry.slot_from_xy_counts(x, y) != slot:
            raise HardwareSafetyError('The specimen lift requires its original registered XY slot.')

    def _delay(self, seconds):
        deadline = self.monotonic() + seconds
        while True:
            remaining = deadline - self.monotonic()
            if remaining <= 0:
                break
            self.table._check_cancel()
            self.table.sleep(min(.05, remaining))
        self.table._check_cancel()

    def _execute(self, session, state, action, command, before, after, *, field_proof=None):
        operation = dict(action=action, transfer_context=state.to_dict(), field_outputs_proof=field_proof)
        child = session.child_store
        token = child.begin('motion', operation, self.table.profile, sample_id=state.sample_id,
            run_id=session.store._queue(session.token)['run_id'])
        observations, error, result = [], '', None
        final_state = replace(state, phase='unverified')
        try:
            initial, stop_error = self.table._stopped()
            observations.append({'phase': 'before_move', 'axes': initial})
            if stop_error:
                raise HardwareSafetyError(stop_error)
            before(initial)
            self.table._check_cancel()
            result = command()
            if result.success is not True:
                raise HardwareSafetyError('Native lift command did not verify completion.')
            observations.append({'phase': 'native_result', 'result': asdict(result)})
            self.table._check_cancel()
        except Exception as exc:
            error = str(exc) or type(exc).__name__
        finally:
            cleanup, cleanup_error = self.table._stopped()
        if not error and not cleanup_error:
            try:
                final_state = after(result, cleanup)
                self.table._check_cancel()
            except Exception as exc:
                error = str(exc) or type(exc).__name__
        if error or cleanup_error:
            final_state = replace(state, phase='unverified')
        record = QueueTableMoveRecord(copy.deepcopy(operation), copy.deepcopy(self.table.profile),
            tuple(observations), tuple({'axis': key, 'samples': values} for key, values in cleanup.items()),
            error, cleanup_error, not error and not cleanup_error,
            schema='rapidpy.queue_lift_transfer.v1', transfer_context=final_state.to_dict())
        self.table.last_record = record
        child.finish(token, self.table.profile, record)
        if not record.safe_state_confirmed:
            raise HardwareSafetyError('Specimen lift remains unverified; original queue ownership is retained: ' + '; '.join(filter(None, (error, cleanup_error))))
        return record

    def pickup(self, session, original_slot, sample_id, *, file_id='', field_outputs_off_verified=False):
        field_proof = self._validate(session, field_outputs_off_verified)
        previous = self.context(session)
        if previous and previous.phase != 'clear':
            raise HardwareSafetyError('Return the original specimen before loading another.')
        state = QueueSpecimenState(session.token, sample_id, self.table.geometry.specimen_slot(original_slot), 'unverified', file_id=file_id)
        QueueSpecimenState.read(state.to_dict(), session, self.table.geometry)
        if self.vacuum.is_valve_connected() is not False:
            raise HardwareSafetyError('Pickup requires the acknowledged gripper valve OFF.')
        def before(samples):
            self._at_slot(samples, state.original_slot)
            self.table._clearance(samples)
        def after(result, samples):
            self._at_slot(samples, state.original_slot)
            actual = samples['updown'][-1]['position_raw']
            height = actual - self.table.motor.config.sample_bottom
            if not 0 < height <= abs(self.table.motor.config.sample_bottom):
                raise HardwareSafetyError('Pickup readback is outside the accepted specimen travel envelope.')
            if self.table.motor.check_internal_status(self.table.axes['updown'], 4) != 0:
                raise HardwareSafetyError('Pickup cannot retain a set top-reference switch.')
            return replace(state, phase='picked', pickup_position_raw=actual)
        return self._execute(session, state, 'specimen_pickup',
            lambda: self.table.motor.sample_pickup(self.table.axes['updown']), before, after, field_proof=field_proof)

    def home_loaded(self, session, *, field_outputs_off_verified=False):
        field_proof = self._validate(session, field_outputs_off_verified)
        state = self._state(session, 'picked')
        if self.vacuum.is_valve_connected() is not True:
            raise HardwareSafetyError('Lift home requires the original acknowledged specimen grip.')
        def before(samples):
            self._at_slot(samples, state.original_slot)
            if samples['updown'][-1]['position_raw'] != state.pickup_position_raw:
                raise HardwareSafetyError('The lift moved after the recorded specimen pickup.')
        def command():
            self._delay(.3)
            return self.table.motor.home_to_top(self.table.axes['updown'])
        def after(result, samples):
            self._at_slot(samples, state.original_slot)
            self.table._clearance(samples)
            height = state.pickup_position_raw - self.table.motor.config.sample_bottom - _integer(result.final_position, 'top reference')
            if not 0 < height <= abs(self.table.motor.config.sample_bottom):
                raise HardwareSafetyError('Referenced specimen height is outside the accepted travel envelope.')
            return replace(state, phase='lifted', sample_height=height)
        return self._execute(session, state, 'specimen_home_reference', command, before, after, field_proof=field_proof)

    def raise_loaded_to_clearance(self, session, *, field_outputs_off_verified=False):
        field_proof = self._validate(session, field_outputs_off_verified)
        state = self._state(session, 'lifted')
        if self.vacuum.is_valve_connected() is not True:
            raise HardwareSafetyError('Returning lift requires the original acknowledged specimen grip.')
        def empty(samples):
            x = samples['changer_x'][-1]['position_raw']
            y = samples['changer_y'][-1]['position_raw']
            slot = self.table.geometry.slot_from_xy_counts(x, y)
            if not self.table.geometry.is_empty(slot):
                raise HardwareSafetyError('Raise the measured specimen only above a verified empty table location.')
        def after(result, samples):
            empty(samples)
            self.table._clearance(samples)
            if abs(samples['updown'][-1]['position_raw']) > 150:
                raise HardwareSafetyError('Returning specimen lift did not reach zero.')
            return state
        return self._execute(session, state, 'specimen_raise_for_return',
            lambda: self.table.motor.updown_move(self.table.axes['updown'], 0, 2), empty, after, field_proof=field_proof)

    def lower_for_dropoff(self, session, *, field_outputs_off_verified=False):
        field_proof = self._validate(session, field_outputs_off_verified)
        state = self._state(session, 'lifted')
        if self.vacuum.is_valve_connected() is not True:
            raise HardwareSafetyError('Dropoff requires the original acknowledged specimen grip.')
        target = int(self.table.motor.config.sample_bottom + .9 * state.sample_height)
        def before(samples):
            self._at_slot(samples, state.original_slot)
            self.table._clearance(samples)
        def after(result, samples):
            self._at_slot(samples, state.original_slot)
            if (abs(samples['updown'][-1]['position_raw'] - target) > 150
                    or self.table.motor.check_internal_status(self.table.axes['updown'], 4) != 0):
                raise HardwareSafetyError('The original-slot supported dropoff pose is unverified.')
            return replace(state, phase='supported')
        return self._execute(session, state, 'specimen_supported_dropoff',
            lambda: self.table.motor.updown_move(self.table.axes['updown'], target, 0), before, after, field_proof=field_proof)

    def clear_after_release(self, session, *, field_outputs_off_verified=False):
        field_proof = self._validate(session, field_outputs_off_verified)
        state = self._state(session, 'supported')
        if self.vacuum.is_valve_connected() is not False:
            raise HardwareSafetyError('Lift clearance requires acknowledged original-slot valve release.')
        target = int(self.table.motor.config.sample_bottom + .9 * state.sample_height)
        def before(samples):
            self._at_slot(samples, state.original_slot)
            if abs(samples['updown'][-1]['position_raw'] - target) > 150:
                raise HardwareSafetyError('The lift moved after the verified supported dropoff.')
        def command():
            self._delay(self.dropoff_delay_s)
            return self.table.motor.updown_move(self.table.axes['updown'], 0, 1)
        def after(result, samples):
            self._at_slot(samples, state.original_slot)
            self.table._clearance(samples)
            if abs(samples['updown'][-1]['position_raw']) > 150:
                raise HardwareSafetyError('Post-release lift clearance did not reach zero.')
            return replace(state, phase='clear')
        return self._execute(session, state, 'specimen_clear_after_release', command, before, after, field_proof=field_proof)
