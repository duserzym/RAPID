"""Original empty-station loading corner and durable human tray confirmations."""
import copy
from dataclasses import dataclass

from rapidpy_common.hardware_safety import HardwareSafetyError
from .queue_table_motion import QueueTableMoveRecord


@dataclass(frozen=True)
class QueueOperatorRecord(QueueTableMoveRecord):
    operator_context: dict | None = None
    schema: str = 'rapidpy.queue_operator_confirmation.v1'


class QueueOperatorStages:
    def __init__(self, backend, coordinator):
        self.backend, self.coordinator = backend, coordinator
        self.session, self.table = coordinator.session, coordinator.table
        self.last_park = None
        self._confirmed_park_id = None

    def _idle(self):
        if (self.backend._queue_coordinator is not self.coordinator
                or self.backend._queue_specimen_geometry is not None):
            raise HardwareSafetyError('Return and clear the original rod before handling the specimen tray.')
        self.backend._validate_queue_bindings(self.coordinator)
        return self.coordinator._validate()

    def _corner(self, samples):
        vacuum = self.coordinator.vacuum
        if (vacuum._queue_binding is None or vacuum._queue_binding.session is not self.session
                or vacuum.output_state_known is not True or vacuum.is_pump_on() is not True):
            raise HardwareSafetyError('Original acknowledged queue vacuum ownership must remain unchanged.')
        vacuum._queue_binding._validate_owner()
        self.table._clearance(samples)
        state = self.coordinator.reference._switches('loading_corner')
        if any(value['negative'] != 0 for value in state['switches'].values()):
            raise HardwareSafetyError('Both original negative XY edges must verify the loading corner.')
        if self.coordinator.vacuum.is_valve_connected() is not False:
            raise HardwareSafetyError('The original valve must remain OFF during tray handling.')
        return state

    def park(self):
        reference = self._idle()
        fields = self.coordinator.fields.verify_off(self.session)
        # VB6 MoveToCorner uses one -3,000,000 count sweep per negative edge;
        # never enlarge a smaller accepted native homing envelope.
        distance = -min(3_000_000, abs(self.table.motor.config.xy_neg_homing_distance))
        operation = dict(action='park_native_loading_corner', connection_id=self.table.motor._connection_id,
            reference_proof=reference.to_dict(), field_outputs_proof=dict(fields), travel_counts=distance)
        token = self.session.child_store.begin('motion', operation, self.table.profile)
        reference_service = self.coordinator.reference
        deadline = reference_service.monotonic() + reference_service.timeout_s
        reference_service._active_session = self.session
        reference_service._active_connection_id = self.table.motor._connection_id
        observations, error = [], ''
        try:
            initial, failure = self.table._stopped()
            observations.append(dict(phase='before_corner', axes=initial))
            if failure:
                raise HardwareSafetyError(failure)
            self.table._clearance(initial)
            reference_service._phase('negative', deadline, observations, distance_override=distance)
        except Exception as exc:
            error = str(exc) or type(exc).__name__
        finally:
            cleanup, cleanup_error = self.table._stopped()
            reference_service._active_session = reference_service._active_connection_id = None
        if not error and not cleanup_error:
            try:
                self.table._check_cancel()
                self.backend._validate_queue_bindings(self.coordinator)
                if self.table.motor._connection_id != operation['connection_id']:
                    raise HardwareSafetyError('Original motor connection changed during loading-corner motion.')
                observations.append(self._corner(cleanup))
            except Exception as exc:
                error = str(exc) or type(exc).__name__
        record = QueueTableMoveRecord(copy.deepcopy(operation), copy.deepcopy(self.table.profile),
            tuple(observations), tuple(dict(axis=key, samples=value) for key, value in cleanup.items()),
            error, cleanup_error, not bool(error or cleanup_error), schema='rapidpy.queue_loading_corner.v1')
        self.session.child_store.finish(token, self.table.profile, record)
        if not record.safe_state_confirmed:
            raise HardwareSafetyError('Native loading corner remains pending: ' + '; '.join(filter(None, (error, cleanup_error))))
        self.last_park = record
        return record

    def directions(self):
        self.session.child_store._owned()
        root = self.session.store._queue(self.session.token)
        context = self.session.store.latest_operator_context(self.session.token)
        if (not isinstance(context, dict) or context.get('schema') != 'rapidpy.queue_tray_state.v1'
                or context.get('queue_token') != self.session.token or context.get('phase') != 'confirmed'
                or not isinstance(context.get('file_directions'), dict)
                or set(context['file_directions']) != set(root['plan']['commands']['file_directions'])
                or any(type(value) is not bool for value in context['file_directions'].values())):
            raise HardwareSafetyError('Record the original operator tray confirmation before selecting specimen orientation.')
        return copy.deepcopy(context['file_directions'])

    def confirm(self, kind, file_id, *, operator):
        self._idle()
        if kind not in {'load', 'flip'} or not isinstance(operator, str) or not operator.strip():
            raise HardwareSafetyError('A named original tray load or flip confirmation is required.')
        root = self.session.store._queue(self.session.token)
        stage = root['stage']
        import json
        if (self.last_park is None or self._confirmed_park_id == self.last_park.record_id
                or stage['status'] != 'verified' or stage['record'] != json.loads(json.dumps(self.last_park.to_dict()))
                or self.last_park.operation['connection_id'] != self.table.motor._connection_id):
            raise HardwareSafetyError('Confirm only the latest original verified loading corner, once.')
        if kind == 'load' and self.session.store.latest_operator_context(self.session.token) is not None:
            raise HardwareSafetyError('Initial tray loading cannot reset an already confirmed orientation.')
        directions = (copy.deepcopy(root['plan']['commands']['file_directions']) if kind == 'load' else self.directions())
        if any(type(value) is not bool for value in directions.values()):
            raise HardwareSafetyError('Original per-file tray orientations must be explicit booleans.')
        if kind == 'flip':
            if file_id not in directions:
                raise HardwareSafetyError('The tray flip must name its original registered file.')
            directions[file_id] = not directions[file_id]
        fields = self.coordinator.fields.verify_off(self.session)
        context = dict(schema='rapidpy.queue_tray_state.v1', queue_token=self.session.token,
            phase='confirmed', kind=kind, file_id=file_id, operator=operator.strip(),
            file_directions=directions, park_record_id=self.last_park.record_id)
        operation = dict(action='confirm_original_tray_' + kind, operator_context=context,
            connection_id=self.table.motor._connection_id, field_outputs_proof=dict(fields))
        token = self.session.child_store.begin('motion', operation, self.table.profile)
        samples, error = self.table._stopped()
        observations = []
        if not error:
            try:
                self.table._check_cancel()
                observations.append(self._corner(samples))
                if self.table.motor._connection_id != operation['connection_id']:
                    raise HardwareSafetyError('Original motor connection changed during operator confirmation.')
            except Exception as exc:
                error = str(exc) or type(exc).__name__
        if error:
            context['phase'] = 'unverified'
        record = QueueOperatorRecord(copy.deepcopy(operation), copy.deepcopy(self.table.profile),
            tuple(observations), tuple(dict(axis=key, samples=value) for key, value in samples.items()),
            error, '', not bool(error), operator_context=context)
        self.session.child_store.finish(token, self.table.profile, record)
        if error:
            raise HardwareSafetyError('Original tray confirmation remains pending: ' + error)
        self._confirmed_park_id = self.last_park.record_id
        return record
