"""Recover explicitly supported native stages after the original worker is gone."""
from dataclasses import asdict, dataclass, field, replace
from datetime import datetime, timezone
import json
import uuid

from rapidpy_common.hardware_safety import HardwareSafetyError, _canonical
from rapidpy_common.queue_safety import QueueSafetyStore, QueueWorkflowSession
from .queue_startup import QueueNativeStartup
from .queue_terminal import QueueTerminalSafeStateRecord


def restart_stage(state):
    """Classify recorded evidence. The operation lease is still required to act."""
    if (not state or state['family'] != 'queue' or state['status'] != 'pending'
            or state['profile'].get('helper') != 'rapid_main_queue'):
        raise HardwareSafetyError('No interrupted native queue is recorded.')
    stage = state['stage']
    if stage is not None:
        action = stage['plan'].get('action')
        if (stage['family'] == 'field_outputs' and stage['status'] == 'pending'
                and action == 'verify_participating_field_outputs_off'):
            return 'field_outputs'
        if stage['family'] == 'motion' and stage['status'] == 'verified' and action == 'queue_terminal_close':
            return 'terminal_close'
        if (stage['status'] == 'pending' and action != 'queue_terminal_close'
                and stage['family'] in {'motion', 'acquisition', 'af', 'arm', 'pulse', 'rrm', 'vacuum'}):
            return 'containment'
    description = (f"{stage['family']} / {stage['plan'].get('action', 'unknown')} ({stage['status']})"
                   if stage is not None else 'startup before the first stage')
    raise HardwareSafetyError(f'Original stage: {description}. Station recovery is required. '
                              'Specimen support and output states cannot be inferred from this record.')


def _safe_record(stage, family, schema):
    record = stage['record']
    if (stage['status'] != 'verified' or stage['family'] != family or record['schema'] != schema
            or record['safe_state_confirmed'] is not True or record['simulated'] is not False):
        raise ValueError(f'Original {schema} evidence is incomplete.')
    return record


def _terminal_record(store, state, operator):
    """Reconstruct completion from the original linked shutdown operations."""
    try:
        stage, token = state['stage'], state['token']
        plan = stage['plan']
        close = _safe_record(stage, 'motion', 'rapidpy.queue_terminal_close.v1')
        observed = close['observations']
        if (close['operation'] != plan or close['station_profile'] != stage['profile']
                or len(observed) != 4 or {item['transport'] for item in observed}
                != {'motors', 'vacuum', 'squid', 'susceptibility'}
                or any(item['closed'] is not True for item in observed)
                or close['error'] or close['cleanup_error']):
            raise ValueError('Original transport settlement is incomplete.')
        close_event = store.stage_event(token, state['history_head'])
        # A failed close can add an earlier event for this same stage token.
        # Walk those attempts without skipping any distinct intervening stage.
        link = close_event['previous']
        while link != plan['vacuum_event']:
            attempt = store.stage_event(token, link)
            if attempt['stage']['token'] != stage['token'] or attempt['stage']['plan'] != plan:
                raise ValueError('Terminal close does not follow the original vacuum release.')
            link = attempt['previous']
        vacuum_event = store.stage_event(token, link)
        vacuum_stage = vacuum_event['stage']
        vacuum = _safe_record(vacuum_stage, 'vacuum', 'rapidpy.queue_vacuum_release.v1')
        commands = vacuum['command_evidence']
        if (vacuum_stage['token'] != plan['vacuum_stage_token'] or vacuum != plan['vacuum_release']
                or vacuum['station_binding'] != vacuum_stage['profile']
                or vacuum['operation'] != 'release' or vacuum_stage['plan']['action'] != 'release'
                or vacuum['pump_enabled'] is not False or vacuum['valve_connected'] is not False
                or vacuum['output_state_acknowledged'] is not True or vacuum['error']
                or [entry['command'] for entry in commands] != ['10V00', '10M00']
                or any(not isinstance(entry['reply'], str) or not entry['reply'].strip() for entry in commands)):
            raise ValueError('Original valve and pump OFF acknowledgements are incomplete.')
        pose_event = store.stage_event(token, vacuum_event['previous'])
        pose_stage = pose_event['stage']
        pose = _safe_record(pose_stage, 'motion', 'rapidpy.queue_terminal_pose.v1')
        operation = pose['operation']
        axes = set(pose_stage['profile']['axes'])
        observations = {item['axis']: item['samples'] for item in pose['observations']}
        if (pose != plan['pose'] or vacuum_stage['plan']['transfer_pose'] != pose
                or operation['action'] != 'verify_terminal_rod_clearance'
                or pose['station_profile'] != pose_stage['profile']
                or pose['error'] or pose['cleanup_error'] or set(observations) != axes
                or len(pose['observations']) != len(axes)
                or any(len(samples) != 2 or any(type(sample['velocity_register']) is not int
                       or sample['velocity_register'] != 0 or type(sample['position_raw']) is not int
                       for sample in samples) or samples[0]['position_raw'] != samples[1]['position_raw']
                       for samples in observations.values())):
            raise ValueError('Original stopped-axis and empty-rod clearance evidence is incomplete.')
        field_stage = store.stage_event(token, pose_event['previous'])['stage']
        field = _safe_record(field_stage, 'field_outputs', 'rapidpy.queue_field_outputs.v1')
        proof = plan['field_outputs_proof']
        if (field['field_context'] != proof or operation['field_outputs_proof'] != proof
                or vacuum_stage['plan']['field_outputs_proof'] != proof or proof['phase'] != 'off'
                or proof['queue_token'] != token or proof['instance_id'] != plan['field_instance_id']
                or field['station_profile'] != field_stage['profile'] or field['errors']
                or any(field[name]['safe_state_confirmed'] is not True for name in ('af', 'arm', 'pulse'))
                or plan['instruments']['profile'] != state['profile']['resources']['instruments']):
            raise ValueError('Original field cutoff or instrument bindings changed.')
    except (KeyError, TypeError, ValueError, AttributeError) as exc:
        raise HardwareSafetyError('Recorded terminal shutdown cannot be settled: ' + str(exc)) from exc
    return QueueTerminalSafeStateRecord(True, True, True, settlement=dict(
        queue_token=token, terminal_stage_token=stage['token'], terminal_record_id=close['record_id'],
        pose_record_id=pose['record_id'], vacuum_release_id=vacuum['record_id'],
        field_evidence_id=proof['evidence_id'], instruments=plan['instruments'],
        restart_publication=dict(operator=operator, timestamp_iso=datetime.now(timezone.utc).isoformat(),
                                 original_close_event=state['history_head'])))


class QueueRestartRecovery:
    def __init__(self, backend, vacuum, expected_state, *, operator):
        self.kind = restart_stage(expected_state)
        if (backend._config.general.nocomm is not False or not isinstance(operator, str) or not operator.strip()
                or getattr(backend._safety_store, 'session', None) is not None):
            raise HardwareSafetyError('Restart recovery requires the idle native station and a named operator.')
        self._expected = _canonical(expected_state)
        self._backend, self._vacuum, self._operator = backend, vacuum, operator.strip()
        store = QueueSafetyStore(backend._safety_store.path)
        self.session = QueueWorkflowSession.recover(store, expected_state['profile'])
        try:
            self._state()
        except BaseException:
            self.session.release()
            raise

    def _state(self):
        state = self.session.store._queue(self.session.token)
        if _canonical(state) != self._expected:
            raise HardwareSafetyError('The original queue changed after recovery was selected.')
        self.session.store.verify_history(state)
        return state

    def release(self):
        self.session.release()

    def run(self):
        with self.session.claim():
            state = self._state()
            stop_owner = ((state['token'], state['stage']['token']) if self.kind == 'containment' else None)
            table, fields, _, profile = QueueNativeStartup.station_bindings(self._backend, self._vacuum,
                                                                             stop_owner=stop_owner)
            # Transport ownership IDs belong to their original process. Every
            # actual setting and participating circuit must still match.
            expected = json.loads(_canonical(state['profile']))
            profile['resources']['instruments']['instance_id'] = expected['resources']['instruments']['instance_id']
            if _canonical(profile) != _canonical(expected):
                raise HardwareSafetyError('Restore the exact original native station configuration before recovery.')
            if self.kind == 'field_outputs':
                fields.verify_off(self.session, recovery=True, recovery_operator=self._operator)
                return
            if self.kind == 'containment':
                self._contain(state, table, fields, stop_owner)
                return
            record = _terminal_record(self.session.store, state, self._operator)
            self.session.store.finish_queue(self.session.token, state['profile'], record)

    def _contain(self, state, table, fields, stop_owner):
        stage = state['stage']
        intent = QueueContainmentRecord(stage['token'], stage['family'], self._operator,
                                        state['history_head'])
        child = self.session.child_store
        # This records intent ahead of even a serial open or an output cutoff.
        # The original plan stays latched and earlier evidence stays linked.
        child.finish(stage['token'], stage['profile'], intent)
        errors = []
        field_record = None
        try:
            field_record = fields.cutoff_for_containment(self.session, intent)
            if not field_record.safe_state_confirmed:
                errors.extend(field_record.errors)
        except Exception as exc:
            errors.append(f'Field cutoff: {exc}')

        samples = {}
        motor_errors = []
        stopped = False
        closed = False
        motor = table.motor
        try:
            available, open_errors = motor.connect_for_stop(stop_owner, baudrate=self._backend._config.motor_station.baud)
            motor_errors.extend(open_errors)
            samples, failure = table._stopped(available_ports=available)
            if failure:
                motor_errors.append(failure)
            stopped = not motor_errors and set(samples) == set(table.axes)
        except Exception as exc:
            motor_errors.append(f'Motor stop: {exc}')
        finally:
            try:
                motor.disconnect()
                closed = not motor._connections
                if not closed:
                    motor_errors.append('Retained recovery motor connections are still open.')
            except Exception as exc:
                motor_errors.append(f'Motor close: {exc}')

        errors.extend(motor_errors)
        result = replace(intent, status='failed' if errors else 'contained',
                         completed_iso=datetime.now(timezone.utc).isoformat(),
                         field_record=field_record.to_dict() if field_record is not None else None,
                         fields_off_verified=field_record is not None and field_record.safe_state_confirmed,
                         motor_observations=samples, motors_stopped_verified=stopped,
                         motor_handles_closed=closed, errors=tuple(errors))
        child.finish(stage['token'], stage['profile'], result)
        if errors:
            raise HardwareSafetyError('Restart containment is incomplete: ' + '; '.join(errors))


@dataclass(frozen=True)
class QueueContainmentRecord:
    stage_token: str
    stage_family: str
    operator: str
    interrupted_history_head: dict | None
    record_id: str = field(default_factory=lambda: uuid.uuid4().hex)
    started_iso: str = field(default_factory=lambda: datetime.now(timezone.utc).isoformat())
    completed_iso: str = ''
    status: str = 'started'
    field_record: dict | None = None
    fields_off_verified: bool = False
    motor_observations: dict = field(default_factory=dict)
    motors_stopped_verified: bool = False
    motor_handles_closed: bool = False
    errors: tuple[str, ...] = ()
    safe_state_confirmed: bool = field(default=False, init=False)
    simulated: bool = field(default=False, init=False)

    def to_dict(self):
        # Component checks do not establish specimen support, vacuum state or
        # any accepted treatment / scientific acquisition.
        return dict(schema='rapidpy.queue_restart_containment.v1', **asdict(self))
