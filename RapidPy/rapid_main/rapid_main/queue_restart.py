"""Recover explicitly supported native stages after the original worker is gone."""
from dataclasses import asdict, dataclass, field, replace
from datetime import datetime, timezone
import hashlib
import json
import uuid

from rapidpy_common.hardware_safety import HardwareSafetyError, _canonical
from rapidpy_common.queue_safety import QueueSafetyStore, QueueWorkflowSession
from .queue_startup import QueueNativeStartup
from .queue_terminal import QueueTerminalSafeStateRecord


def require_interrupted_native_queue(state):
    if (not state or state['family'] != 'queue' or state['status'] != 'pending'
            or state['profile'].get('helper') != 'rapid_main_queue'):
        raise HardwareSafetyError('No interrupted native queue is recorded.')


def restart_stage(state):
    """Classify recorded evidence. The operation lease is still required to act."""
    require_interrupted_native_queue(state)
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


@dataclass(frozen=True)
class RestartShutdownAttestation:
    queue_token: str
    snapshot_sha256: str
    operator: str
    specimen_secured: bool
    empty_rod: bool
    work_area_clear: bool
    recorded_iso: str = field(default_factory=lambda: datetime.now(timezone.utc).isoformat())

    @classmethod
    def confirm(cls, state, operator, *, specimen_secured, empty_rod, work_area_clear):
        require_interrupted_native_queue(state)
        if not isinstance(operator, str) or not operator.strip():
            raise HardwareSafetyError('A named operator must confirm restart shutdown.')
        result = cls(state['token'], hashlib.sha256(_canonical(state)).hexdigest(), operator.strip(),
                     specimen_secured, empty_rod, work_area_clear)
        result.require(state, operator)
        return result

    def require(self, state, operator):
        if (self.queue_token != state['token'] or self.snapshot_sha256 != hashlib.sha256(_canonical(state)).hexdigest()
                or self.operator != operator.strip() or self.specimen_secured is not True
                or self.empty_rod is not True or self.work_area_clear is not True):
            raise HardwareSafetyError('Confirm independent specimen support, an empty rod and a clear work area for this queue.')


class QueueRestartRecovery:
    def __init__(self, backend, vacuum, expected_state, *, operator, shutdown=None, live_startup=None):
        require_interrupted_native_queue(expected_state)
        if backend._config.general.nocomm is not False or not isinstance(operator, str) or not operator.strip():
            raise HardwareSafetyError('Queue recovery requires the native station and a named operator.')
        original_session = getattr(backend._safety_store, 'session', None)
        if live_startup is None and original_session is not None:
            raise HardwareSafetyError('Restart recovery requires the idle native station.')
        if live_startup is not None and (not isinstance(live_startup, QueueNativeStartup)
                or live_startup.backend is not backend or live_startup.vacuum is not vacuum
                or live_startup.session is not original_session or original_session.is_recovery or shutdown is None):
            raise HardwareSafetyError('Live shutdown requires the original queue owner and operator confirmation.')
        if shutdown is None:
            self.kind = restart_stage(expected_state)
        else:
            if type(shutdown) is not RestartShutdownAttestation:
                raise HardwareSafetyError('Explicit restart shutdown confirmation is required.')
            shutdown.require(expected_state, operator)
            self.kind = 'shutdown'
        self._shutdown = shutdown
        self._live_startup = live_startup
        self._expected = _canonical(expected_state)
        self._backend, self._vacuum, self._operator = backend, vacuum, operator.strip()
        if live_startup is None:
            store = QueueSafetyStore(backend._safety_store.path)
            self.session = QueueWorkflowSession.recover(store, expected_state['profile'])
        else:
            self.session = live_startup.session
        try:
            self._state()
        except BaseException:
            self.release()
            raise

    def _state(self):
        state = self.session.store._queue(self.session.token)
        if _canonical(state) != self._expected:
            raise HardwareSafetyError('The original queue changed after recovery was selected.')
        self.session.store.verify_history(state)
        return state

    def release(self):
        if self._live_startup is None:
            self.session.release()

    def run(self):
        with self.session.claim():
            state = self._state()
            if self._live_startup is not None:
                self._run_live_shutdown(state)
                return
            stopping = self.kind in {'containment', 'shutdown'}
            stage = state['stage']
            stop_owner = (state['token'], stage['token']) if stopping and stage is not None else None
            table, fields, _, profile = QueueNativeStartup.station_bindings(self._backend, self._vacuum,
                stop_owner=stop_owner, restart_vacuum_token=state['token'] if self.kind == 'shutdown' else None)
            # Transport ownership IDs belong to their original process. Every
            # actual setting and participating circuit must still match.
            expected = json.loads(_canonical(state['profile']))
            profile['resources']['instruments']['instance_id'] = expected['resources']['instruments']['instance_id']
            if _canonical(profile) != _canonical(expected):
                raise HardwareSafetyError('Restore the exact original native station configuration before recovery.')
            if self.kind == 'field_outputs':
                fields.verify_off(self.session, recovery=True, recovery_operator=self._operator)
                return
            if stopping:
                self._start_containment(state, table, fields)
                return
            record = _terminal_record(self.session.store, state, self._operator)
            self.session.store.finish_queue(self.session.token, state['profile'], record)

    def _run_live_shutdown(self, state):
        from .hardware_contracts import _build_motor_controller_config
        from .queue_station import QueueStationGeometry
        startup = self._live_startup
        backend = self._backend
        terminal = backend._queue_terminal_cleanup
        if terminal is not None and terminal._close_token is not None:
            raise HardwareSafetyError('Use Retry Queue Shutdown to settle the original journaled closes.')
        startup._validate()
        binding = self._vacuum._queue_binding
        controller = self._vacuum._pump_only_controller
        serial = getattr(controller, '_serial', None)
        vacuum_handles = (binding, controller, serial)
        if (state['profile']['stage_profiles']['motion'] != startup.table.profile
                or startup.table.geometry != QueueStationGeometry.from_config(
                    backend._config, use_xy_table=backend._config.motor_station.use_xy_table)
                or startup.table.motor.config != _build_motor_controller_config(backend._config)
                or not backend._config.motor_station.routing_matches(startup.table.axes)
                or backend._config.motor_station.baud != 57600
                or asdict(backend._config.vacuum) != asdict(self._vacuum._cfg)
                or self._vacuum.queue_station_binding() != state['profile']['stage_profiles']['vacuum']
                or (binding is not None and binding.session is not self.session)
                or (startup.vacuum_handles is None and (binding is not None or serial is not None))
                or (startup.vacuum_handles is not None and any(
                    old is not current for old, current in zip(startup.vacuum_handles, vacuum_handles)))
                or any(port not in startup.motor_handles or client is not startup.motor_handles[port][0]
                       or client._serial is not startup.motor_handles[port][1]
                       for port, client in startup.table.motor._connections.items())):
            raise HardwareSafetyError('Restore the original live station configuration and handles before shutdown.')
        self._start_containment(state, startup.table, startup.fields)

    def _start_containment(self, state, table, fields):
        stage = state['stage']
        if self.kind == 'shutdown' and (stage is None or stage['status'] != 'pending'):
            action = 'live_shutdown' if self._live_startup is not None else 'restart_shutdown'
            self.session.child_store.begin('motion', dict(action=action,
                interrupted_stage_token=stage['token'] if stage is not None else None), table.profile)
            state = self.session.store._queue(self.session.token)
        self._contain(state, table, fields, (state['token'], state['stage']['token']))

    def _contain(self, state, table, fields, stop_owner):
        stage = state['stage']
        startup = self._live_startup
        intent = QueueContainmentRecord(stage['token'], stage['family'], self._operator,
            state['history_head'], shutdown_attestation=asdict(self._shutdown) if self._shutdown else None,
            scope='live' if startup is not None else 'restart')
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
        clearance = False
        vacuum_record = None
        vacuum_closed = False
        instrument_closures = None
        motor = table.motor
        try:
            available, open_errors = motor.connect_for_stop(stop_owner, baudrate=self._backend._config.motor_station.baud,
                original_handles=startup.motor_handles if startup is not None else None)
            motor_errors.extend(open_errors)
            samples, failure = table._stopped(available_ports=available)
            if failure:
                motor_errors.append(failure)
            stopped = not motor_errors and set(samples) == set(table.axes)
            if self._shutdown is not None and stopped and not errors:
                try:
                    table._clearance(samples)
                    try:
                        self._vacuum.queue_connect(self.session)
                    finally:
                        if startup is not None:
                            startup.remember_vacuum_handles()
                    # Confirm that clearance and stationarity still hold on the
                    # live motor connection immediately before vacuum commands.
                    samples, error = table._stopped(available_ports=available)
                    if error:
                        stopped = False
                        raise HardwareSafetyError(error)
                    table._clearance(samples)
                    clearance = True
                    ready = replace(intent, status='release_ready', field_record=field_record.to_dict(),
                        fields_off_verified=True, motor_observations=samples, motors_stopped_verified=True,
                        clearance_verified=True)
                    child.finish(stage['token'], stage['profile'], ready)
                    vacuum_record = self._vacuum.queue_release_for_shutdown(self.session, ready)
                    if not vacuum_record.safe_state_confirmed:
                        errors.append('Vacuum OFF: ' + vacuum_record.error)
                    elif startup is not None:
                        instrument_closures = self._close_live_instruments(startup.instruments, ready, errors)
                except Exception as exc:
                    errors.append(f'Queue shutdown: {exc}')
        except Exception as exc:
            motor_errors.append(f'Motor stop: {exc}')
        finally:
            vacuum_binding = self._vacuum._queue_binding
            if (self._shutdown is not None and vacuum_binding is not None and vacuum_binding.session is self.session
                    and (startup is None or vacuum_record is not None and vacuum_record.safe_state_confirmed)):
                try:
                    self._vacuum.queue_close_shutdown_transport(self.session, vacuum_record)
                    vacuum_closed = True
                except Exception as exc:
                    errors.append(f'Vacuum close: {exc}')
                finally:
                    if startup is not None:
                        startup.remember_vacuum_handles()
            try:
                motor.disconnect()
                closed = not motor._connections
                if not closed:
                    motor_errors.append('Retained recovery motor connections are still open.')
            except Exception as exc:
                motor_errors.append(f'Motor close: {exc}')
            finally:
                if startup is not None:
                    startup.remember_motor_handles()

        errors.extend(motor_errors)
        aborted = (self._shutdown is not None and not errors and clearance and vacuum_closed
                   and vacuum_record is not None and vacuum_record.safe_state_confirmed is True)
        result = replace(intent, status='aborted' if aborted else 'failed' if errors else 'contained',
                         completed_iso=datetime.now(timezone.utc).isoformat(),
                         field_record=field_record.to_dict() if field_record is not None else None,
                         fields_off_verified=field_record is not None and field_record.safe_state_confirmed,
                         motor_observations=samples, motors_stopped_verified=stopped,
                         motor_handles_closed=closed, errors=tuple(errors), clearance_verified=clearance,
                         vacuum_release=vacuum_record.to_dict() if vacuum_record is not None else None,
                         vacuum_handle_closed=vacuum_closed, instrument_closures=instrument_closures)
        if aborted:
            self.session.store.finish_interrupted_queue(self.session.token, stage['token'], state['profile'], result)
            if startup is not None:
                self._backend._detach_queue_lifetime(self.session)
        else:
            child.finish(stage['token'], stage['profile'], result)
        if errors:
            raise HardwareSafetyError('Queue containment is incomplete: ' + '; '.join(errors))

    @staticmethod
    def _close_live_instruments(instruments, readiness, errors):
        for name in ('squid', 'susceptibility'):
            try:
                instruments.close_for_shutdown(name, readiness)
            except Exception as exc:
                errors.append(f'{name} close: {exc}')
        if not instruments.is_settled():
            errors.append('Original scientific instrument handles remain unsettled.')
        return dict(profile=instruments.profile, closed=dict(instruments.closed))


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
    shutdown_attestation: dict | None = None
    clearance_verified: bool = False
    vacuum_release: dict | None = None
    vacuum_handle_closed: bool = False
    scope: str = 'restart'
    instrument_closures: dict | None = None
    safe_state_confirmed: bool = field(default=False, init=False)
    simulated: bool = field(default=False, init=False)

    def to_dict(self):
        # Component checks do not establish specimen support, vacuum state or
        # any accepted treatment / scientific acquisition.
        return dict(schema='rapidpy.queue_restart_containment.v1', **asdict(self))
