"""A queue owns its outputs across individually recoverable native stages.

This journal is independent of measurement destinations. Immutable, linked stage
events survive crashes; completing a treatment never clears the queue's latch.
"""
from contextlib import contextmanager
from dataclasses import dataclass, field
from datetime import datetime, timezone
import hashlib
import json
import os
import threading
import uuid

from .hardware_safety import HardwareSafetyError, HardwareSafetyStore, _canonical

STAGE_FAMILIES = frozenset({'pulse', 'rrm', 'af', 'arm', 'motion', 'acquisition', 'vacuum'})
QUEUE_CHECKS = frozenset({'vacuum_off_verified', 'motors_stopped_verified', 'field_outputs_off_verified'})


@dataclass(frozen=True)
class QueueSafeStateRecord:
    vacuum_off_verified: bool
    motors_stopped_verified: bool
    field_outputs_off_verified: bool
    cleanup_errors: tuple[str, ...] = ()
    simulated: bool = False
    record_id: str = field(default_factory=lambda: uuid.uuid4().hex)

    @property
    def safe_state_confirmed(self):
        return (self.vacuum_off_verified is True and self.motors_stopped_verified is True
                and self.field_outputs_off_verified is True and not self.cleanup_errors
                and self.simulated is False)

    def to_dict(self):
        return dict(schema='rapidpy.queue_safe_state.v1', record_id=self.record_id,
            safe_state_confirmed=self.safe_state_confirmed, simulated=self.simulated,
            checks={name: getattr(self, name) for name in sorted(QUEUE_CHECKS)},
            cleanup_errors=list(self.cleanup_errors))


@dataclass(frozen=True)
class QueueVacuumHoldRecord:
    hold_acknowledged: bool
    raw_acknowledgements: tuple[str, ...]
    simulated: bool = False
    record_id: str = field(default_factory=lambda: uuid.uuid4().hex)
    safe_state_confirmed: bool = field(default=False, init=False)

    def to_dict(self):
        return dict(schema='rapidpy.queue_vacuum_hold.v1', record_id=self.record_id,
            hold_acknowledged=self.hold_acknowledged, raw_acknowledgements=list(self.raw_acknowledgements),
            safe_state_confirmed=False, simulated=self.simulated)


def _verified_hold(record):
    return (isinstance(record, dict) and record.get('schema') == 'rapidpy.queue_vacuum_hold.v1'
        and record.get('hold_acknowledged') is True and record.get('simulated') is False
        and record.get('safe_state_confirmed') is False
        and isinstance(record.get('raw_acknowledgements'), list)
        and bool(record['raw_acknowledgements'])
        and all(isinstance(reply, str) and reply.strip() for reply in record['raw_acknowledgements']))


def _verified_pump_ready(record):
    if not isinstance(record, dict):
        return False
    evidence = record.get('command_evidence')
    return (record.get('schema') == 'rapidpy.queue_vacuum_phase.v1'
        and record.get('operation') == 'pump_ready'
        and record.get('pump_enabled') is True and record.get('valve_connected') is False
        and record.get('output_state_acknowledged') is True
        and record.get('safe_state_confirmed') is False and record.get('simulated') is False
        and isinstance(evidence, list) and len(evidence) == 2
        and all(isinstance(item, dict) for item in evidence)
        and [item.get('command') for item in evidence] == ['10V00', '10MFF']
        and all(isinstance(item.get('reply'), str) and item['reply'].strip() for item in evidence))


def _verified_vacuum_stage(stage):
    action = stage['plan'].get('action')
    return (stage['family'] == 'vacuum' and (
        (action == 'enable' and _verified_hold(stage['record']))
        or (action == 'pump_ready' and _verified_pump_ready(stage['record']))))


def _validate_queue_record(record):
    if (not isinstance(record, dict) or record.get('schema') != 'rapidpy.queue_safe_state.v1'
            or not isinstance(record.get('checks'), dict) or set(record['checks']) != QUEUE_CHECKS
            or any(type(value) is not bool for value in record['checks'].values())
            or not isinstance(record.get('cleanup_errors'), list)
            or any(not isinstance(error, str) for error in record['cleanup_errors'])):
        raise ValueError('queue release lacks explicit independent output verification')
    _token(record['record_id'])
    expected_safe = all(record['checks'].values()) and not record['cleanup_errors'] and record.get('simulated') is False
    if record.get('safe_state_confirmed') is not expected_safe or type(record.get('simulated')) is not bool:
        raise ValueError('queue safe-state evidence is inconsistent')


def _snapshot(value):
    return json.loads(_canonical(value))


def _token(value):
    if not isinstance(value, str) or len(value) != 32:
        raise ValueError('invalid queue identity')
    int(value, 16)
    if value != value.lower() or any(c not in '0123456789abcdef' for c in value):
        raise ValueError('invalid queue identity')


def _head(value):
    if value is None:
        return
    _token(value['id'])
    digest = value['sha256']
    if not isinstance(digest, str) or len(digest) != 64 or any(c not in '0123456789abcdef' for c in digest):
        raise ValueError('invalid queue evidence digest')


def _record_snapshot(record):
    snapshot = _snapshot(record.to_dict())
    if (not isinstance(snapshot, dict)
            or type(snapshot.get('safe_state_confirmed')) is not bool
            or type(snapshot.get('simulated')) is not bool
            or snapshot['safe_state_confirmed'] is not record.safe_state_confirmed
            or snapshot['simulated'] is not record.simulated):
        raise HardwareSafetyError('Queue evidence flags must be consistent physical booleans.')
    return snapshot


def validate_queue_state(state):
    """Called by every reader, including helpers which cannot operate a queue."""
    _token(state['token'])
    bindings = state['profile']['stage_profiles']
    if (not isinstance(bindings, dict) or not bindings or set(bindings) - STAGE_FAMILIES
            or any(not isinstance(value, dict) for value in bindings.values())):
        raise ValueError('invalid queue stage bindings')
    count = state['stage_count']
    if isinstance(count, bool) or not isinstance(count, int) or count < 0:
        raise ValueError('invalid queue stage count')
    _head(state['history_head'])
    if state['status'] == 'verified':
        record = state['record']
        _validate_queue_record(record)
        if (not isinstance(record, dict) or record.get('safe_state_confirmed') is not True
                or record.get('simulated') is not False):
            raise ValueError('queue lacks verified physical evidence')
    stage = state['stage']
    if stage is None:
        if count != 0 or state['history_head'] is not None:
            raise ValueError('missing queue stage')
        return
    _token(stage['token'])
    if (stage['family'] not in bindings or stage['status'] not in {'pending', 'verified', 'held'}
            or _canonical(stage['profile']) != _canonical(bindings[stage['family']])
            or not isinstance(stage['plan'], dict) or not isinstance(stage['sample_id'], str)
            or not isinstance(stage['run_id'], str) or count == 0):
        raise ValueError('invalid queue stage')
    datetime.fromisoformat(stage['started_at'])
    if stage['status'] == 'verified':
        record = stage['record']
        if (not isinstance(record, dict) or record.get('safe_state_confirmed') is not True
                or record.get('simulated') is not False or state['history_head'] is None):
            raise ValueError('queue stage lacks verified physical evidence')
    if stage['status'] == 'held' and (not _verified_vacuum_stage(stage) or state['history_head'] is None):
        raise ValueError('queue vacuum hold lacks native acknowledgement evidence')
    if state['status'] == 'verified' and stage['status'] != 'verified':
        raise ValueError('queue was cleared with an unfinished stage')


class QueueSafetyStore(HardwareSafetyStore):
    def begin_queue(self, plan, profile, *, run_id=''):
        plan, profile = _snapshot(plan), _snapshot(profile)
        state = dict(schema=self.schema, status='pending', family='queue', token=uuid.uuid4().hex,
            plan=plan, profile=profile, sample_id='', run_id=run_id,
            started_at=datetime.now(timezone.utc).isoformat(), record=None,
            stage=None, stage_count=0, history_head=None)
        validate_queue_state(state)
        with self._locked():
            if self.pending():
                raise HardwareSafetyError('An unfinished hardware operation requires original-station recovery.')
            self._write(state)
        return state['token']

    def _queue(self, token, profile=None):
        state = self.pending(profile)
        if not state or state['family'] != 'queue' or state['token'] != token:
            raise HardwareSafetyError('Stale queue ownership token.')
        return state

    def begin_stage(self, token, family, plan, profile, *, sample_id='', run_id=''):
        plan, profile = _snapshot(plan), _snapshot(profile)
        with self._locked():
            state = self._queue(token)
            if state['stage'] and state['stage']['status'] == 'pending':
                raise HardwareSafetyError('The unfinished queue stage requires recovery before another stage can begin.')
            self.verify_history(state, latest_only=True)
            binding = state['profile']['stage_profiles'].get(family)
            if binding is None or _canonical(binding) != _canonical(profile):
                raise HardwareSafetyError('Queue stage does not match its original wiring/calibration binding.')
            stage_token = uuid.uuid4().hex
            state['stage'] = dict(token=stage_token, family=family, status='pending',
                plan=plan, profile=profile, sample_id=sample_id, run_id=run_id,
                started_at=datetime.now(timezone.utc).isoformat(), record=None)
            state['stage_count'] += 1
            validate_queue_state(state)
            self._write(state)
            return stage_token

    def _event_path(self, queue_token, event_id):
        _token(queue_token)
        _token(event_id)
        return self.path.parent / (self.path.name + '.queue-events') / queue_token / (event_id + '.json')

    def _publish_event(self, state, stage):
        event_id = uuid.uuid4().hex
        event = dict(schema='rapidpy.queue_stage.v1', id=event_id, queue_token=state['token'],
            stage_count=state['stage_count'], stage=stage, previous=state['history_head'])
        payload = _canonical(event)
        path = self._event_path(state['token'], event_id)
        try:
            path.parent.mkdir(parents=True, exist_ok=True)
            with path.open('xb') as handle:
                handle.write(payload)
                handle.flush()
                os.fsync(handle.fileno())
            if os.name != 'nt':
                descriptor = os.open(path.parent, os.O_RDONLY)
                try:
                    os.fsync(descriptor)
                finally:
                    os.close(descriptor)
        except OSError as exc:
            raise HardwareSafetyError('Queue stage evidence could not be persisted: ' + str(exc)) from exc
        return dict(id=event_id, sha256=hashlib.sha256(payload).hexdigest())

    def verify_history(self, state, *, latest_only=False):
        """Validate every linked event before declaring the entire queue safe."""
        head, seen = state['history_head'], set()
        expected_count = state['stage_count']
        if state['stage'] is not None and state['stage']['record'] is None:
            expected_count -= 1
        first = True
        while head is not None:
            try:
                _head(head)
                if head['id'] in seen:
                    raise ValueError('queue evidence cycle')
                seen.add(head['id'])
                payload = self._event_path(state['token'], head['id']).read_bytes()
                if hashlib.sha256(payload).hexdigest() != head['sha256']:
                    raise ValueError('queue evidence checksum mismatch')
                event = json.loads(payload)
                if (event['schema'] != 'rapidpy.queue_stage.v1' or event['id'] != head['id']
                        or event['queue_token'] != state['token']):
                    raise ValueError('queue evidence identity mismatch')
                count = event['stage_count']
                if type(count) is not int or count < 1 or count != expected_count:
                    raise ValueError('queue evidence stage count mismatch')
                event_state = dict(state, status='pending', stage=event['stage'], stage_count=count,
                    history_head=head)
                validate_queue_state(event_state)
                if first and state['stage']['record'] is not None and _canonical(event['stage']) != _canonical(state['stage']):
                    raise ValueError('latest queue evidence does not match the latched stage')
                if latest_only:
                    return 1
                first = False
                head = event['previous']
                if head is not None:
                    _head(head)
                    prior_payload = self._event_path(state['token'], head['id']).read_bytes()
                    prior_count = json.loads(prior_payload)['stage_count']
                    if prior_count not in {count, count - 1}:
                        raise ValueError('queue evidence stages are missing or reordered')
                    if prior_count == count - 1:
                        expected_count -= 1
            except (OSError, ValueError, KeyError, TypeError) as exc:
                raise HardwareSafetyError('Queue evidence cannot be verified: ' + str(exc)) from exc
        if expected_count not in {0, 1}:
            raise HardwareSafetyError('Queue evidence history is incomplete.')
        return len(seen)

    def finish_stage(self, token, stage_token, profile, record):
        snapshot = _record_snapshot(record)
        with self._locked():
            state = self._queue(token)
            stage = state['stage']
            if not stage or stage['status'] != 'pending' or stage['token'] != stage_token:
                raise HardwareSafetyError('Stale queue stage token.')
            if _canonical(stage['profile']) != _canonical(profile):
                raise HardwareSafetyError('Queue stage wiring/calibration changed; original-station recovery is required.')
            stage['record'] = snapshot
            if record.safe_state_confirmed is True and record.simulated is False:
                stage['status'] = 'verified'
            elif _verified_vacuum_stage(stage):
                stage['status'] = 'held'
            state['history_head'] = self._publish_event(state, stage)
            validate_queue_state(state)
            self._write(state)

    def finish_queue(self, token, profile, record):
        snapshot = _record_snapshot(record)
        try:
            _validate_queue_record(snapshot)
        except (ValueError, KeyError, TypeError) as exc:
            raise HardwareSafetyError(str(exc)) from exc
        with self._locked():
            state = self._queue(token, profile)
            if state['stage'] and state['stage']['status'] != 'verified':
                raise HardwareSafetyError('Queue cannot finish before its active stage is verified.')
            self.verify_history(state)
            state['record'] = snapshot
            if record.safe_state_confirmed is True and record.simulated is False:
                state['status'] = 'verified'
            validate_queue_state(state)
            self._write(state)


class QueueWorkflowSession:
    """An OS lease plus an exclusive worker capability for borrowed treatments."""
    def __init__(self, store, token, lease, *, recovering=False):
        self.store, self.token, self._lease = store, token, lease
        self.is_recovery = recovering
        self._lock = threading.RLock()
        self._owner = None
        self._depth = 0
        self.child_store = QueueStageStore(self)

    @classmethod
    def start(cls, store, plan, profile, *, run_id=''):
        lease = store.operation_lease()
        lease.__enter__()
        try:
            token = store.begin_queue(plan, profile, run_id=run_id)
            return cls(store, token, lease)
        except BaseException:
            lease.__exit__(None, None, None)
            raise

    @classmethod
    def recover(cls, store, profile):
        lease = store.operation_lease()
        lease.__enter__()
        try:
            state = store.pending(profile)
            if not state or state['family'] != 'queue':
                raise HardwareSafetyError('No matching unfinished queue exists.')
            store.verify_history(state)
            return cls(store, state['token'], lease, recovering=True)
        except BaseException:
            lease.__exit__(None, None, None)
            raise

    @contextmanager
    def claim(self):
        if not self._lock.acquire(blocking=False):
            raise HardwareSafetyError('Another queue worker still owns native I/O.')
        try:
            if self._lease is None:
                raise HardwareSafetyError('Queue ownership has already been released.')
            self.store._queue(self.token)
            self._owner = threading.get_ident()
            self._depth += 1
            try:
                yield self.child_store
            finally:
                self._depth -= 1
                if not self._depth:
                    self._owner = None
        finally:
            self._lock.release()

    def release(self):
        """Release only after the driver has stopped; the journal stays latched."""
        if not self._lock.acquire(blocking=False):
            raise HardwareSafetyError('Cannot release ownership while a queue worker is live.')
        try:
            if self._depth:
                raise HardwareSafetyError('Cannot release ownership inside a native queue claim.')
            if self._lease is not None:
                self._lease.__exit__(None, None, None)
                self._lease = None
        finally:
            self._lock.release()


class QueueStageStore:
    """Treatment store adapter; only the current claimed worker may borrow it."""
    def __init__(self, session):
        self.session = session

    def _owned(self):
        if self.session._lease is None or self.session._owner != threading.get_ident():
            raise HardwareSafetyError('A native queue worker claim is required before stage I/O.')

    @contextmanager
    def operation_lease(self):
        self._owned()
        with self.session.claim():
            yield

    def begin(self, family, plan, profile, *, sample_id='', run_id=''):
        self._owned()
        return self.session.store.begin_stage(self.session.token, family, plan, profile,
            sample_id=sample_id, run_id=run_id)

    def finish(self, token, profile, record):
        self._owned()
        self.session.store.finish_stage(self.session.token, token, profile, record)

    def pending(self, profile=None):
        state = self.session.store._queue(self.session.token)
        if self.session._owner != threading.get_ident():
            return state
        stage = state['stage']
        if not stage or stage['status'] in {'verified', 'held'}:
            return None
        if profile is not None and _canonical(profile) != _canonical(stage['profile']):
            raise HardwareSafetyError('Restore the original queue stage wiring/calibration before recovery.')
        return _snapshot(stage)
