"""Verified empty-rod shutdown and original-handle settlement for a native queue."""
import copy
from dataclasses import dataclass, field
import json

from rapidpy_common.hardware_safety import HardwareSafetyError
from rapidpy_common.queue_safety import QueueSafeStateRecord
from .queue_table_motion import QueueTableMoveRecord


@dataclass(frozen=True)
class QueueTerminalPoseProof:
    owner: object
    record: QueueTableMoveRecord

    def require(self, session, valve_connected):
        if valve_connected is not False or self.owner.session is not session:
            raise HardwareSafetyError('Terminal rod clearance can authorize vacuum OFF only.')
        self.owner._validate_open()
        root = session.store._queue(session.token)
        stage = root['stage']
        expected = json.loads(json.dumps(self.record.to_dict(), allow_nan=False))
        if (not stage or stage['status'] != 'verified' or stage['family'] != 'motion'
                or stage['record'] != expected or self.record.safe_state_confirmed is not True
                or self.record.operation.get('connection_id') != self.owner.table.motor._connection_id):
            raise HardwareSafetyError('The latest original verified terminal rod clearance is required.')
        return copy.deepcopy(self.record.operation)


@dataclass(frozen=True)
class QueueTerminalSafeStateRecord(QueueSafeStateRecord):
    settlement: dict = field(default_factory=dict)

    def to_dict(self):
        value = super().to_dict()
        value['settlement'] = copy.deepcopy(self.settlement)
        return value


class QueueTerminalCleanup:
    def __init__(self, coordinator, instruments):
        from .queue_transfer_coordinator import QueueTransferCoordinator
        if not isinstance(coordinator, QueueTransferCoordinator):
            raise HardwareSafetyError('Terminal cleanup requires the original native transfer coordinator.')
        self.coordinator = coordinator
        self.instruments = instruments
        self.session, self.table = coordinator.session, coordinator.table
        self._motor = self.table.motor
        self.vacuum, self.fields = coordinator.vacuum, coordinator.fields
        self.profile = copy.deepcopy(self.session.store._queue(self.session.token)['profile'])
        self._close_token = None
        self._close_plan = None
        self._vacuum_controller = None
        self._vacuum_binding = None
        self._motor_handles = None
        self._field_handles = (self.fields.af, self.fields.pulse.daq, self.fields.pulse.relays, self.fields.arm.controller)
        self.last_record = None
        self.last_root_record = None
        self.finished = False

    def _validate_original(self):
        self.session.child_store._owned()
        root = self.session.store._queue(self.session.token)
        if root['profile'] != self.profile or self.session.is_recovery:
            raise HardwareSafetyError('Terminal cleanup requires its original live queue and station profile.')
        if (self.table.motor is not self._motor or self.table._profile() != self.table.profile
                or self._field_handles != (self.fields.af, self.fields.pulse.daq, self.fields.pulse.relays, self.fields.arm.controller)):
            raise HardwareSafetyError('Restore the original motor calibration before terminal cleanup.')
        self.fields._validate(self.session)
        self.instruments.validate()
        self.session.store.verify_history(root)
        return root

    def _validate_open(self):
        self._validate_original()
        self.coordinator._validate()
        if self.vacuum.is_valve_connected() is not False:
            raise HardwareSafetyError('Return and release the original specimen before terminal cleanup.')
        specimen = self.coordinator.lift.context(self.session)
        if specimen is not None and specimen.phase != 'clear':
            raise HardwareSafetyError('Verify the original specimen return before terminal cleanup.')
        return specimen

    def _clearance(self, fields):
        specimen = self._validate_open()
        reference = self.coordinator.reference.require(self.session)
        operation = dict(action='verify_terminal_rod_clearance', connection_id=self.table.motor._connection_id,
            reference_proof=reference.to_dict(), field_outputs_proof=dict(fields),
            specimen_context=specimen.to_dict() if specimen is not None else None,
            holder_context=self.session.store.latest_holder_context(self.session.token))
        child = self.session.child_store
        token = child.begin('motion', operation, self.table.profile)
        samples, failure = self.table._stopped()
        error = failure
        if not error:
            try:
                self.table._clearance(samples)
                self.table._check_cancel()
                if self.table.motor._connection_id != operation['connection_id']:
                    raise HardwareSafetyError('Original motor connection changed during terminal clearance.')
                if self.vacuum.is_valve_connected() is not False or self.vacuum.is_pump_on() is not True:
                    raise HardwareSafetyError('Original pump/valve state changed during terminal clearance.')
            except Exception as exc:
                error = str(exc)
        record = QueueTableMoveRecord(copy.deepcopy(operation), copy.deepcopy(self.table.profile),
            tuple(dict(axis=key, samples=value) for key, value in samples.items()), (), error, '', not bool(error),
            schema='rapidpy.queue_terminal_pose.v1')
        child.finish(token, self.table.profile, record)
        if error:
            raise HardwareSafetyError('Terminal rod clearance remains pending: ' + error)
        return QueueTerminalPoseProof(self, record)

    def finish(self):
        if self.last_root_record is not None:
            state = self.session.store.read()
            if state and state['status'] == 'verified':
                return self._settle_verified_finish(state)
        if self._close_token is None:
            self._validate_open()
            fields = self.fields.verify_off(self.session)
            pose = self._clearance(fields)
            release = self.vacuum.queue_set_outputs(self.session, pump_enabled=False, valve_connected=False,
                transfer_pose=pose, field_outputs_off_verified=fields)
            root = self._validate_original()
            binding = self.vacuum._queue_binding
            controller = self.vacuum._pump_only_controller
            if (binding is None or binding.session is not self.session or binding.held
                    or self.vacuum.output_state_known is not True or controller._valve_connected is not False
                    or controller._motor_powered is not False or release.safe_state_confirmed is not True):
                raise HardwareSafetyError('Original native vacuum release is unverified before terminal close.')
            self._vacuum_controller, self._vacuum_binding = controller, binding
            self._motor_handles = dict(self._motor._connections)
            self._close_plan = json.loads(json.dumps(dict(action='queue_terminal_close', pose=pose.record.to_dict(),
                vacuum_release=release.to_dict(), vacuum_stage_token=root['stage']['token'],
                vacuum_event=copy.deepcopy(root['history_head']), field_outputs_proof=dict(fields),
                field_instance_id=self.fields.instance_id, instruments=self.instruments.settlement_plan()), allow_nan=False))
            self._close_token = self.session.child_store.begin('motion', self._close_plan, self.table.profile)
        return self._close_original_handles()

    def _close_original_handles(self):
        root = self._validate_original()
        stage = root['stage']
        if (not stage or stage['token'] != self._close_token or stage['family'] != 'motion'
                or stage['profile'] != self.table.profile or stage['plan'] != self._close_plan
                or self._close_plan['field_instance_id'] != self.fields.instance_id
                or self.vacuum._queue_binding is not self._vacuum_binding
                or self.vacuum._pump_only_controller is not self._vacuum_controller
                or self._vacuum_binding.held or self._vacuum_controller._valve_connected is not False
                or self._vacuum_controller._motor_powered is not False
                or self._motor._connections != {port: client for port, client in self._motor_handles.items()
                                               if client._serial is not None}):
            raise HardwareSafetyError('Only original terminal handles and immutable release evidence may be retried.')
        if stage['status'] == 'verified':
            self.instruments.require_settled()
            if (self.last_record is None or stage['record'] != json.loads(json.dumps(self.last_record.to_dict()))
                    or self._vacuum_controller._serial is not None or self.table.motor._connections
                    or any(client._serial is not None for client in self._motor_handles.values())):
                raise HardwareSafetyError('Original verified transport settlement changed before root publication.')
            return self._publish_root_finish()
        errors, observations = [], []
        for name, close in (('vacuum', self._vacuum_controller.disconnect), ('motors', self.table.motor.disconnect),
                           ('squid', lambda: self.instruments.close('squid')),
                           ('susceptibility', lambda: self.instruments.close('susceptibility'))):
            try:
                close()
                observations.append(dict(transport=name, closed=True))
            except Exception as exc:
                errors.append(name + ': ' + str(exc))
                observations.append(dict(transport=name, closed=False, error=str(exc)))
        if self._vacuum_controller._serial is not None:
            errors.append('The original vacuum handle has not settled.')
        if self.table.motor._connections:
            errors.append('Original motor handles remain unsettled.')
        if any(client._serial is not None for client in self._motor_handles.values()):
            errors.append('Original motor serial handles have not settled.')
        if not self.instruments.is_settled():
            errors.append('Original SQUID/susceptibility handles have not settled.')
        record = QueueTableMoveRecord(copy.deepcopy(self._close_plan), copy.deepcopy(self.table.profile),
            tuple(observations), (), '; '.join(errors), '', not bool(errors), schema='rapidpy.queue_terminal_close.v1')
        self.last_record = record
        self.session.child_store.finish(self._close_token, self.table.profile, record)
        if errors:
            raise HardwareSafetyError('Original terminal transport close remains pending: ' + '; '.join(errors))
        return self._publish_root_finish()

    def _publish_root_finish(self):
        self.instruments.require_settled()
        if self.last_root_record is None:
            self.last_root_record = QueueTerminalSafeStateRecord(True, True, True, settlement=dict(queue_token=self.session.token,
                terminal_stage_token=self._close_token, terminal_record_id=self.last_record.record_id,
                pose_record_id=self._close_plan['pose']['record_id'],
                vacuum_release_id=self._close_plan['vacuum_release']['record_id'],
                field_evidence_id=self._close_plan['field_outputs_proof']['evidence_id'],
                instruments=copy.deepcopy(self._close_plan['instruments'])))
        root_record = self.last_root_record
        self.session.store.finish_queue(self.session.token, self.profile, root_record)
        return self._settle_verified_finish(self.session.store.read())

    def _settle_verified_finish(self, state):
        self.session.child_store._owned()
        self.instruments.require_settled()
        if (state['family'] != 'queue' or state['token'] != self.session.token or state['status'] != 'verified'
                or state['profile'] != self.profile or self.last_root_record is None
                or state['record'] != self.last_root_record.to_dict()
                or self.vacuum._pump_only_controller is not self._vacuum_controller
                or self._vacuum_controller._serial is not None or self.table.motor._connections
                or self.table.motor is not self._motor
                or any(client._serial is not None for client in self._motor_handles.values())
                or self._vacuum_controller._valve_connected is not False
                or self._vacuum_controller._motor_powered is not False
                or self.fields._profile() != self.fields.profile
                or self._field_handles != (self.fields.af, self.fields.pulse.daq, self.fields.pulse.relays, self.fields.arm.controller)
                or self.fields.instance_id != self._close_plan['field_instance_id']):
            raise HardwareSafetyError('Exact original settled output/transport evidence is required before detaching.')
        self.session.store.verify_history(state)
        self.vacuum._queue_binding = None
        self.vacuum._hold_session = None
        self.vacuum.output_state_known = False
        self.finished = True
        return self.last_root_record
