"""Borrow the native vacuum transport under a claimed durable queue owner."""
import copy
from dataclasses import dataclass, asdict
from datetime import datetime, timezone
import math
import uuid

from rapidpy_common.hardware_safety import HardwareSafetyError
from rapidpy_common.queue_safety import QueueWorkflowSession
from rapidpy_common.vacuum_diagnostic_safety import vacuum_binding


@dataclass(frozen=True)
class QueueVacuumRecord:
    operation: str
    station_binding: dict
    raw_acknowledgements: tuple
    hold_acknowledged: bool
    safe_state_confirmed: bool
    error: str = ''
    simulated: bool = False
    record_id: str = ''
    timestamp_iso: str = ''
    pump_enabled: bool = False
    valve_connected: bool = False
    output_state_acknowledged: bool = False

    def to_dict(self):
        snapshot = asdict(self)
        snapshot['schema'] = ('rapidpy.queue_vacuum_hold.v1' if self.operation == 'enable'
            else 'rapidpy.queue_vacuum_phase.v1' if self.operation == 'pump_ready'
            else 'rapidpy.queue_vacuum_release.v1')
        # Retain exact commands/replies in addition to the hold schema's text.
        snapshot['command_evidence'] = list(snapshot['raw_acknowledgements'])
        snapshot['raw_acknowledgements'] = [f"{item['command']}: {item['reply']}"
            for item in self.raw_acknowledgements]
        return snapshot


def queue_vacuum_binding(adapter):
    threshold = float(adapter._cfg.warn_threshold)
    if not math.isfinite(threshold) or threshold < 0:
        raise HardwareSafetyError('Queue vacuum acceptance threshold must be finite and nonnegative.')
    if threshold > 0 and adapter.pressure_telemetry_available is not True:
        raise HardwareSafetyError('Queue vacuum requires pressure telemetry at a positive threshold. Explicit command-only operation uses threshold zero.')
    return dict(adapter._binding(), warn_threshold=threshold,
        pressure_telemetry=adapter.pressure_telemetry_available is True)


class QueueVacuumBinding:
    def __init__(self, adapter, session):
        self.adapter, self.session = adapter, session
        self.held = True  # An unqueried cached OFF value does not prove outputs off.
        self.error = ''
        self._connected_once = False
        self._recovery_required = False
        self._validate_owner()

    def _validate_owner(self):
        if not isinstance(self.session, QueueWorkflowSession):
            raise HardwareSafetyError('A durable native queue session is required.')
        self.session.child_store._owned()
        if self.session.store.path.resolve() != self.adapter._safety_store.path.resolve():
            raise HardwareSafetyError('Queue and vacuum must share the same fixed safety journal.')
        root = self.session.store._queue(self.session.token)
        expected = queue_vacuum_binding(self.adapter)
        if (root['profile'].get('helper') != 'rapid_main_queue'
                or root['profile']['stage_profiles'].get('vacuum') != expected
                or root['profile'].get('resources', {}).get('vacuum') != self.adapter._binding()):
            raise HardwareSafetyError('Restore the original queue vacuum port, baud and acceptance binding before I/O.')
        self.profile = copy.deepcopy(expected)
        return root

    def connect(self, factory):
        self._validate_owner()
        if self.adapter._hold_session and self.adapter._hold_session.active:
            raise HardwareSafetyError('Release the diagnostic vacuum hold before starting a queue owner.')
        controller = self.adapter._pump_only_controller
        if controller is not None and controller.is_connected:
            self._validate_controller(controller)
            if self.session.is_recovery or not self._connected_once:
                self.adapter.output_state_known = False
            self._connected_once = True
            return
        if controller is not None:
            self._validate_controller(controller)
            controller.disconnect()  # Retain/retry any failed original handle close.
        if self._connected_once:
            self._recovery_required = True
        controller = factory()
        self.adapter._pump_only_controller = controller
        self.adapter.output_state_known = False
        try:
            controller.connect(self.profile['port'], baudrate=self.profile['baud'])
            if controller.is_connected is not True:
                raise HardwareSafetyError('Queue vacuum transport did not connect.')
            self._validate_controller(controller)
        except BaseException:
            try:
                controller.disconnect()
            except Exception:
                pass
            raise
        self.adapter._pump_only_controller = controller
        self.adapter.output_state_known = False
        self._connected_once = True

    def _validate_controller(self, controller):
        if vacuum_binding(controller) != {key: self.profile[key] for key in ('port', 'baud')}:
            raise HardwareSafetyError('The actual vacuum transport differs from the original queue binding.')

    def set_outputs(self, *, pump_enabled, valve_connected, motors_stopped_verified=False,
                    specimen_secured=False, specimen_at_pickup_verified=False,
                    field_outputs_off_verified=False, transfer_pose=None):
        if type(pump_enabled) is not bool or type(valve_connected) is not bool:
            raise HardwareSafetyError('Queue pump and valve states must be explicit booleans.')
        if valve_connected and not pump_enabled:
            raise HardwareSafetyError('Queue grip requires the vacuum pump powered.')
        if transfer_pose is not None:
            from .queue_lift_transfer import QueueVacuumPoseProof
            from .queue_terminal import QueueTerminalPoseProof
            from .queue_startup import QueueEmptyRodStopProof
            if not isinstance(transfer_pose, (QueueVacuumPoseProof, QueueTerminalPoseProof, QueueEmptyRodStopProof)):
                raise HardwareSafetyError('Native original specimen pose proof is required.')
            transfer_pose.require(self.session, valve_connected)
            motors_stopped_verified = True
            specimen_at_pickup_verified = valve_connected
            specimen_secured = not valve_connected
        if valve_connected:
            from .queue_field_outputs import field_outputs_proof
            field_outputs_proof(self.session, field_outputs_off_verified)
        if valve_connected and not (motors_stopped_verified is True
                and specimen_at_pickup_verified is True):
            raise HardwareSafetyError('Verify motor stop, specimen pickup position and field outputs off before connecting grip.')
        return self._set_enabled(valve_connected, pump_ready=pump_enabled and not valve_connected,
            motors_stopped_verified=motors_stopped_verified, specimen_secured=specimen_secured,
            field_outputs_off_verified=field_outputs_off_verified, transfer_pose=transfer_pose)

    def set_enabled(self, enabled, *, motors_stopped_verified=False, specimen_secured=False,
                    field_outputs_off_verified=False):
        return self._set_enabled(enabled, motors_stopped_verified=motors_stopped_verified,
            specimen_secured=specimen_secured, field_outputs_off_verified=field_outputs_off_verified)

    def _set_enabled(self, enabled, *, pump_ready=False, motors_stopped_verified=False,
                     specimen_secured=False, field_outputs_off_verified=False, transfer_pose=None):
        self._validate_owner()
        controller = self.adapter._require_controller()
        self._validate_controller(controller)
        if type(enabled) is not bool:
            raise HardwareSafetyError('Queue vacuum command must be an explicit boolean.')
        if (enabled or pump_ready) and (self.session.is_recovery or self._recovery_required):
            raise HardwareSafetyError('Queue restart recovery may release outputs but never replay vacuum enable.')
        field_proof = None
        if not enabled or field_outputs_off_verified is not False:
            from .queue_field_outputs import field_outputs_proof
            field_proof = field_outputs_proof(self.session, field_outputs_off_verified)
        if not enabled and not (motors_stopped_verified is True and specimen_secured is True):
            raise HardwareSafetyError('Verify motor stop, specimen support and field outputs off before queue vacuum release.')
        child = self.session.child_store
        pending = child.pending()
        if pending:
            if enabled or pump_ready or pending['family'] != 'vacuum' or pending['profile'] != self.profile:
                raise HardwareSafetyError('Recover the unfinished original queue stage before changing vacuum outputs.')
            token = pending['token']  # An explicit OFF request recovers; never replay ON.
        else:
            token = child.begin('vacuum', {'action': 'pump_ready' if pump_ready else 'enable' if enabled else 'release',
                                         'field_outputs_proof': field_proof,
                                         'transfer_pose': transfer_pose.record.to_dict() if transfer_pose else None}, self.profile)
        self.held = True
        self.adapter.output_state_known = False
        reset = getattr(controller, 'reset_acknowledgements', None)
        if not callable(reset):
            raise HardwareSafetyError('Native queue vacuum requires per-command acknowledgement evidence.')
        reset()
        error = ''
        try:
            if pump_ready:
                controller.set_valve_connect(False)
                controller.set_motor_power(True)
            else:
                controller.set_enabled(enabled)
            expected_commands = ['10V00', '10MFF'] if pump_ready else ['10MFF', '10VFF'] if enabled else ['10V00', '10M00']
            acknowledgements = controller.acknowledgements
            if ([item.get('command') for item in acknowledgements] != expected_commands
                    or any(not isinstance(item.get('reply'), str) or not item['reply'].strip() for item in acknowledgements)):
                raise HardwareSafetyError('Queue vacuum lacks both exact native command acknowledgements.')
            if enabled:
                if controller.is_enabled is not True:
                    raise HardwareSafetyError('Queue vacuum hold state is unverified.')
            elif controller._valve_connected is not False or controller._motor_powered is not pump_ready:
                raise HardwareSafetyError('The requested queue valve and pump states must both be acknowledged.')
            if enabled and self.profile['warn_threshold'] > 0:
                pressure = float(self.adapter.read_pressure())
                if not math.isfinite(pressure) or pressure < 0 or pressure > self.profile['warn_threshold']:
                    raise HardwareSafetyError('Queue vacuum pressure is outside its original acceptance threshold.')
        except Exception as exc:
            error = str(exc) or type(exc).__name__
        record = QueueVacuumRecord('pump_ready' if pump_ready else 'enable' if enabled else 'release', copy.deepcopy(self.profile),
            tuple(copy.deepcopy(controller.acknowledgements)), enabled and not error,
            not enabled and not pump_ready and not error, error, record_id=uuid.uuid4().hex,
            timestamp_iso=datetime.now(timezone.utc).isoformat(), pump_enabled=enabled or pump_ready,
            valve_connected=enabled, output_state_acknowledged=not bool(error))
        try:
            child.finish(token, self.profile, record)
        except Exception:
            self.error = 'Queue vacuum evidence could not be persisted; the original stage remains pending.'
            raise
        if error:
            self.error = error
            raise HardwareSafetyError('Queue vacuum state remains unverified: ' + error)
        self.held = enabled or pump_ready
        self.error = ''
        self.adapter.output_state_known = True
        return record

    def detach(self):
        self._validate_owner()
        controller = self.adapter._pump_only_controller
        if controller is None:
            raise HardwareSafetyError('The original queue vacuum transport handle is unavailable.')
        self._validate_controller(controller)
        root = self.session.store._queue(self.session.token)
        stage = root['stage']
        if (self.held or not self.adapter.output_state_known or child_pending(self.session)
                or not stage or stage['family'] != 'vacuum' or stage['status'] != 'verified'
                or stage['record'].get('schema') != 'rapidpy.queue_vacuum_release.v1'
                or stage['record'].get('safe_state_confirmed') is not True):
            raise HardwareSafetyError('Persist a verified original queue vacuum release before disconnecting its transport.')
        controller.disconnect()
        self.adapter.output_state_known = False


def child_pending(session):
    return session.child_store.pending() is not None
