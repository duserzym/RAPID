"""Journal native empty-rod startup before opening the station's transports."""
import copy
from dataclasses import asdict
from datetime import datetime, timezone
import json

from rapidpy_common.hardware_safety import HardwareSafetyError
from rapidpy_common.queue_safety import QueueSafetyStore, QueueWorkflowSession
from .queue_field_outputs import QueueFieldOutputs, field_outputs_proof
from .queue_lift_transfer import QueueLiftTransfer
from .queue_station import QueueStationGeometry
from .queue_table_motion import QueueTableMoveRecord, QueueXYTableMotion
from .queue_transfer_coordinator import QueueTransferCoordinator
from .queue_xy_reference import QueueXYReference


class QueueEmptyRodStopProof:
    def __init__(self, owner, record):
        self.owner, self.record = owner, record

    def require(self, session, valve_connected):
        owner = self.owner
        owner._validate()
        stage = session.store._queue(session.token)['stage']
        if (session is not owner.session or valve_connected is not False
                or stage['status'] != 'verified' or stage['family'] != 'motion'
                or stage['record'] != json.loads(json.dumps(self.record.to_dict()))
                or self.record.safe_state_confirmed is not True
                or self.record.operation['connection_id'] != owner.table.motor._connection_id):
            raise HardwareSafetyError('Original empty-rod stop proof permits valve OFF only.')
        return copy.deepcopy(self.record.operation)


class QueueNativeStartup:
    """Original live startup only; an interrupted startup cannot replay its motion."""
    @classmethod
    def prepare(cls, backend, vacuum, plan, *, run_id, operator, empty_rod_confirmed=False):
        from rapidpy_common.motor_routing import RoutedMotorSerialClient
        from .diagnostic_services import VacuumBackendAdapter
        if (backend._config.general.nocomm is not False or empty_rod_confirmed is not True
                or not isinstance(operator, str) or not operator.strip()
                or not isinstance(run_id, str) or not run_id.strip() or not isinstance(plan, dict)):
            raise HardwareSafetyError('Native startup requires a queue plan, run identity and named empty-rod confirmation.')
        if (not isinstance(backend._client, RoutedMotorSerialClient)
                or backend._client._connections or getattr(backend, '_queue_coordinator', None) is not None
                or getattr(backend, '_queue_startup', None) is not None
                or getattr(backend, '_queue_specimen_geometry', None) is not None
                or getattr(backend._safety_store, 'session', None) is not None):
            raise HardwareSafetyError('Restore the original idle native backend and close retained motor handles before startup.')
        if (not isinstance(vacuum, VacuumBackendAdapter) or vacuum.simulated is not False
                or vacuum._queue_binding is not None
                or (vacuum._hold_session is not None and vacuum._hold_session.active)
                or (vacuum._pump_only_controller is not None and vacuum._pump_only_controller._serial is not None)
                or vacuum._safety_store.path.resolve() != backend._safety_store.path.resolve()):
            raise HardwareSafetyError('Startup requires the idle original native vacuum and the same fixed safety journal.')
        geometry = QueueStationGeometry.from_config(backend._config, use_xy_table=True)
        from .hardware_contracts import _build_motor_controller_config
        station = backend._config.motor_station
        if (asdict(backend._client.config) != asdict(_build_motor_controller_config(backend._config))
                or station.ports != {key: axis.port for key, axis in backend._axes.items()}
                or station.addresses != {key: axis.address for key, axis in backend._axes.items()}
                or station.baud != 57600 or asdict(backend._config.vacuum) != asdict(vacuum._cfg)
                or getattr(backend, '_backend_errors', ())):
            raise HardwareSafetyError('Restore the accepted native motor/vacuum calibration and available components before startup.')
        table = QueueXYTableMotion(backend._client, backend._axes, geometry,
            should_cancel=lambda: getattr(backend, '_halt_check', None) is not None and backend._halt_check(),
            transfer_settings=dict(grip_settle_s=.3, dropoff_delay_s=backend._config.vacuum.dropoff_delay_s))
        pulse = getattr(backend, '_pulse_irm', None)
        if pulse is None or not callable(getattr(pulse, 'make_circuit', None)):
            raise HardwareSafetyError('Original participating pulse cutoff circuitry is required before startup.')
        fields = QueueFieldOutputs(getattr(backend._af_demag, '_controller', None),
            pulse.make_circuit(), backend._arm_bias)
        if (asdict(fields.pulse.cfg) != asdict(backend._config.pulse_irm)
                or asdict(fields.arm.cfg) != asdict(backend._config.irm_arm)):
            raise HardwareSafetyError('Restore the original accepted ARM and pulse configuration before startup.')
        profile = dict(helper='rapid_main_queue', resources=dict(vacuum=vacuum._binding()),
            stage_profiles=dict(motion=table.profile, vacuum=vacuum.queue_station_binding(),
                field_outputs=fields.profile, acquisition=backend._acquisition_safety_profile(),
                **{family: backend._safety_profile() for family in ('af', 'arm', 'pulse', 'rrm')}))
        attestation = dict(schema='rapidpy.queue_empty_rod_attestation.v1', operator=operator.strip(),
            empty_rod_confirmed=True, timestamp_iso=datetime.now(timezone.utc).isoformat())
        original_plan = json.loads(json.dumps(dict(commands=plan, startup=attestation), allow_nan=False))
        store = QueueSafetyStore(backend._safety_store.path)
        session = QueueWorkflowSession.start(store, original_plan, profile, run_id=run_id)
        instance = cls(backend, vacuum, table, fields, session, attestation)
        backend._queue_startup = instance
        backend._safety_store = session.child_store
        return instance

    def __init__(self, backend, vacuum, table, fields, session, attestation):
        self.backend, self.vacuum, self.table, self.fields = backend, vacuum, table, fields
        self.session, self.attestation = session, copy.deepcopy(attestation)
        self.completed = False
        self.started = False
        self.last_record = None

    def _validate(self):
        self.session.child_store._owned()
        root = self.session.store._queue(self.session.token)
        if (self.session.is_recovery or self.backend._queue_startup is not self
                or self.backend._safety_store is not self.session.child_store
                or self.backend._client is not self.table.motor
                or self.backend._axes != self.table.axes or self.table._profile() != self.table.profile
                or root['plan']['startup'] != self.attestation
                or self.backend._config.general.nocomm is not False):
            raise HardwareSafetyError('Restore the original live startup bindings and empty-rod attestation.')
        # Scientific and circuit bindings must still match the prepared snapshot.
        if (root['profile']['stage_profiles']['acquisition'] != self.backend._acquisition_safety_profile()
                or any(root['profile']['stage_profiles'][family] != self.backend._safety_profile()
                       for family in ('af', 'arm', 'pulse', 'rrm'))
                or self.fields.af is not getattr(self.backend._af_demag, '_controller', None)
                or self.fields.arm is not self.backend._arm_bias
                or self.fields.pulse.daq is not self.backend._pulse_irm.daq
                or self.fields.pulse.relays is not self.backend._pulse_irm.relays):
            raise HardwareSafetyError('Prepared startup scientific or field circuit bindings changed.')
        self.fields._validate(self.session)

    def _publish(self, token, operation, observations, cleanup, error, cleanup_error, schema):
        record = QueueTableMoveRecord(copy.deepcopy(operation), copy.deepcopy(self.table.profile),
            tuple(observations), tuple(dict(axis=axis, samples=samples) for axis, samples in cleanup.items()),
            error, cleanup_error, not bool(error or cleanup_error), schema=schema)
        self.last_record = record
        self.session.child_store.finish(token, self.table.profile, record)
        if not record.safe_state_confirmed:
            raise HardwareSafetyError('Native startup remains pending: ' + '; '.join(filter(None, (error, cleanup_error))))
        return record

    def _connect_and_stop(self, fields):
        self._validate()
        operation = dict(action='connect_empty_station_and_verify_stop', attestation=self.attestation,
            field_outputs_proof=field_outputs_proof(self.session, fields), connection_id=None)
        token = self.session.child_store.begin('motion', operation, self.table.profile)
        error, cleanup_error, observations, cleanup = '', '', [], {}
        try:
            self.table._check_cancel()
            self.table.motor.connect(baudrate=self.backend._config.motor_station.baud)
            self.table._validate(self.session)
            operation['connection_id'] = self.table.motor._connection_id
        except Exception as exc:
            error = str(exc) or type(exc).__name__
        if self.table.motor.is_connected:
            cleanup, cleanup_error = self.table._stopped()
        elif not error:
            error = 'All original motor transports must connect before startup stop verification.'
        self._validate()
        self.table._check_cancel()
        record = self._publish(token, operation, observations, cleanup, error, cleanup_error,
            'rapidpy.queue_startup_stop.v1')
        return QueueEmptyRodStopProof(self, record)

    def _home_empty_lift(self, fields):
        self._validate()
        self.table._validate(self.session)
        if (self.vacuum.output_state_known is not True or self.vacuum.is_valve_connected() is not False
                or self.vacuum.is_pump_on() is not True):
            raise HardwareSafetyError('Empty lift reference requires acknowledged original pump ON and valve OFF.')
        operation = dict(action='reference_confirmed_empty_lift', attestation=self.attestation,
            connection_id=self.table.motor._connection_id, field_outputs_proof=field_outputs_proof(self.session, fields))
        token = self.session.child_store.begin('motion', operation, self.table.profile)
        observations, error = [], ''
        motor, axis = self.table.motor, self.table.axes['updown']
        try:
            initial, failure = self.table._stopped()
            observations.append(dict(phase='before_reference', axes=initial))
            if failure:
                raise HardwareSafetyError(failure)
            self.table._check_cancel()
            top = motor.check_internal_status(axis, 4)
            if type(top) is not int or top not in (0, 1):
                raise HardwareSafetyError('The empty-lift top switch must return a binary value.')
            if top == 0:
                target = -2 * motor.config.meas_pos
                if type(target) is not int or not 0 < target < 2**31:
                    raise HardwareSafetyError('Accepted empty-lift homing travel must be positive signed controller counts.')
                speed = int(.25 * (motor.config.lift_speed_normal + 3 * motor.config.lift_speed_slow))
                motor.move_motor(axis, target, speed, wait_for_stop=True, stop_enable=-1, stop_condition=1)
                self.table._check_cancel()
                stopped, failure = self.table._stopped()
                observations.append(dict(phase='at_top_edge', axes=stopped))
                if failure:
                    raise HardwareSafetyError(failure)
            if motor.check_internal_status(axis, 4) != 1:
                raise HardwareSafetyError('Empty-lift reference did not reach the original live top switch.')
            self.table._check_cancel()
            motor.zero_target_pos(axis)
        except Exception as exc:
            error = str(exc) or type(exc).__name__
        cleanup, cleanup_error = self.table._stopped()
        if not error and not cleanup_error:
            try:
                self._validate()
                self.table._check_cancel()
                self.table._clearance(cleanup)
                if (motor._connection_id != operation['connection_id']
                        or any(abs(sample['position_raw']) > 150 for sample in cleanup['updown'])
                        or self.vacuum.is_valve_connected() is not False):
                    raise HardwareSafetyError('Original empty lift zero, top switch and valve OFF must remain verified.')
            except Exception as exc:
                error = str(exc) or type(exc).__name__
        return self._publish(token, operation, observations, cleanup, error, cleanup_error,
            'rapidpy.queue_empty_lift_reference.v1')

    def run(self):
        self._validate()
        if self.started:
            raise HardwareSafetyError('An original startup cannot replay; recover its unfinished stage explicitly.')
        self.started = True
        fields = self.fields.verify_off(self.session)
        pose = self._connect_and_stop(fields)
        self.vacuum.queue_connect(self.session)
        self.vacuum.queue_set_outputs(self.session, pump_enabled=True, valve_connected=False,
            transfer_pose=pose, field_outputs_off_verified=fields)
        self._home_empty_lift(fields)
        reference = QueueXYReference(self.table, self.vacuum)
        reference.home(self.session, empty_rod_confirmed=True, field_outputs_off_verified=fields)
        lift = QueueLiftTransfer(self.table, self.vacuum)
        coordinator = QueueTransferCoordinator(self.session, lift, reference, self.fields)
        self.backend.bind_queue_coordinator(coordinator)
        self.backend._connected = True
        self.completed = True
        return coordinator
