"""A held vacuum and participating lift share one durable station lifetime."""
import copy
from contextlib import contextmanager
from dataclasses import dataclass, asdict
from datetime import datetime, timezone
import uuid

from .hardware_safety import HardwareSafetyError, HardwareSafetyStore, default_safety_path
from .motor_diagnostic_safety import verify_stopped_in_place
from .af_diagnostic_safety import publish_diagnostic_record


def vacuum_binding(vacuum, port=None, baud=None):
    return {'port': (vacuum._port if port is None else port).strip().upper(),
            'baud': vacuum._baud if baud is None else baud}


@dataclass(frozen=True)
class StationDiagnosticRecord:
    treatment_id: str
    operation: dict
    station_profile: dict
    timestamp_iso: str
    observations: tuple
    error: str
    cleanup_error: str
    outputs_held: bool
    safe_state_confirmed: bool
    vacuum_evidence_basis: str = 'terminated serial command acknowledgements; no pressure telemetry'
    simulated: bool = False
    schema: str = 'rapidpy.station.diagnostic.v1'

    def to_dict(self):
        return asdict(self)


class VacuumHoldSession:
    def __init__(self, vacuum, *, store=None, helper='updown_control'):
        if helper not in {'updown_control', 'rapid_main_vacuum'}:
            raise HardwareSafetyError('Unknown vacuum diagnostic owner.')
        self.helper = helper
        self.vacuum = vacuum
        self.store = store if store is not None else HardwareSafetyStore(default_safety_path())
        self.active = False
        self.faulted = False
        self.vacuum_unverified = False
        self.evidence_unverified = False
        self.profile = None
        self.token = None
        self._lease = None
        self._lift = None
        self._axis = None
        self.observations = []
        self._ack_start = 0
        self.operation = {'action': 'updown_vacuum_hold' if helper == 'updown_control' else 'rapid_main_vacuum_hold'}

    def _start_acknowledgement_log(self):
        reset = getattr(self.vacuum, 'reset_acknowledgements', None)
        if callable(reset):
            reset()
        self._ack_start = len(self.vacuum.acknowledgements)

    def _publish(self, *, error='', cleanup_error='', held=True, safe=False, recovery=False):
        acknowledgements = copy.deepcopy(self.vacuum.acknowledgements[self._ack_start:])
        record = StationDiagnosticRecord('station-' + uuid.uuid4().hex, self.operation, copy.deepcopy(self.profile),
            datetime.now(timezone.utc).isoformat(), tuple(copy.deepcopy(self.observations) +
                [{'vacuum_acknowledgements': acknowledgements}]), error, cleanup_error, held, safe,
            schema='rapidpy.station.diagnostic_recovery.v1' if recovery else 'rapidpy.station.diagnostic.v1')
        publish_diagnostic_record(self.store, record, family='station')
        self.store.finish(self.token, self.profile, record)
        self.evidence_unverified = False
        self.observations = [{'previous_record_id': record.treatment_id, 'operation_token': self.token}]
        return record

    def start(self, lift=None, *, enabled=True):
        if self.active:
            raise HardwareSafetyError('The vacuum station is already owned.')
        if not self.vacuum.is_connected:
            raise HardwareSafetyError('Connect the original vacuum controller first.')
        self._lease = self.store.operation_lease()
        self._lease.__enter__()
        try:
            self.profile = {'helper': self.helper, 'resources': {'vacuum': vacuum_binding(self.vacuum)}}
            self.token = self.store.begin('station_diagnostic', self.operation, self.profile)
            self._lift = self._axis = None
            self.observations = []
            self.faulted = self.vacuum_unverified = self.evidence_unverified = False
            self._start_acknowledgement_log()
            self.active = True
            if lift is not None and lift.is_connected:
                self.bind_lift(lift)
            if enabled:
                self.vacuum.set_enabled(True)
                if self.vacuum.is_enabled is not True:
                    raise HardwareSafetyError('Vacuum enable state was not acknowledged.')
                self.observations.append({'vacuum_command': 'enable', 'acknowledged': True})
                self._publish()
            else:
                return self.close()
        except BaseException as exc:
            if self.active:
                if enabled:
                    self.close(error=str(exc))
            else:
                self._release()
            raise

    def _release(self):
        lease, self._lease = self._lease, None
        self.active = False
        if lease is not None:
            lease.__exit__(None, None, None)

    def bind_lift(self, lift, port=None):
        if not self.active:
            raise HardwareSafetyError('No held station is available for this lift.')
        binding = lift.safety_profile(port)
        if self.faulted and self.profile['resources'].get('lift') != binding:
            raise HardwareSafetyError('Only the original participating lift may reconnect during recovery.')
        self.profile = self.store.join_diagnostic_resource(self.token, 'lift', binding)
        self._lift, self._axis = lift, copy.deepcopy(lift.profile.updown_axis)

    def verify_lift(self):
        if self._lift is None:
            return []
        observations, error = verify_stopped_in_place(self._lift.motor, self._axis)
        self.observations.append({'lift_stop': observations, 'error': error})
        if error:
            self.faulted = True
            self._publish(cleanup_error=error)
            raise HardwareSafetyError('Lift stop remains unverified; vacuum is retained: ' + error)
        self.faulted = self.vacuum_unverified or self.evidence_unverified
        return observations

    @contextmanager
    def motion_operation(self, lift, operation):
        if self.faulted:
            raise HardwareSafetyError('Recover the held station before further lift motion.')
        self.bind_lift(lift)
        self.observations.append({'lift_operation': copy.deepcopy(operation)})
        error = ''
        try:
            yield self.observations
        except BaseException as exc:
            error = str(exc)
            raise
        finally:
            self.verify_lift()
            try:
                self._publish(error=error)
            except Exception:
                self.evidence_unverified = True
                self.faulted = True
                raise

    def close(self, *, error='', recovery=False):
        if not self.active:
            return None
        # Never release a held specimen while motor stopping remains uncertain.
        self.verify_lift()
        cleanup_error = ''
        try:
            self.vacuum.set_enabled(False)
            if self.vacuum.is_enabled is not False:
                raise HardwareSafetyError('Vacuum off state was not acknowledged.')
        except Exception as exc:
            cleanup_error = str(exc)
            self.vacuum_unverified = True
            self.faulted = True
        else:
            self.vacuum_unverified = False
            self.faulted = False
        try:
            record = self._publish(error=error, cleanup_error=cleanup_error,
                                   held=bool(cleanup_error), safe=not cleanup_error, recovery=recovery)
        finally:
            # Publication failure keeps a durable pending latch. When outputs
            # actually acknowledged off, no live hold must trap the application.
            if not cleanup_error:
                self._release()
        if cleanup_error:
            raise HardwareSafetyError('Vacuum release remains unverified: ' + cleanup_error)
        return record

    def recover(self, lift=None):
        """Reattach original bindings and release; never reset or enable outputs."""
        if self.active:
            return self.close(recovery=True)
        self._lease = self.store.operation_lease()
        self._lease.__enter__()
        try:
            pending = self.store.pending()
            if pending is None or pending['family'] != 'station_diagnostic':
                raise HardwareSafetyError('No held vacuum station is available for this recovery.')
            if pending['profile'].get('helper') != self.helper:
                raise HardwareSafetyError('Recover the held vacuum in its original panel or helper.')
            resources = pending['profile'].get('resources', {})
            if resources.get('vacuum') != vacuum_binding(self.vacuum):
                raise HardwareSafetyError('Restore the original vacuum port/baud before recovery.')
            self._lift = self._axis = None
            if 'lift' in resources:
                if lift is None or not lift.is_connected or resources['lift'] != lift.safety_profile():
                    raise HardwareSafetyError('Connect the original participating lift before releasing vacuum.')
                self._lift, self._axis = lift, copy.deepcopy(lift.profile.updown_axis)
            self.token, self.profile, self.operation = pending['token'], pending['profile'], pending['plan']
            self.observations = [{'recovery_of_token': pending['token'], 'previous_record_id': (pending.get('record') or {}).get('treatment_id', '')}]
            self._start_acknowledgement_log()
            self.active = True
            return self.close(recovery=True)
        except BaseException:
            if not self.active:
                self._release()
            raise
