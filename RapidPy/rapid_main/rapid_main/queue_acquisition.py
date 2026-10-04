"""Durable acquisition settlement under the original specimen queue claim."""
from dataclasses import asdict, dataclass, field
from datetime import datetime, timezone
import copy
import math
import uuid

from rapidpy_common.hardware_safety import HardwareSafetyError
from rapidpy_common.motor_diagnostic_safety import verify_stopped_in_place
from .acquisition import BracketedAcquisition, RecoveryRecord
from .magnetometer import validate_block_observations
from .susceptibility_acquisition import SusceptibilityAcquisitionRecord


@dataclass(frozen=True)
class QueueAcquisitionRecord:
    operation: dict
    evidence: dict | None
    cleanup_observations: tuple
    error: str
    cleanup_error: str
    safe_state_confirmed: bool
    simulated: bool = False
    record_id: str = field(default_factory=lambda: uuid.uuid4().hex)
    timestamp_iso: str = field(default_factory=lambda: datetime.now(timezone.utc).isoformat())
    schema: str = 'rapidpy.queue_acquisition.v1'

    def to_dict(self):
        return asdict(self)


def _evidence(kind, backend, result, operation):
    if kind == 'flux_recovery':
        if (backend._bracketed.simulated is not False
                or not isinstance(result, RecoveryRecord) or not result.commands
                or result.attempt < 1 or any(event.ok is not True for event in result.commands)):
            raise HardwareSafetyError('Flux recovery lacks completed native command evidence.')
        return asdict(result)
    if kind == 'squid':
        acquisition = backend._bracketed.last_acquisition
        if not isinstance(acquisition, BracketedAcquisition) or acquisition.block is not result:
            raise HardwareSafetyError('The returned SQUID block lacks its original completed acquisition.')
        block = acquisition.block
        audit = block.audit
        if (audit is None or audit.simulated is not False
                or audit.sample_name != operation['sample_id'] or audit.run_id != operation['run_id']
                or (audit.zero_position, audit.measurement_position) != tuple(operation['positions'])
                or not audit.block_id or not acquisition.commands
                or any(event.ok is not True for event in acquisition.commands)):
            raise HardwareSafetyError('SQUID acquisition identity, geometry, or command evidence is unverified.')
        validate_block_observations(block)
        payload = asdict(acquisition)
        payload['transport_recoveries'] = [asdict(record) for record in backend._bracketed.transport_recovery_records]
        return payload
    record = backend._susceptibility_records[-1]
    if (not isinstance(record, SusceptibilityAcquisitionRecord) or record.simulated is not False
            or record.sample_id != operation['sample_id'] or record.is_holder is not False
            or record.sample_height != operation['specimen_geometry']['sample_height']
            or record.target_position != operation['positions'][0]
            or record.outcome != 'completed' or record.safe_state_confirmed is not True
            or record.error or record.safe_return_error or not record.phases
            or any(phase.ok is not True for phase in record.phases)
            or record.susceptibility is None or not math.isfinite(float(record.susceptibility))
            or result != record.susceptibility):
        raise HardwareSafetyError('Susceptibility acquisition evidence is incomplete or unverified.')
    return record.to_dict()


def run_queue_acquisition(backend, kind, acquire):
    """Begin before instrument I/O; settlement never releases the specimen grip."""
    geometry = backend._queue_specimen_geometry
    geometry.height(backend._sample_name)
    child = geometry.session.child_store
    if backend._safety_store is not child:
        raise HardwareSafetyError('Acquisition must borrow the original claimed queue journal.')
    if backend._halt_check is not None and backend._halt_check():
        raise InterruptedError('Queue acquisition cancelled before I/O.')
    height = geometry.context.sample_height
    positions = (geometry.positions(backend._sample_name) if kind in {'squid', 'flux_recovery'} else
                 (math.floor(backend._config.susceptibility.coil_position + height / 2),))
    if any(type(value) is not int or not -(2**31) <= value < 2**31 for value in positions):
        raise HardwareSafetyError('Acquisition targets exceed the accepted motor count range.')
    operation = dict(action=kind, sample_id=backend._sample_name, run_id=backend._run_id,
                     specimen_geometry=geometry.context.to_dict(), positions=list(positions))
    profile = backend._acquisition_safety_profile()
    axes = copy.deepcopy(backend._axes)
    previous_susc_count = len(getattr(backend, '_susceptibility_records', ()))
    token = child.begin('acquisition', operation, profile,
                        sample_id=backend._sample_name, run_id=backend._run_id)
    previous = backend._geometry_stage_token
    backend._geometry_stage_token = token
    evidence, error, result, cleanup, failures = None, '', None, [], []
    try:
        try:
            result = acquire()
            if kind == 'susceptibility' and len(backend._susceptibility_records) != previous_susc_count + 1:
                raise HardwareSafetyError('Susceptibility did not produce a new acquisition record.')
            evidence = _evidence(kind, backend, result, operation)
            # Configuration, routing, identity and cancellation may change during a read.
            geometry.height(backend._sample_name, own_stage_token=token)
            if backend._halt_check is not None and backend._halt_check():
                raise InterruptedError('Queue acquisition cancelled before settlement.')
        except Exception as exc:
            error = str(exc) or type(exc).__name__
            if kind == 'susceptibility' and len(backend._susceptibility_records) == previous_susc_count + 1:
                failed = backend._susceptibility_records[-1]
                if isinstance(failed, SusceptibilityAcquisitionRecord):
                    evidence = failed.to_dict()
        finally:
            # Every axis is attempted independently, including after failed instrument I/O.
            for key, axis in axes.items():
                samples, failure = verify_stopped_in_place(backend._client, axis)
                cleanup.append(dict(axis=key, samples=samples, error=failure))
                if failure:
                    failures.append(f'{key}: {failure}')
        if backend._halt_check is not None and backend._halt_check():
            error = error or 'Queue acquisition cancelled during cleanup.'
        if not error:
            try:
                geometry.height(backend._sample_name, own_stage_token=token)
            except Exception as exc:
                error = str(exc) or type(exc).__name__
        record = QueueAcquisitionRecord(operation, evidence, tuple(cleanup), error,
            '; '.join(failures), not error and not failures and evidence is not None)
        backend._last_queue_acquisition = record
        # The store publishes immutable linked evidence before updating the stage.
        child.finish(token, profile, record)
        if not record.safe_state_confirmed:
            raise HardwareSafetyError('Queue acquisition remains pending: ' +
                                      '; '.join(filter(None, (error, record.cleanup_error))))
        return result
    finally:
        backend._geometry_stage_token = previous
