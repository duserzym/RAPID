"""Original-station field cutoff evidence before queue transfer and release."""
import copy
from dataclasses import asdict, dataclass, field
from datetime import datetime, timezone
import json
import uuid

from rapidpy_common.adwin_af import AdwinAFController
from rapidpy_common.mcc_daq import MccDaq
from rapidpy_common.hardware_safety import HardwareSafetyError
from .arm_bias import ArmBiasBackend
from .pulse_circuit import PulseCircuit, PulseCircuitError, validate_pulse_bindings


def _snapshot(value):
    return json.loads(json.dumps(value, allow_nan=False))


@dataclass(frozen=True)
class QueueFieldOutputsRecord:
    operation: dict
    station_profile: dict
    arm: dict | None
    pulse: dict | None
    af: dict | None
    errors: tuple
    safe_state_confirmed: bool
    field_context: dict
    simulated: bool = False
    record_id: str = field(default_factory=lambda: uuid.uuid4().hex)
    timestamp_iso: str = field(default_factory=lambda: datetime.now(timezone.utc).isoformat())
    schema: str = 'rapidpy.queue_field_outputs.v1'

    def to_dict(self):
        return asdict(self)


class QueueFieldOffProof(dict):
    def __init__(self, owner, context):
        super().__init__(_snapshot(context))
        self.owner = owner


def field_outputs_proof(session, value):
    if value is True:
        return None  # Existing explicit operator-attestation callers.
    if not isinstance(value, QueueFieldOffProof) or value.owner.require(session) != value:
        raise HardwareSafetyError('Verify original participating field outputs off before transfer or release.')
    return _snapshot(dict(value))


class QueueFieldOutputs:
    """Off-only: no boot, charge, fire, specimen motion or vacuum release."""
    def __init__(self, af, pulse, arm):
        if (not isinstance(af, AdwinAFController) or not isinstance(pulse, PulseCircuit)
                or not isinstance(pulse.daq, MccDaq) or not isinstance(pulse.relays, AdwinAFController)
                or not isinstance(arm, ArmBiasBackend) or not isinstance(arm.controller, MccDaq)):
            raise HardwareSafetyError('Queue field cutoff requires native AF, pulse and ARM circuits.')
        self.af, self.pulse, self.arm = af, pulse, arm
        self.instance_id = uuid.uuid4().hex
        self.profile = self._profile()
        self._validate_circuits()
        self.last_record = None

    def _profile(self):
        return _snapshot(dict(helper='rapid_main_queue_field_outputs',
            af=dict(board=asdict(self.af.board), limits=asdict(self.af.limits)),
            pulse=dict(configuration=asdict(self.pulse.cfg), relay_board=asdict(self.pulse.relays.board)),
            arm=asdict(self.arm.cfg)))

    def _validate_circuits(self):
        validate_pulse_bindings(self.pulse.cfg)
        cfg, arm = self.pulse.cfg, self.arm.cfg
        if (self.pulse.daq.board != cfg.board or self.pulse.relays.board.board_num != cfg.relay_board
                or self.arm.controller.board != arm.arm_board
                or any(getattr(device, 'simulated', False) is True for device in
                       (self.af, self.pulse, self.pulse.daq, self.pulse.relays, self.arm, self.arm.controller))):
            raise HardwareSafetyError('Field cutoff circuit boards differ from the original physical station.')
        if cfg.board == arm.arm_board and (cfg.dac_channel == arm.arm_dac_channel
                or (cfg.digital_port == arm.arm_digital_port and arm.arm_gate_bit in {cfg.fire_bit, cfg.trim_bit})):
            raise HardwareSafetyError('ARM and pulse output bindings must not alias on the same MCC board.')

    def _validate(self, session):
        session.child_store._owned()
        root = session.store._queue(session.token)
        if (root['profile'].get('helper') != 'rapid_main_queue'
                or root['profile']['stage_profiles'].get('field_outputs') != self.profile
                or self._profile() != self.profile):
            raise HardwareSafetyError('Restore the original queue field circuit calibration and binding before I/O.')
        self._validate_circuits()
        return root

    def require(self, session):
        self._validate(session)
        context = session.store.latest_field_context(session.token)
        if (not isinstance(context, dict) or context.get('schema') != 'rapidpy.queue_field_off_state.v1'
                or context.get('queue_token') != session.token or context.get('instance_id') != self.instance_id
                or context.get('phase') != 'off'):
            raise HardwareSafetyError('Verify participating outputs off on the original live queue circuits.')
        return QueueFieldOffProof(self, context)

    def verify_off(self, session, *, recovery=False):
        root = self._validate(session)
        child, operation = session.child_store, dict(action='verify_participating_field_outputs_off', recovery=recovery)
        stage = root['stage']
        if stage and stage['status'] == 'pending':
            if (recovery is not True or session.is_recovery is not True or stage['family'] != 'field_outputs'
                    or stage['profile'] != self.profile):
                raise HardwareSafetyError('Recover the original unfinished stage before field-output verification.')
            token = stage['token']
        else:
            token = child.begin('field_outputs', operation, self.profile, run_id=root['run_id'])
        errors, arm_record, pulse_record, af_record = [], None, None, None
        # ARM cutoff is independent and attempted even if capacitor discharge fails.
        try:
            arm_record = self.arm.clear_bias_verified().to_dict()
            if arm_record['safe_state_confirmed'] is not True:
                errors.append('ARM cutoff: ' + arm_record['error'])
        except Exception as exc:
            errors.append('ARM cutoff: ' + str(exc))
        # Stop AF processes and zero both DACs before pulse recovery can switch
        # shared relays. Capacitor charge is still unknown at this boundary.
        af_zero = False
        try:
            self.af.recover_safe_field(clear_relays=False)
            af_record = copy.deepcopy(self.af.last_safe_field_recovery)
            af_zero = af_record['outputs_zero_confirmed'] is True
        except Exception as exc:
            af_record = copy.deepcopy(getattr(self.af, 'last_safe_field_recovery', None))
            errors.append('AF cutoff: ' + str(exc))
        initial_af_record = copy.deepcopy(af_record)
        discharged = False
        try:
            pulse_record = self.pulse.recover_safe_state(sample_id='queue-field-cutoff',
                run_id=root['run_id'], clear_relays=af_zero).to_dict()
            discharged = pulse_record['safe_state_confirmed'] is True and pulse_record['simulated'] is False
        except PulseCircuitError as exc:
            pulse_record = exc.record.to_dict()
            errors.append('Pulse discharge: ' + str(exc))
        except Exception as exc:
            errors.append('Pulse discharge: ' + str(exc))
        if af_zero and discharged:
            try:
                self.af.recover_safe_field(clear_relays=True)
                af_record = copy.deepcopy(self.af.last_safe_field_recovery)
            except Exception as exc:
                af_record = copy.deepcopy(getattr(self.af, 'last_safe_field_recovery', None))
                errors.append('AF cutoff: ' + str(exc))
        else:
            errors.append('AF relay cutoff withheld until AF outputs and pulse discharge are verified.')
        if af_record is not None:
            af_record['initial_output_cutoff'] = initial_af_record
        try:
            self._validate(session)
        except Exception as exc:
            errors.append(str(exc))
        safe = not errors and discharged and arm_record is not None and af_record is not None
        context = dict(schema='rapidpy.queue_field_off_state.v1', queue_token=session.token,
            instance_id=self.instance_id, evidence_id=uuid.uuid4().hex, phase='off' if safe else 'unverified')
        record = QueueFieldOutputsRecord(operation, copy.deepcopy(self.profile), arm_record, pulse_record,
                                          af_record, tuple(errors), safe, context)
        self.last_record = record
        child.finish(token, self.profile, record)
        if not safe:
            raise HardwareSafetyError('Queue field cutoff remains pending: ' + '; '.join(errors))
        return self.require(session)
