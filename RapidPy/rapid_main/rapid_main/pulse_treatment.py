"""Queue pulse specimen positioning and safe-return evidence."""
from dataclasses import asdict,dataclass
from datetime import datetime,timezone
import math
import uuid

from .pulse_irm import plan_pulse_irm
from .pulse_circuit import PulseCircuit, PulseCircuitError, PulsePhase, validate_pulse_bindings


def pulse_motion_target(cfg,sample_height):
    if not isinstance(cfg.coil_position,int) or cfg.coil_position==0 or int(sample_height)<=0:
        raise ValueError("Pulse IRM coil position and specimen height must be configured.")
    target = math.floor(cfg.coil_position+int(sample_height)/2)
    if target==0 or (target>0)!=(cfg.coil_position>0):
        raise ValueError("Specimen centre would cross the configured pulse coil side.")
    return target


@dataclass(frozen=True)
class PulseTreatmentRecord:
    treatment_id: str
    sample_id: str
    run_id: str
    target_position: int
    orientation_deg: float
    circuit: object
    phases: tuple[PulsePhase,...]
    safe_state_confirmed: bool
    error: str
    cleanup_errors: tuple[str,...]
    simulated: bool = False
    schema: str = "rapidpy.irm.treatment.v1"

    def to_dict(self):
        payload = asdict(self)
        if self.circuit is not None:
            payload["circuit"] = self.circuit.to_dict()
        return payload


class PulseTreatmentError(RuntimeError):
    def __init__(self,record):
        self.record = record
        super().__init__("; ".join([record.error,*record.cleanup_errors]).strip("; "))


class PulseTreatmentService:
    def __init__(self,circuit,vertical,turning,*,clock=None):
        self.circuit,self.vertical,self.turning = circuit,vertical,turning
        self.clock = clock or (lambda:datetime.now(timezone.utc))

    def execute(self,plan,*,sample_id,run_id,sample_height,orientation_deg=0):
        phases,cleanup = [],[]
        pulse_record,error,moved,pulse_started = None,"",False,False
        target = pulse_motion_target(self.circuit.cfg,sample_height)
        def motion(name,operation):
            try:
                outcome = operation()
                if not outcome.ok or not math.isfinite(float(outcome.actual)):
                    raise RuntimeError(getattr(outcome,"detail","") or "Pulse specimen motion was not verified.")
                phases.append(PulsePhase(name,self.clock().isoformat(),True,str(outcome.actual)))
            except Exception as exc:
                phases.append(PulsePhase(name,self.clock().isoformat(),False,str(exc)))
                raise
        try:
            if not sample_id.strip() or not run_id.strip():
                raise ValueError("Pulse requires specimen and run identities before motion.")
            validate_pulse_bindings(self.circuit.cfg)
            if plan_pulse_irm(plan.field_mT,plan.coil,self.circuit.cfg)!=plan:
                raise ValueError("Pulse calibration changed after specimen planning.")
            if not math.isfinite(float(orientation_deg)) or orientation_deg not in {0,90}:
                raise ValueError("Pulse orientation must be 0 or 90 degrees.")
            if self.circuit.should_cancel():
                raise InterruptedError("Pulse treatment cancelled before motion.")
            moved = True
            motion("home_before",self.vertical.home_to_top)
            motion("orient_specimen",lambda:self.turning.rotate_to(orientation_deg))
            def position():
                motion("move_to_pulse_coil",lambda:self.vertical.move_to(target,speed_index=1))
            try:
                pulse_started = True
                pulse_record = self.circuit.execute(plan,sample_id=sample_id,run_id=run_id,position_for_pulse=position)
            except PulseCircuitError as exc:
                pulse_record = exc.record
                raise
        except Exception as exc:
            error = f"{type(exc).__name__}: {exc}"
        finally:
            if moved and ((pulse_record is None and not pulse_started) or (pulse_record is not None and pulse_record.safe_state_confirmed)):
                for name,operation in (("orient_reference",lambda:self.turning.rotate_to(360)),("home_after",self.vertical.home_to_top)):
                    try:
                        motion(name,operation)
                    except Exception as exc:
                        cleanup.append(f"{name}: {exc}")
            elif moved:
                cleanup.append("Motor return withheld because capacitor/relay safe state is unverified.")
        safe = pulse_record is not None and pulse_record.safe_state_confirmed and not cleanup
        record = PulseTreatmentRecord("irm-"+uuid.uuid4().hex,sample_id,run_id,target,orientation_deg,
                                     pulse_record,tuple(phases),safe,error,tuple(cleanup),
                                     schema="rapidpy.irm.zero_treatment.v1" if plan.operation == "zero_field" else "rapidpy.irm.treatment.v1")
        if error or cleanup:
            raise PulseTreatmentError(record)
        return record


class PulseIrmBackend:
    simulated = False

    def __init__(self,cfg,relay_controller,*,daq=None):
        from rapidpy_common.mcc_daq import MccDaq
        validate_pulse_bindings(cfg)
        if relay_controller is None:
            raise ValueError("Pulse IRM requires the configured ADwin relay board.")
        if int(relay_controller.board.board_num)!=cfg.relay_board:
            raise ValueError("Pulse relay driver board differs from the imported channel mapping.")
        self.cfg,self.relays = cfg,relay_controller
        self.daq = daq if daq is not None else MccDaq(cfg.board)

    def is_connected(self):
        return self.daq is not None and self.relays is not None

    def make_circuit(self,**options):
        return PulseCircuit(self.cfg,self.daq,self.relays,**options)
