"""Bounded capacitor pulse execution with verified discharge and phase evidence."""
from dataclasses import asdict, dataclass
from datetime import datetime, timezone
import json
import math
import time
import uuid

from .pulse_irm import plan_pulse_irm


def validate_pulse_bindings(cfg):
    integers = (cfg.board,cfg.dac_channel,cfg.capacitor_adc_channel,cfg.voltage_range,
                cfg.digital_port,cfg.fire_bit,cfg.trim_bit,cfg.relay_board)
    if any(not isinstance(value,int) or isinstance(value,bool) or value < 0 for value in integers):
        raise ValueError("Pulse IRM MCC/ADwin boards and channels must be explicitly configured.")
    bits = (cfg.irm_relay_bit,cfg.axial_relay_bit,cfg.transverse_relay_bit)
    if any(not isinstance(bit,int) or isinstance(bit,bool) or not 0 <= bit <= 5 for bit in bits) or len(set(bits)) != 3:
        raise ValueError("Pulse coil/polarity relay bits must be distinct ADwin-light-16 outputs in 0..5.")
    if cfg.fire_bit == cfg.trim_bit:
        raise ValueError("Pulse fire and trim must use distinct MCC output bits.")
    if any(not math.isfinite(float(value)) or value <= 0 for value in
           (cfg.charge_timeout_s,cfg.discharge_timeout_s,cfg.discharged_max_v,cfg.poll_s)):
        raise ValueError("Pulse timeout, discharge threshold and poll interval must be positive and finite.")
    if cfg.discharged_max_v > 10 or cfg.poll_s > 1:
        raise ValueError("Pulse discharge threshold must be at most 10 V and polling at most one second.")
    if cfg.temperature_channels:
        if any(not isinstance(channel,int) or channel<0 for channel in cfg.temperature_channels):
            raise ValueError("Enabled pulse temperature channels must be configured.")
        if not all(math.isfinite(float(value)) for value in (cfg.temperature_slope,cfg.temperature_offset,cfg.temperature_hot_c)) or cfg.temperature_slope<=0:
            raise ValueError("Enabled pulse temperature sensors require finite calibration and a positive slope.")


@dataclass(frozen=True)
class PulsePhase:
    name: str
    timestamp_iso: str
    ok: bool
    detail: str


@dataclass(frozen=True)
class PulseCircuitRecord:
    treatment_id: str
    sample_id: str
    run_id: str
    plan: object
    phases: tuple[PulsePhase,...]
    capacitor_readings_v: tuple[float,...]
    fired: bool
    discharged: bool
    safe_state_confirmed: bool
    error: str
    cleanup_errors: tuple[str,...]
    simulated: bool = False
    schema: str = "rapidpy.irm.pulse.v1"
    configuration_json: str = ""
    residual_pulses_completed: int = 0

    def to_dict(self):
        payload = asdict(self)
        payload["charged_pulse_fired"] = self.fired
        payload["configuration"] = json.loads(payload.pop("configuration_json")) if self.configuration_json else None
        return payload


class PulseCircuitError(RuntimeError):
    def __init__(self,record):
        self.record = record
        super().__init__("; ".join([record.error,*record.cleanup_errors]).strip("; "))


class PulseCircuit:
    simulated = False

    def __init__(self,cfg,daq,relays,*,should_cancel=None,sleep=time.sleep,monotonic=time.monotonic,clock=None):
        self.cfg,self.daq,self.relays = cfg,daq,relays
        self.should_cancel = should_cancel or (lambda:False)
        self.sleep,self.monotonic = sleep,monotonic
        self.clock = clock or (lambda:datetime.now(timezone.utc))

    def recover_safe_state(self,*,sample_id,run_id,clear_relays=True):
        """Zero/bleed only: never charge or assert the active-low fire switch."""
        if type(clear_relays) is not bool:
            raise ValueError('Relay recovery permission must be explicit.')
        phases,readings,errors = [],[],[]
        discharged = False
        def action(name,operation):
            try:
                value = operation()
                phases.append(PulsePhase(name,self.clock().isoformat(),True,str(value)))
                return value
            except Exception as exc:
                phases.append(PulsePhase(name,self.clock().isoformat(),False,str(exc)))
                errors.append(f"{name}: {exc}")
                return None
        try:
            validate_pulse_bindings(self.cfg)
            if not math.isfinite(self.cfg.feedback_v_per_capacitor_v) or self.cfg.feedback_v_per_capacitor_v <= 0:
                raise ValueError("Capacitor feedback calibration is required for discharge recovery.")
        except Exception as exc:
            errors.append(str(exc))
        else:
            action("recovery_charge_zero",lambda:self.daq.analog_output(self.cfg.dac_channel,0,self.cfg.voltage_range))
            action("recovery_fire_inhibit",lambda:self.daq.digital_output(self.cfg.digital_port,self.cfg.fire_bit,True))
            action("recovery_trim_on",lambda:self.daq.digital_output(self.cfg.digital_port,self.cfg.trim_bit,self.cfg.trim_on_high))
            deadline,stable = self.monotonic()+self.cfg.discharge_timeout_s,0
            try:
                while stable<3:
                    raw = float(self.daq.analog_input(self.cfg.capacitor_adc_channel,self.cfg.voltage_range))
                    if not math.isfinite(raw) or not 0<=raw<=10:
                        raise RuntimeError("Invalid capacitor readback during discharge recovery.")
                    cap = raw/self.cfg.feedback_v_per_capacitor_v
                    readings.append(cap)
                    phases.append(PulsePhase("recovery_readback",self.clock().isoformat(),True,str(cap)))
                    stable = stable+1 if cap<=self.cfg.discharged_max_v else 0
                    if self.monotonic()>deadline:
                        raise TimeoutError("Discharge recovery exceeded its deadline.")
                    if stable<3:
                        self.sleep(self.cfg.poll_s)
                discharged = True
            except Exception as exc:
                errors.append(str(exc))
            if not clear_relays:
                errors.append('Recovery relay clear withheld until participating AF outputs are verified stopped and zero.')
            if discharged and not errors:
                action("recovery_relay_clear",lambda:self.relays.set_digout(0))
                actual = action("recovery_relay_readback",self.relays.get_digout)
                if actual!=0:
                    errors.append("Recovery relay clear was not verified.")
        record = PulseCircuitRecord("irm-"+uuid.uuid4().hex,sample_id,run_id,None,tuple(phases),tuple(readings),
                                    False,discharged,discharged and not errors,"",tuple(errors))
        if discharged:
            snapshot = {name:getattr(self.cfg,name) for name in ("calibration_source","board","dac_channel","capacitor_adc_channel",
                        "feedback_v_per_capacitor_v","voltage_range","digital_port","fire_bit","trim_bit","trim_on_high",
                        "relay_board","discharged_max_v","discharge_timeout_s")}
            record = PulseCircuitRecord(record.treatment_id,sample_id,run_id,None,tuple(phases),tuple(readings),False,
                         discharged,discharged and not errors,"",tuple(errors),configuration_json=json.dumps(snapshot,allow_nan=False))
        if errors:
            raise PulseCircuitError(record)
        return record

    def execute(self,plan,*,sample_id,run_id,position_for_pulse=None):
        phases,readings,cleanup = [],[],[]
        fired,discharged,started,error = False,False,False,""
        residual_pulses_completed = 0
        def phase(name,operation):
            try:
                value = operation()
                phases.append(PulsePhase(name,self.clock().isoformat(),True,str(value)))
                return value
            except Exception as exc:
                phases.append(PulsePhase(name,self.clock().isoformat(),False,str(exc)))
                raise
        def cancel():
            if self.should_cancel():
                raise InterruptedError("Pulse IRM cancelled.")
        def voltage(value):
            return self.daq.analog_output(self.cfg.dac_channel,value,self.cfg.voltage_range)
        def fire(high):
            return self.daq.digital_output(self.cfg.digital_port,self.cfg.fire_bit,high)
        def trim(on):
            return self.daq.digital_output(self.cfg.digital_port,self.cfg.trim_bit,self.cfg.trim_on_high if on else not self.cfg.trim_on_high)
        def read():
            value = float(self.daq.analog_input(self.cfg.capacitor_adc_channel,self.cfg.voltage_range))
            if not math.isfinite(value) or value < 0 or value > 10:
                raise RuntimeError("Invalid capacitor ADC voltage; cannot fire or confirm discharge.")
            cap = value/self.cfg.feedback_v_per_capacitor_v
            readings.append(cap)
            return cap
        def wait(seconds,*,cancellable=True):
            end = self.monotonic()+seconds
            while self.monotonic()<end:
                if cancellable:
                    cancel()
                self.sleep(min(self.cfg.poll_s,max(0,end-self.monotonic())))
        def temperatures():
            if not self.cfg.temperature_channels:
                return "No temperature sensors enabled by station configuration."
            for channel in self.cfg.temperature_channels:
                raw = float(self.daq.analog_input(channel,self.cfg.voltage_range))
                temperature = raw*self.cfg.temperature_slope-self.cfg.temperature_offset
                if not math.isfinite(raw) or not math.isfinite(temperature) or abs(self.cfg.temperature_offset+temperature)<20:
                    raise RuntimeError("Pulse coil temperature sensor is invalid or zeroed.")
                if temperature>=self.cfg.temperature_hot_c:
                    raise RuntimeError(f"Pulse coil temperature {temperature:g} C exceeds the configured hot threshold.")
            return "Enabled temperature sensors are below the hot threshold."

        def discharge():
            deadline = self.monotonic()+self.cfg.discharge_timeout_s
            stable = 0
            while stable < 3:
                cap = phase("discharge_readback",read)
                stable = stable+1 if cap <= self.cfg.discharged_max_v else 0
                if self.monotonic()>deadline:
                    raise TimeoutError("Capacitor discharge could not be confirmed before its deadline.")
                if stable < 3:
                    self.sleep(self.cfg.poll_s)
        configuration_json = ""
        try:
            configuration_json = json.dumps(asdict(self.cfg),allow_nan=False,sort_keys=True)
            validate_pulse_bindings(self.cfg)
            if plan_pulse_irm(plan.field_mT,plan.coil,self.cfg) != plan:
                raise ValueError("Pulse calibration changed after planning.")
            if plan.operation == "zero_field" and position_for_pulse is None:
                raise ValueError("Zero-field IRM requires the verified load-to-coil specimen lifecycle.")
            if not sample_id.strip() or not run_id.strip():
                raise ValueError("Pulse requires specimen and run identities.")
            if any(getattr(device,"simulated",False) is True for device in (self.daq,self.relays)):
                raise ValueError("Simulated circuit cannot execute a live pulse.")
            cancel()
            started = True
            phase("temperature_preflight",temperatures)
            phase("fire_inhibit",lambda:fire(True))
            phase("charge_zero",lambda:voltage(0))
            phase("initial_trim",lambda:trim(True))
            phase("initial_discharge",discharge)
            phase("relay_boot",self.relays.boot_board)
            # IRM axial selects the transverse relay HIGH, and vice versa.
            word = (0 if plan.backfield else 1<<self.cfg.irm_relay_bit) | (1<<(self.cfg.transverse_relay_bit if plan.coil=="axial" else self.cfg.axial_relay_bit))
            phase("relay_select",lambda:self.relays.set_digout(word))
            actual = phase("relay_readback",self.relays.get_digout)
            if actual != word:
                raise RuntimeError(f"Pulse relay readback mismatch: requested {word}, read {actual}.")
            wait(1)
            cancel()
            if position_for_pulse is not None:
                # RockmagStep discharges twice at the load position before
                # moving the specimen into the pulse coil.
                for index in range(2):
                    cancel()
                    phase(f"residual_temperature_{index}", temperatures)
                    phase(f"residual_discharge_{index}",discharge)
                    cancel()
                    phase(f"residual_fire_{index}",lambda:fire(False))
                    wait(plan.fire_hold_s)
                    phase(f"residual_fire_off_{index}",lambda:fire(True))
                    residual_pulses_completed += 1
                    wait(1)
                phase("position_specimen",position_for_pulse)
                cancel()
            if plan.operation == "zero_field":
                phase("zero_field_no_charge", lambda: "Two residual pulses at load; no charged pulse at specimen.")
            else:
                phase("temperature_pre_charge", temperatures)
                phase("charge_target",lambda:voltage(plan.charge_control_v))
                phase("trim_off",lambda:trim(False))
                deadline = self.monotonic()+self.cfg.charge_timeout_s
                stable = 0
                while stable < 3:
                    cancel()
                    phase("temperature_charge",temperatures)
                    cap = phase("charge_readback",read)
                    if cap > getattr(self.cfg,f"{plan.coil}_capacitor_max_v"):
                        raise RuntimeError("Capacitor exceeded its calibrated voltage limit.")
                    if abs(cap-plan.capacitor_v) <= plan.capacitor_tolerance_v:
                        phase("charge_hold",lambda:voltage(0))
                        phase("trim_hold",lambda:trim(False))
                        stable += 1
                    elif cap > plan.capacitor_v:
                        stable = 0
                        phase("charge_stop",lambda:voltage(0))
                        phase("trim_target",lambda:trim(True))
                    else:
                        stable = 0
                        phase("trim_off",lambda:trim(False))
                        phase("charge_resume",lambda:voltage(plan.charge_control_v))
                    if self.monotonic()>deadline:
                        raise TimeoutError("Capacitor charge did not stabilize before its deadline.")
                    if stable < 3:
                        wait(self.cfg.poll_s)
                cancel()
                # One final unaveraged read avoids firing on stale charge evidence.
                phase("temperature_pre_fire",temperatures)
                cap = phase("pre_fire_readback",read)
                if abs(cap-plan.capacitor_v)>plan.capacitor_tolerance_v:
                    raise RuntimeError("Capacitor drifted outside tolerance before firing.")
                phase("fire_on",lambda:fire(False))
                fired = True
                wait(plan.fire_hold_s)
        except Exception as exc:
            error = f"{type(exc).__name__}: {exc}"
        finally:
            if started:
                for name,operation in (("charge_zero",lambda:voltage(0)),("fire_off",lambda:fire(True)),("trim_on",lambda:trim(True))):
                    try:
                        phase("cleanup_"+name,operation)
                    except Exception as exc:
                        cleanup.append(f"{name}: {exc}")
                try:
                    phase("final_discharge",discharge)
                    discharged = True
                except Exception as exc:
                    cleanup.append(f"discharge: {exc}")
                if discharged and not cleanup:
                    try:
                        phase("relay_clear",lambda:self.relays.set_digout(0))
                        actual = phase("relay_clear_readback",self.relays.get_digout)
                        if actual != 0:
                            raise RuntimeError("Pulse relays did not return to zero.")
                    except Exception as exc:
                        cleanup.append(f"relay_clear: {exc}")
        record = PulseCircuitRecord("irm-"+uuid.uuid4().hex,sample_id,run_id,plan,tuple(phases),tuple(readings),
                                    fired,discharged,started and discharged and not cleanup,error,tuple(cleanup),configuration_json=configuration_json,
                                    schema="rapidpy.irm.zero_field.v1" if plan.operation == "zero_field" else "rapidpy.irm.pulse.v1",
                                    residual_pulses_completed=residual_pulses_completed)
        if error or cleanup:
            raise PulseCircuitError(record)
        return record
