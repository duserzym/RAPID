from collections import deque
import unittest
from unittest.mock import Mock

from rapid_main.pulse_circuit import PulseCircuit, PulseCircuitError, validate_pulse_bindings
from rapid_main.pulse_irm import plan_pulse_irm
from tests.test_pulse_irm import pulse_config
from tests.acquisition_fakes import FakeClock


def circuit_config(**overrides):
    values = dict(board=0,dac_channel=1,capacitor_adc_channel=0,fire_bit=1,trim_bit=3,
                  relay_board=1,irm_relay_bit=0,axial_relay_bit=1,transverse_relay_bit=2,
                  charge_timeout_s=1,discharge_timeout_s=1,coil_position=-20000)
    values.update(overrides)
    return pulse_config(**values)


class PulseCircuitTests(unittest.TestCase):
    def circuit(self,values=None,**overrides):
        self.cfg = circuit_config(**overrides)
        self.journal = []
        self.readings = deque(values if values is not None else [0,0,0,.5,.5,.5,.5,0,0,0])
        self.daq,self.relays = Mock(simulated=False),Mock(simulated=False)
        self.daq.analog_input.side_effect = lambda *args:self.readings.popleft() if self.readings else 0
        self.daq.analog_output.side_effect = lambda ch,v,r:self.journal.append(("dac",v))
        self.daq.digital_output.side_effect = lambda port,bit,high:self.journal.append(("ttl",bit,high))
        self.word = 0
        def set_word(word):
            self.word = word
            self.journal.append(("relays",word))
        self.relays.set_digout.side_effect = set_word
        self.relays.get_digout.side_effect = lambda:self.word
        self.clock = FakeClock()
        return PulseCircuit(self.cfg,self.daq,self.relays,sleep=self.clock.sleep,
                            monotonic=self.clock.monotonic,clock=self.clock.now)

    def run_plan(self,circuit,field=50):
        return circuit.execute(plan_pulse_irm(field,"axial",self.cfg),sample_id="S1",run_id="R1")

    def test_fire_requires_stable_voltage_and_cleanup_verifies_discharge(self):
        record = self.run_plan(self.circuit())
        self.assertTrue(record.fired)
        self.assertTrue(record.safe_state_confirmed)
        self.assertEqual(record.schema,"rapidpy.irm.pulse.v1")
        self.assertIn(("relays",5),self.journal)
        self.assertEqual(self.journal[-1],("relays",0))
        phases = [p.name for p in record.phases]
        self.assertEqual(phases.count("charge_readback"),3)
        self.assertLess(phases.index("pre_fire_readback"),phases.index("fire_on"))
        self.assertGreater(phases.index("relay_clear"),phases.index("final_discharge"))
        self.assertEqual(record.to_dict()["configuration"]["feedback_v_per_capacitor_v"],.01)

    def test_backfield_changes_polarity_relay_instead_of_negative_dac_voltage(self):
        record = self.run_plan(self.circuit(backfield_enabled=True),-50)
        self.assertTrue(record.plan.backfield)
        self.assertIn(("relays",4),self.journal)
        self.assertTrue(all(entry[1]>=0 for entry in self.journal if entry[0]=="dac"))

    def test_charge_timeout_never_fires_and_keeps_failure_evidence(self):
        circuit = self.circuit([0]*30)
        with self.assertRaises(PulseCircuitError) as outcome:
            self.run_plan(circuit)
        record = outcome.exception.record
        self.assertIn("charge did not stabilize",record.error)
        self.assertFalse(record.fired)
        self.assertTrue(record.safe_state_confirmed)
        self.assertNotIn(("ttl",1,False),self.journal)

    def test_unverified_discharge_retains_relay_selection_and_reports_unsafe_state(self):
        circuit = self.circuit()
        def read(*args):
            if self.word==0:
                return 0
            return .5
        self.daq.analog_input.side_effect = read
        with self.assertRaises(PulseCircuitError) as outcome:
            self.run_plan(circuit)
        self.assertTrue(outcome.exception.record.fired)
        self.assertFalse(outcome.exception.record.discharged)
        self.assertFalse(outcome.exception.record.safe_state_confirmed)
        self.assertEqual(self.word,5)

    def test_relay_mismatch_blocks_charge_and_firing(self):
        circuit = self.circuit()
        self.relays.get_digout.side_effect = lambda:0
        with self.assertRaises(PulseCircuitError) as outcome:
            self.run_plan(circuit)
        self.assertIn("relay readback mismatch",outcome.exception.record.error)
        self.assertNotIn(("dac",1),self.journal)
        self.assertNotIn(("ttl",1,False),self.journal)

    def test_cancel_during_charge_disables_fire_and_discharges(self):
        circuit = self.circuit()
        circuit.should_cancel = lambda:("dac",1) in self.journal
        with self.assertRaises(PulseCircuitError) as outcome:
            self.run_plan(circuit)
        self.assertIn("cancelled",outcome.exception.record.error)
        self.assertFalse(outcome.exception.record.fired)
        self.assertTrue(outcome.exception.record.discharged)

    def test_invalid_readback_or_pre_fire_drift_never_fires(self):
        for charge in ([float("nan")],[.5,.5,.5,.7]):
            with self.subTest(charge=charge):
                circuit = self.circuit([0,0,0,*charge,0,0,0])
                with self.assertRaises(PulseCircuitError) as outcome:
                    self.run_plan(circuit)
                self.assertFalse(outcome.exception.record.fired)
                self.assertTrue(outcome.exception.record.discharged)

    def test_each_cleanup_output_is_attempted_after_a_zero_voltage_failure(self):
        circuit = self.circuit()
        calls = 0
        def voltage(ch,v,r):
            nonlocal calls
            calls += 1
            if calls>1:
                raise RuntimeError("DAC lost")
        self.daq.analog_output.side_effect = voltage
        with self.assertRaises(PulseCircuitError) as outcome:
            self.run_plan(circuit)
        self.assertTrue(any("DAC lost" in error for error in outcome.exception.record.cleanup_errors))
        self.assertIn(("ttl",1,True),self.journal)
        self.assertIn(("ttl",3,False),self.journal)
        self.assertFalse(outcome.exception.record.safe_state_confirmed)

    def test_invalid_bindings_and_changed_calibration_precede_all_output(self):
        for overrides in ({"fire_bit":-1},{"fire_bit":3},{"irm_relay_bit":9},{"poll_s":0}):
            with self.subTest(overrides=overrides),self.assertRaises(ValueError):
                validate_pulse_bindings(circuit_config(**overrides))
        circuit = self.circuit()
        plan = plan_pulse_irm(50,"axial",self.cfg)
        self.cfg.control_v_per_capacitor_v = .03
        with self.assertRaises(PulseCircuitError) as outcome:
            circuit.execute(plan,sample_id="S1",run_id="R1")
        self.assertIn("changed after planning",outcome.exception.record.error)
        self.assertEqual(self.journal,[])

    def test_discharge_recovery_never_asserts_fire_or_charges(self):
        circuit = self.circuit([.5,.2,0,0,0])
        self.word = 5
        record = circuit.recover_safe_state(sample_id="S1",run_id="R1")
        self.assertTrue(record.safe_state_confirmed)
        self.assertFalse(record.fired)
        self.assertNotIn(("ttl",1,False),self.journal)
        self.assertTrue(all(entry[1]==0 for entry in self.journal if entry[0]=="dac"))
        self.assertEqual(self.word,0)

    def test_hot_or_zeroed_temperature_sensor_blocks_charging_and_fire(self):
        for raw in (5,0,float("nan")):
            with self.subTest(raw=raw):
                circuit = self.circuit(temperature_channels=[2],temperature_slope=10,temperature_offset=0,temperature_hot_c=40)
                base = self.daq.analog_input.side_effect
                self.daq.analog_input.side_effect = lambda channel,vrange:raw if channel==2 else base(channel,vrange)
                with self.assertRaises(PulseCircuitError) as outcome:
                    self.run_plan(circuit)
                self.assertFalse(outcome.exception.record.fired)
                self.assertNotIn(("dac",1),self.journal)

    def test_positioning_service_withholds_motion_after_unverified_discharge(self):
        from types import SimpleNamespace
        from rapid_main.pulse_treatment import PulseTreatmentService,PulseTreatmentError
        circuit = self.circuit()
        self.daq.analog_input.side_effect = lambda *args:0 if self.word==0 else .5
        vertical,turning = Mock(),Mock()
        outcome = SimpleNamespace(ok=True,actual=0)
        vertical.home_to_top.return_value = vertical.move_to.return_value = outcome
        turning.rotate_to.return_value = outcome
        service = PulseTreatmentService(circuit,vertical,turning,clock=self.clock.now)
        with self.assertRaises(PulseTreatmentError) as result:
            service.execute(plan_pulse_irm(50,"axial",self.cfg),sample_id="S1",run_id="R1",sample_height=3000)
        self.assertFalse(result.exception.record.safe_state_confirmed)
        self.assertIn("Motor return withheld",result.exception.record.cleanup_errors[0])
        vertical.home_to_top.assert_called_once()

    def test_worker_indexes_current_run_pulse_evidence_and_digest(self):
        import hashlib
        import json
        from pathlib import Path
        import tempfile
        from rapid_main.measurement_worker import MeasurementWorker
        from tests.test_susceptibility_queue import _WorkerBackend,_meta
        record = self.run_plan(self.circuit())
        class Backend(_WorkerBackend):
            def __init__(self):
                super().__init__()
                self.pulse_treatment_records = []
            def set_demag_step(self,label):
                super().set_demag_step(label)
                self.pulse_treatment_records.append(record)
        with tempfile.TemporaryDirectory() as folder:
            output = Path(folder)/"S1"
            worker = MeasurementWorker(meta=_meta("S1"),labels=["IRM50"],output_dir=output,backend=Backend())
            errors = []
            worker.error_occurred.connect(errors.append)
            worker.run()
            self.assertEqual(errors,[])
            artifact = output/"pulse_treatments"/f"{record.treatment_id}.json"
            self.assertEqual(json.loads(artifact.read_text())["schema"],"rapidpy.irm.pulse.v1")
            index = json.loads((output/"artifact_index.json").read_text())
            row = next(row for row in index["artifacts"] if row["name"]==f"pulse_treatment:{record.treatment_id}")
            self.assertEqual(row["sha256"],hashlib.sha256(artifact.read_bytes()).hexdigest())
            self.assertTrue(row["required"])
            summary = json.loads((output/"workflow_summary.json").read_text())
            self.assertEqual(summary["pulse_treatment_ids"],[record.treatment_id])
