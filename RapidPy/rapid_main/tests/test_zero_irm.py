import hashlib
import json
from pathlib import Path
import tempfile
from types import SimpleNamespace
import unittest
from unittest.mock import Mock

from rapid_main.pulse_circuit import PulseCircuit, PulseCircuitError
from rapid_main.pulse_irm import plan_pulse_irm
from rapid_main.pulse_treatment import PulseTreatmentService, PulseTreatmentError
from tests.test_pulse_circuit import circuit_config
from tests.acquisition_fakes import FakeClock


class ZeroIrmTests(unittest.TestCase):
    def fixture(self, coil="axial"):
        self.cfg = circuit_config(axial_calibrated=False, transverse_calibrated=False)
        self.log = []
        self.clock = FakeClock()
        self.word = 0
        self.daq, self.relays = Mock(simulated=False), Mock(simulated=False)
        self.daq.analog_input.return_value = 0
        self.daq.analog_output.side_effect = lambda ch, value, vrange: self.log.append(("dac", value))
        self.daq.digital_output.side_effect = lambda port, bit, high: self.log.append(("ttl", bit, high))
        def relay(word):
            self.word = word
            self.log.append(("relay", word))
        self.relays.set_digout.side_effect = relay
        self.relays.get_digout.side_effect = lambda: self.word
        self.circuit = PulseCircuit(self.cfg, self.daq, self.relays, sleep=self.clock.sleep,
                                    monotonic=self.clock.monotonic, clock=self.clock.now)
        self.vertical, self.turning = Mock(), Mock()
        def motion(name):
            def run(*args, **kwargs):
                self.log.append((name, args))
                return SimpleNamespace(ok=True, actual=0)
            return run
        self.vertical.home_to_top.side_effect = motion("home")
        self.vertical.move_to.side_effect = motion("move")
        self.turning.rotate_to.side_effect = motion("turn")
        self.plan = plan_pulse_irm(0, coil, self.cfg)
        return PulseTreatmentService(self.circuit, self.vertical, self.turning, clock=self.clock.now)

    def execute(self, service):
        return service.execute(self.plan, sample_id="S1", run_id="R1", sample_height=3000)

    def test_two_residual_pulses_at_load_then_position_without_charging(self):
        record = self.execute(self.fixture())
        self.assertTrue(record.safe_state_confirmed)
        self.assertEqual(record.schema, "rapidpy.irm.zero_treatment.v1")
        self.assertEqual(record.circuit.schema, "rapidpy.irm.zero_field.v1")
        self.assertFalse(record.circuit.fired)
        self.assertEqual(record.circuit.residual_pulses_completed, 2)
        self.assertFalse(record.to_dict()["circuit"]["charged_pulse_fired"])
        pulses = [index for index, entry in enumerate(self.log) if entry == ("ttl", 1, False)]
        self.assertEqual(len(pulses), 2)
        move = next(index for index, entry in enumerate(self.log) if entry[0] == "move")
        self.assertLess(pulses[-1], move)
        self.assertTrue(all(entry[1] == 0 for entry in self.log if entry[0] == "dac"))
        phases = [phase.name for phase in record.circuit.phases]
        self.assertIn("zero_field_no_charge", phases)
        self.assertNotIn("charge_target", phases)
        self.assertNotIn("fire_on", phases)
        self.assertEqual(self.log[-1][0], "home")

    def test_transverse_zero_field_uses_source_three_second_residual_fire(self):
        service = self.fixture("transverse")
        record = self.execute(service)
        self.assertEqual(record.circuit.plan.fire_hold_s, 3)
        self.assertIn(("relay", 3), self.log)
        self.assertGreaterEqual(sum(self.clock.slept), 9)

    def test_direct_zero_without_load_lifecycle_fails_before_all_outputs(self):
        self.fixture()
        with self.assertRaises(PulseCircuitError) as result:
            self.circuit.execute(self.plan, sample_id="S1", run_id="R1")
        self.assertIn("load-to-coil", result.exception.record.error)
        self.assertEqual(self.log, [])

    def test_cancel_during_first_residual_pulse_never_moves_specimen_and_inhibits_fire(self):
        service = self.fixture()
        self.circuit.should_cancel = lambda: ("ttl", 1, False) in self.log
        with self.assertRaises(PulseTreatmentError) as result: self.execute(service)
        self.assertIn("cancelled", result.exception.record.error)
        self.vertical.move_to.assert_not_called()
        self.assertTrue(result.exception.record.safe_state_confirmed)
        self.assertEqual([entry for entry in self.log if entry[:2] == ("ttl", 1)][-1], ("ttl", 1, True))

    def test_unverified_discharge_never_fires_and_withholds_motor_return(self):
        service = self.fixture()
        self.daq.analog_input.return_value = .5  # 50 capacitor V, above zero-pulse threshold
        with self.assertRaises(PulseTreatmentError) as result: self.execute(service)
        self.assertFalse(result.exception.record.safe_state_confirmed)
        self.assertNotIn(("ttl", 1, False), self.log)
        self.vertical.move_to.assert_not_called()
        self.vertical.home_to_top.assert_called_once()

    def test_hot_sensor_between_residual_pulses_blocks_second_fire_and_specimen_move(self):
        service = self.fixture()
        self.cfg.temperature_channels = [2]
        self.cfg.temperature_slope = 10
        self.cfg.temperature_offset = 0
        self.cfg.temperature_hot_c = 40
        self.daq.analog_input.side_effect = lambda channel, vrange: (5 if ("ttl", 1, False) in self.log else 3) if channel == 2 else 0
        with self.assertRaises(PulseTreatmentError) as result: self.execute(service)
        self.assertIn("hot threshold", result.exception.record.error)
        self.assertEqual(self.log.count(("ttl", 1, False)), 1)
        self.assertEqual(result.exception.record.circuit.residual_pulses_completed, 1)
        self.vertical.move_to.assert_not_called()

    def test_hot_sensor_after_positioning_blocks_nonzero_charge_command(self):
        service = self.fixture()
        self.cfg.axial_calibrated = True
        self.cfg.temperature_channels = [2]
        self.cfg.temperature_slope = 10
        self.cfg.temperature_offset = 0
        self.cfg.temperature_hot_c = 40
        self.plan = plan_pulse_irm(50, "axial", self.cfg)
        self.daq.analog_input.side_effect = lambda channel, vrange: (5 if any(entry[0] == "move" for entry in self.log) else 3) if channel == 2 else 0
        with self.assertRaises(PulseTreatmentError) as result: self.execute(service)
        self.assertIn("hot threshold", result.exception.record.error)
        self.assertTrue(all(entry[1] == 0 for entry in self.log if entry[0] == "dac"))
        self.assertFalse(result.exception.record.circuit.fired)

    def test_live_queue_routes_zero_and_worker_indexes_zero_treatment_schema(self):
        from tests.test_live_treatment_validation import LiveTreatmentValidationTests
        backend = LiveTreatmentValidationTests().backend()
        backend._config.pulse_irm.axial_calibrated = False
        backend._config.pulse_irm.axial_calibration = []
        self.assertTrue(backend.validate_treatment_plan(("IRM0",)).ok)
        backend.set_demag_step("IRM0")
        record = backend.pulse_treatment_records[-1]
        self.assertEqual(record.schema, "rapidpy.irm.zero_treatment.v1")
        self.assertFalse(record.circuit.fired)
        from rapid_main.measurement_worker import MeasurementWorker
        from tests.test_susceptibility_queue import _WorkerBackend, _meta
        class Backend(_WorkerBackend):
            def __init__(self):
                super().__init__()
                self.pulse_treatment_records = []
            def set_demag_step(self, label): self.pulse_treatment_records.append(record)
        with tempfile.TemporaryDirectory() as folder:
            output = Path(folder)/"S1"
            worker = MeasurementWorker(meta=_meta("S1"), labels=["IRM0"], output_dir=output, backend=Backend())
            errors = []
            worker.error_occurred.connect(errors.append)
            worker.run()
            self.assertEqual(errors, [])
            path = output/"pulse_treatments"/f"{record.treatment_id}.json"
            self.assertEqual(json.loads(path.read_text())["schema"], "rapidpy.irm.zero_treatment.v1")
            index = json.loads((output/"artifact_index.json").read_text())
            row = next(row for row in index["artifacts"] if row["name"] == f"pulse_treatment:{record.treatment_id}")
            self.assertTrue(row["required"])
            self.assertEqual(row["sha256"], hashlib.sha256(path.read_bytes()).hexdigest())
