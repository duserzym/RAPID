"""Requests must fail before any live ramp or configuration mutation."""
from dataclasses import asdict
import unittest
from unittest.mock import Mock
from tests.af_fakes import configured_af
from tests.acquisition_fakes import FakeClock, FakeMotorSerialClient
from rapidpy_common.hardware import MotorAxisConfig
from tests.test_pulse_circuit import circuit_config
from rapid_main.pulse_circuit import PulseCircuit

from rapid_main.config import AppConfig, IrmArmConfig, AfDemagConfig
from rapid_main.diagnostic_services import IrmArmBackendAdapter, plan_af_demag_command, DiagnosticContractError
from rapid_main.hardware_contracts import HardwareError, QueueHardwareBackend


class LiveTreatmentValidationTests(unittest.TestCase):
    def test_axial_and_maximum_af_labels_preserve_the_requested_field(self):
        for label in ("AF25", "AFZ25", "AFMAX25"):
            with self.subTest(label=label):
                command = plan_af_demag_command(label, AfDemagConfig(peak=100))
                self.assertEqual(command.field_mT, 25)
                self.assertEqual(command.ramp_peak_voltage, 2.5)

    def test_invalid_or_over_limit_af_never_becomes_default_peak(self):
        for label in ("AFBOGUS", "AFnan", "AFinf", "AF-10", "AF101", "AFZ101"):
            with self.subTest(label=label):
                with self.assertRaises(DiagnosticContractError):
                    plan_af_demag_command(label, AfDemagConfig(peak=100))

    def test_af_plan_blocks_an_over_limit_step_before_ramping(self):
        backend = self.backend()
        result = backend.validate_treatment_plan(("AF20", "AF2000"))
        self.assertFalse(result.ok)
        self.assertIn("calibration limits", result.blockers[0])
        backend._af_demag.apply_af.assert_not_called()

    def backend(self):
        backend = object.__new__(QueueHardwareBackend)
        backend._config = AppConfig()
        backend._config.general.nocomm = False
        backend._config.af_demag = configured_af()
        backend._config.pulse_irm = circuit_config()
        cfg = backend._config.irm_arm
        cfg.arm_enabled = True
        cfg.arm_calibration_source = "test-fixture"
        cfg.arm_bias_max_mT = 1
        cfg.arm_voltage_per_mT = 2
        cfg.arm_voltage_max = 2
        cfg.arm_board = cfg.arm_dac_channel = cfg.arm_gate_bit = 0
        backend._arm_bias = Mock(simulated=False)
        backend._arm_bias.is_connected.return_value = True
        backend._config.motion.sample_top = 0
        backend._config.motion.sample_bottom = -3000
        backend._client = FakeMotorSerialClient()
        backend._client.is_connected = True
        backend._connected = True
        backend._axes = {"turning": MotorAxisConfig("Turning",2,2), "updown": MotorAxisConfig("UpDown",3,3)}
        backend._halt_check = lambda: False
        backend._acquisition_clock = FakeClock()
        backend._af_treatment_records = []
        backend._pulse_treatment_records = []
        backend._pulse_irm = Mock(simulated=False)
        backend._pulse_irm.is_connected.return_value = True
        daq,relays = Mock(simulated=False),Mock(simulated=False)
        state = {"cap":0,"word":0}
        def dac(channel,value,vrange):
            if value>0:
                state["cap"] = value/.02
        def digital(port,bit,high):
            if (bit==1 and not high) or (bit==3 and not high):
                state["cap"] = 0
        daq.analog_output.side_effect = dac
        daq.digital_output.side_effect = digital
        daq.analog_input.side_effect = lambda *args:state["cap"]*.01
        relays.set_digout.side_effect = lambda word:state.update(word=word)
        relays.get_digout.side_effect = lambda:state["word"]
        backend._pulse_irm.make_circuit.side_effect = lambda **options:PulseCircuit(backend._config.pulse_irm,daq,relays,**options)
        backend._sample_name = "S1"
        backend._run_id = "R1"
        backend._measurement = None
        backend._af_demag = Mock(simulated=False)
        backend._af_demag.is_connected.return_value = True
        backend._irm_arm = Mock(simulated=False)
        backend._irm_arm.is_connected.return_value = True
        backend._treatment_label = "NRM"
        return backend

    def test_missing_simulated_disconnected_and_unreadable_actuators_block_before_calls(self):
        for kind, attribute in (("AF20", "_af_demag"), ("IRM20", "_pulse_irm"), ("ARM20", "_arm_bias")):
            for fault in ("missing", "simulated", "disconnected", "unreadable", "no-method"):
                with self.subTest(kind=kind, fault=fault):
                    backend = self.backend()
                    actuator = getattr(backend, attribute)
                    if fault == "missing":
                        setattr(backend, attribute, None)
                    elif fault == "simulated":
                        actuator.simulated = True
                    elif fault == "disconnected":
                        actuator.is_connected.return_value = False
                    elif fault == "unreadable":
                        actuator.is_connected.side_effect = RuntimeError("connection lost")
                    else:
                        setattr(backend, attribute, object())
                    self.assertFalse(backend.validate_treatment_plan((kind,)).ok)
                    with self.assertRaises(HardwareError):
                        backend.set_demag_step(kind)
                    self.assertEqual(backend._treatment_label, "NRM")
                    actuator.apply_af.assert_not_called()
                    actuator.apply_irm.assert_not_called()
                    actuator.apply_arm.assert_not_called()

    def test_whole_plan_rejects_overvoltage_before_earlier_steps_execute(self):
        backend = self.backend()
        result = backend.validate_treatment_plan(("NRM", "IRM20", "IRM2000"))
        self.assertFalse(result.ok)
        self.assertIn("IRM2000", result.blockers[0])
        backend._irm_arm.apply_irm.assert_not_called()
        with self.assertRaises(HardwareError):
            backend.set_demag_step("IRM2000")
        backend._irm_arm.apply_irm.assert_not_called()

    def test_valid_live_routes_still_execute(self):
        backend = self.backend()
        self.assertTrue(backend.validate_treatment_plan(("AF20", "IRM30", "ARM20_0.5")).ok)
        for label in ("AF20", "IRM30", "ARM20_0.5"):
            backend.set_demag_step(label)
        self.assertEqual(backend._af_demag.apply_calibrated_af.call_count, 4)
        self.assertEqual(backend.af_treatment_records[0].completed_passes, 3)
        backend._irm_arm.apply_irm.assert_not_called()
        backend._pulse_irm.make_circuit.assert_called_once()
        self.assertTrue(backend.pulse_treatment_records[0].safe_state_confirmed)
        backend._irm_arm.apply_arm.assert_not_called()
        backend._arm_bias.set_bias_mT.assert_called_once_with(.5)
        backend._arm_bias.clear_bias.assert_called_once()
        self.assertEqual(backend.af_treatment_records[1].bias_mT, .5)

    def test_invalid_irm_never_mutates_config_or_sends_partial_ramps(self):
        for field, steps in ((-20, 1), (float("nan"), 1), (float("inf"), 1), (2000, 2), (20, 0), (20, 1.5)):
            with self.subTest(field=field, steps=steps):
                config = IrmArmConfig()
                before = asdict(config)
                controller = Mock()
                controller.test_version.return_value = 91
                adapter = IrmArmBackendAdapter(config, controller=controller)
                with self.assertRaises((ValueError, RuntimeError)):
                    adapter.apply_irm(max_field_mT=field, axis="Z", ramp_label="Fast (10 s)", steps=steps)
                controller.run_ramp.assert_not_called()
                self.assertEqual(asdict(config), before)
                self.assertEqual(adapter.communication_events(), ())

    def test_invalid_arm_never_mutates_or_clamps_a_request(self):
        for peak, bias in ((-20, 0.5), (10, 20), (float("inf"), 0), (20, float("nan"))):
            with self.subTest(peak=peak, bias=bias):
                config = IrmArmConfig()
                before = asdict(config)
                controller = Mock()
                controller.test_version.return_value = 91
                adapter = IrmArmBackendAdapter(config, controller=controller)
                with self.assertRaises(ValueError):
                    adapter.apply_arm(peak_af_mT=peak, bias_mT=bias, steps=2)
                controller.run_ramp.assert_not_called()
                self.assertEqual(asdict(config), before)
