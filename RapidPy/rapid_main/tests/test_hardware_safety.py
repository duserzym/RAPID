from dataclasses import asdict
import json
import os
from pathlib import Path
import tempfile
from types import SimpleNamespace
import unittest
from unittest.mock import Mock, patch

from rapid_main.hardware_safety import HardwareSafetyStore, HardwareSafetyError, default_safety_path
from rapid_main.hardware_contracts import QueueHardwareBackend
from rapid_main.config import AppConfig


def record(safe=True, simulated=False):
    return SimpleNamespace(safe_state_confirmed=safe, simulated=simulated,
                           to_dict=lambda: {"safe_state_confirmed": safe, "simulated": simulated})


class SafetyStoreTests(unittest.TestCase):
    def setUp(self):
        self.directory = tempfile.TemporaryDirectory()
        self.addCleanup(self.directory.cleanup)
        self.path = Path(self.directory.name) / "safety.json"
        self.store = HardwareSafetyStore(self.path)
        self.profile = {"motor": {"port": "COM5"}, "relay": 2}

    def begin(self):
        return self.store.begin("pulse", {"field": 10}, self.profile, sample_id="specimen", run_id="run")

    def test_restart_retains_pending_identity_plan_and_profile(self):
        token = self.begin()
        pending = HardwareSafetyStore(self.path).pending(self.profile)
        self.assertEqual(pending["token"], token)
        self.assertEqual(pending["sample_id"], "specimen")
        self.assertEqual(pending["plan"], {"field": 10})
        with self.assertRaises(HardwareSafetyError):
            HardwareSafetyStore(self.path).begin("rrm", {}, self.profile)

    def test_only_physical_verified_record_clears_then_new_operation_can_begin(self):
        token = self.begin()
        self.store.finish(token, self.profile, record(False))
        self.assertIsNotNone(self.store.pending())
        self.store.finish(token, self.profile, record(True, True))
        self.assertIsNotNone(self.store.pending())
        self.store.finish(token, self.profile, record())
        self.assertIsNone(HardwareSafetyStore(self.path).pending())
        new_token = self.begin()
        self.assertNotEqual(new_token, token)
        with self.assertRaises(HardwareSafetyError):
            self.store.finish(token, self.profile, record())
        self.assertEqual(self.store.pending()["token"], new_token)

    def test_profile_change_cannot_clear_or_recover_original_station(self):
        token = self.begin()
        with self.assertRaises(HardwareSafetyError):
            self.store.finish(token, {"motor": {"port": "COM9"}}, record())
        with self.assertRaises(HardwareSafetyError):
            self.store.pending({"motor": {"port": "COM9"}})
        self.assertIsNotNone(self.store.pending())

    def test_corruption_fails_closed_without_replacing_file(self):
        for payload in ("{", "null", '{"state":NaN}', '{"state":{},"sha256":"wrong"}'):
            self.path.write_text(payload)
            with self.assertRaises(HardwareSafetyError):
                self.begin()
            self.assertEqual(self.path.read_text(), payload)

    def test_failed_publication_preserves_pending_on_restart(self):
        token = self.begin()
        with patch("rapid_main.hardware_safety.os.replace", side_effect=OSError("disk failure")):
            with self.assertRaises(HardwareSafetyError):
                self.store.finish(token, self.profile, record())
        self.assertIsNotNone(HardwareSafetyStore(self.path).pending())
        self.assertEqual(list(self.path.parent.glob("*.tmp-*")), [])

    def test_begin_flush_failure_does_not_publish_or_start_operation(self):
        with patch("rapid_main.hardware_safety.os.fsync", side_effect=OSError("disk full")):
            with self.assertRaises(HardwareSafetyError):
                self.begin()
        self.assertFalse(self.path.exists())

    def test_config_destination_does_not_change_default_latch(self):
        with patch.dict(os.environ, {"RAPID_CONFIG": "different/config.json"}):
            before = default_safety_path()
            os.environ["RAPID_CONFIG"] = "another/config.json"
            self.assertEqual(default_safety_path(), before)

    def test_concurrent_instances_cannot_begin_while_locked(self):
        with self.store._locked():
            with self.assertRaises(HardwareSafetyError):
                HardwareSafetyStore(self.path).begin("pulse", {}, self.profile)
        self.begin()

    def test_lifetime_lease_blocks_recovery_until_owner_exits(self):
        with self.store.operation_lease():
            token = self.begin()
            with self.assertRaises(HardwareSafetyError):
                with HardwareSafetyStore(self.path).operation_lease():
                    self.fail("A second hardware owner entered the operation lease.")
            # Journal transactions remain possible within the lifetime lease.
            self.store.finish(token, self.profile, record(False))
        with HardwareSafetyStore(self.path).operation_lease():
            self.assertIsNotNone(self.store.pending())

    def test_exception_releases_lifetime_lease_but_preserves_pending(self):
        with self.assertRaises(KeyboardInterrupt):
            with self.store.operation_lease():
                self.begin()
                raise KeyboardInterrupt()
        with HardwareSafetyStore(self.path).operation_lease():
            self.assertIsNotNone(self.store.pending())


class QueueRestartSafetyTests(unittest.TestCase):
    def setUp(self):
        self.directory = tempfile.TemporaryDirectory()
        self.addCleanup(self.directory.cleanup)
        self.path = Path(self.directory.name) / "safety.json"
        self.backend = self.make_backend()

    def make_backend(self):
        backend = object.__new__(QueueHardwareBackend)
        backend._config = AppConfig()
        backend._config.general.nocomm = False
        backend._safety_store = HardwareSafetyStore(self.path)
        backend._sample_name, backend._run_id = "specimen", "run"
        backend._pulse_treatment_records, backend._af_treatment_records = [], []
        backend._pulse_irm = Mock()
        backend._pulse_irm.make_circuit.return_value.recover_safe_state.return_value = record()
        backend._acquisition_clock = None
        backend._halt_check = None
        backend._arm_bias = None
        backend._client = Mock()
        backend._connected = False
        return backend

    def test_restart_blocks_motion_before_any_motor_io(self):
        self.backend._begin_safety_operation("pulse", self.backend._config.pulse_irm)
        restarted = self.make_backend()
        self.assertTrue(restarted.has_unresolved_pulse_fault)
        with self.assertRaises(RuntimeError):
            restarted._ensure_connected()
        restarted._client.assert_not_called()
        self.assertEqual(restarted._client.mock_calls, [])

    def test_restart_discharge_precedes_verified_motion_and_clears_without_charge(self):
        self.backend._begin_safety_operation("pulse", self.backend._config.pulse_irm)
        restarted = self.make_backend()
        circuit = restarted._pulse_irm.make_circuit.return_value
        def discharge(**_kwargs):
            self.assertEqual(restarted._client.mock_calls, [])
            return record()
        circuit.recover_safe_state.side_effect = discharge
        restarted._axes = {name: Mock() for name in ("turning", "updown", "changer_x", "changer_y")}
        restarted._sample_loaded = False
        restarted._motor_communication_logger = Mock()
        restarted._af_demag = Mock()
        vertical, turning = Mock(), Mock()
        confirmed = SimpleNamespace(ok=True, actual=0., success=True)
        turning.stop_spin.return_value = confirmed
        turning.restore_spin_reference.return_value = confirmed
        vertical.home_to_top.return_value = confirmed
        with patch("rapid_main.squid_transport.MotorVerticalController", return_value=vertical), patch("rapid_main.squid_transport.MotorTurningController", return_value=turning):
            restarted.return_to_safe_state()
        circuit.recover_safe_state.assert_called_once_with(sample_id="specimen", run_id="run")
        turning.spin.assert_not_called()
        turning.stop_spin.assert_called_once()
        vertical.home_to_top.assert_called_once()
        self.assertFalse(self.make_backend().has_unresolved_pulse_fault)

    def test_pulse_discharge_without_verified_home_remains_pending(self):
        self.backend._begin_safety_operation("pulse", self.backend._config.pulse_irm)
        restarted = self.make_backend()
        restarted._axes = {"turning": Mock(), "updown": Mock()}
        vertical, turning = Mock(), Mock()
        turning.stop_spin.return_value = SimpleNamespace(success=True)
        turning.restore_spin_reference.return_value = SimpleNamespace(ok=True, actual=0.)
        vertical.home_to_top.return_value = SimpleNamespace(ok=False, actual=10.)
        with patch("rapid_main.squid_transport.MotorVerticalController", return_value=vertical), patch("rapid_main.squid_transport.MotorTurningController", return_value=turning):
            with self.assertRaisesRegex(RuntimeError, "motion recovery is unverified"):
                restarted.return_to_safe_state()
        self.assertTrue(self.make_backend().has_unresolved_pulse_fault)

    def test_recovery_io_failure_retains_latch_and_inhibits_all_motion(self):
        self.backend._begin_safety_operation("pulse", self.backend._config.pulse_irm)
        restarted = self.make_backend()
        restarted._pulse_irm.make_circuit.return_value.recover_safe_state.side_effect = OSError("DAQ lost")
        with self.assertRaises(OSError):
            restarted.return_to_safe_state()
        self.assertTrue(self.make_backend().has_unresolved_pulse_fault)
        self.assertEqual(restarted._client.mock_calls, [])

    def test_other_process_cannot_recover_active_treatment(self):
        self.backend._begin_safety_operation("pulse", self.backend._config.pulse_irm)
        restarted = self.make_backend()
        with self.backend._safety_store.operation_lease():
            with self.assertRaises(HardwareSafetyError):
                restarted.return_to_safe_state()
        self.assertEqual(restarted._pulse_irm.mock_calls, [])
        self.assertEqual(restarted._client.mock_calls, [])

    def test_changed_wiring_blocks_recovery_before_daq_io(self):
        self.backend._begin_safety_operation("pulse", self.backend._config.pulse_irm)
        restarted = self.make_backend()
        restarted._config.pulse_irm.board += 1
        with self.assertRaises(HardwareSafetyError):
            restarted.return_to_safe_state()
        self.assertEqual(restarted._pulse_irm.mock_calls, [])

    def test_corrupt_latch_blocks_pulse_and_rotation_without_io(self):
        self.path.write_text("corrupt")
        self.assertTrue(self.backend.has_unresolved_pulse_fault)
        self.assertTrue(self.backend.has_unresolved_rotation_fault)
        with self.assertRaises(HardwareSafetyError):
            self.backend.return_to_safe_state()
        self.assertEqual(self.backend._client.mock_calls, [])

    def test_nocomm_does_not_clear_live_latch(self):
        self.backend._begin_safety_operation("pulse", self.backend._config.pulse_irm)
        self.backend._config.general.nocomm = True
        self.backend.return_to_safe_state()
        self.assertIsNotNone(self.backend._safety_store.pending())

    def pulse_execution_fixture(self):
        from tests.test_pulse_circuit import circuit_config
        self.backend._config.pulse_irm = circuit_config()
        self.backend._config.motion.sample_top = 10
        self.backend._config.motion.sample_bottom = 0
        self.backend._axes = {"turning": Mock(), "updown": Mock()}
        self.backend._ensure_connected = Mock()
        return patch("rapid_main.pulse_treatment.PulseTreatmentService")

    def test_execution_persists_latch_before_first_treatment_action_and_interrupt(self):
        with self.pulse_execution_fixture() as service:
            def interrupted(*args, **kwargs):
                self.assertIsNotNone(HardwareSafetyStore(self.path).pending())
                raise KeyboardInterrupt()
            service.return_value.execute.side_effect = interrupted
            with self.assertRaises(KeyboardInterrupt):
                self.backend._execute_pulse_treatment(50, axis="Z")
        self.assertTrue(self.make_backend().has_unresolved_pulse_fault)

    def test_latch_write_failure_prevents_treatment_execute(self):
        with self.pulse_execution_fixture() as service:
            with patch("rapid_main.hardware_safety.os.replace", side_effect=OSError("disk full")):
                with self.assertRaises(HardwareSafetyError):
                    self.backend._execute_pulse_treatment(50, axis="Z")
            service.return_value.execute.assert_not_called()

    def test_successful_physical_treatment_clears_latch(self):
        with self.pulse_execution_fixture() as service:
            service.return_value.execute.return_value = record()
            self.backend._execute_pulse_treatment(50, axis="Z")
        self.assertIsNone(HardwareSafetyStore(self.path).pending())

    def test_rrm_restart_recovers_original_plan_without_spin_or_ramp(self):
        from tests.test_rrm_treatment import station_config
        from rapid_main.rrm_treatment import plan_rrm_treatment
        self.backend._config = station_config()
        plan = plan_rrm_treatment("RRM50/-5", self.backend._config)
        self.backend._begin_safety_operation("rrm", plan)
        restarted = self.make_backend()
        restarted._config = self.backend._config
        restarted._af_demag = Mock()
        restarted._axes = {name: Mock() for name in ("turning", "updown", "changer_x", "changer_y")}
        restarted._client.is_connected = False
        restarted._sample_loaded = False
        restarted._motor_communication_logger = Mock()
        vertical, turning = Mock(), Mock()
        confirmed = SimpleNamespace(ok=True, actual=0., success=True)
        turning.stop_spin.return_value = confirmed
        turning.restore_spin_reference.return_value = confirmed
        vertical.home_to_top.return_value = confirmed
        with patch("rapid_main.squid_transport.MotorVerticalController", return_value=vertical), patch("rapid_main.squid_transport.MotorTurningController", return_value=turning):
            restarted.return_to_safe_state()
        self.assertIsNone(restarted._safety_store.pending())
        self.assertEqual(restarted.af_treatment_records[-1].plan, plan)
        self.assertEqual(restarted.af_treatment_records[-1].sample_id, "specimen")
        turning.stop_spin.assert_called_once()
        turning.spin.assert_not_called()
        restarted._af_demag.apply_calibrated_af.assert_not_called()
        restarted._client.connect.assert_called_once_with(restarted._config.changer.port, baudrate=restarted._config.motor_station.baud)

    def test_failed_rrm_stationary_readback_persists_latch_and_withholds_lift(self):
        from tests.test_rrm_treatment import station_config
        from rapid_main.rrm_treatment import plan_rrm_treatment
        self.backend._config = station_config()
        self.backend._begin_safety_operation("rrm", plan_rrm_treatment("RRM50/5", self.backend._config))
        self.backend._af_demag = Mock()
        self.backend._axes = {"turning": Mock(), "updown": Mock()}
        vertical, turning = Mock(), Mock()
        turning.stop_spin.return_value = SimpleNamespace(success=False)
        with patch("rapid_main.squid_transport.MotorVerticalController", return_value=vertical), patch("rapid_main.squid_transport.MotorTurningController", return_value=turning):
            with self.assertRaisesRegex(RuntimeError, "recovery failed"):
                self.backend.return_to_safe_state()
        vertical.home_to_top.assert_not_called()
        self.assertTrue(self.make_backend().has_unresolved_rotation_fault)

    def test_af_and_arm_restart_recovery_never_reapply_field_or_bias(self):
        from tests.af_fakes import configured_af
        from rapid_main.af_treatment import plan_af_treatment
        for family in ("af", "arm"):
            with self.subTest(family=family):
                backend = self.make_backend()
                backend._config.af_demag = configured_af()
                backend._config.motion.sample_top = 10
                plan = plan_af_treatment("AFZ50", backend._config.af_demag, 10)
                backend._begin_safety_operation(family, plan, bias_mT=.05 if family == "arm" else None)
                restarted = self.make_backend()
                restarted._config = backend._config
                restarted._af_demag, restarted._arm_bias = Mock(), Mock()
                restarted._client.is_connected = False
                restarted._axes = {name: Mock() for name in ("turning", "updown", "changer_x", "changer_y")}
                restarted._sample_loaded = False
                restarted._motor_communication_logger = Mock()
                self.assertTrue(restarted.has_unresolved_hardware_fault)
                with self.assertRaisesRegex(RuntimeError, "inhibited"):
                    restarted._ensure_connected()
                vertical, turning = Mock(), Mock()
                confirmed = SimpleNamespace(ok=True, actual=0., success=True)
                turning.stop_spin.return_value = confirmed
                turning.restore_spin_reference.return_value = confirmed
                vertical.home_to_top.return_value = confirmed
                with patch("rapid_main.squid_transport.MotorVerticalController", return_value=vertical), patch("rapid_main.squid_transport.MotorTurningController", return_value=turning):
                    restarted.return_to_safe_state()
                self.assertIsNone(restarted._safety_store.pending())
                self.assertEqual(restarted.af_treatment_records[-1].plan, plan)
                self.assertEqual(restarted.af_treatment_records[-1].schema, f"rapidpy.{family}.recovery.v1")
                restarted._af_demag.apply_calibrated_af.assert_not_called()
                restarted._arm_bias.set_bias_mT.assert_not_called()
                turning.spin.assert_not_called()

    def test_arm_clear_failure_retains_durable_latch_and_withholds_motion(self):
        from tests.af_fakes import configured_af
        from rapid_main.af_treatment import plan_af_treatment
        self.backend._config.af_demag = configured_af()
        plan = plan_af_treatment("AFZ50", self.backend._config.af_demag, 10)
        self.backend._begin_safety_operation("arm", plan, bias_mT=.05)
        self.backend._af_demag = Mock()
        self.backend._arm_bias = Mock()
        self.backend._arm_bias.clear_bias.side_effect = RuntimeError("bias cannot clear")
        with self.assertRaisesRegex(RuntimeError, "bias cannot clear"):
            self.backend.return_to_safe_state()
        self.backend._af_demag.recover_field.assert_called_once()
        self.assertEqual(self.backend._client.mock_calls, [])
        self.assertTrue(self.make_backend().has_unresolved_hardware_fault)


class NativeRestartRecoveryTests(unittest.TestCase):
    def controller(self):
        from rapidpy_common.adwin_af import AdwinAFController, AdwinBoardConfig
        controller = object.__new__(AdwinAFController)
        controller.board = AdwinBoardConfig()
        controller._dll = Mock()
        controller._dll.ADB_Stop.return_value = 0
        controller._raise_if_error = Mock()
        controller.get_par = Mock(return_value=0)
        controller.get_digout = Mock(return_value=0)
        controller.set_digout, controller.set_dac, controller.boot_board = Mock(), Mock(), Mock()
        return controller

    def test_stop_running_process_verify_zero_output_and_relays_without_boot(self):
        controller = self.controller()
        statuses = iter([1, 0] + [0] * 18)
        controller.get_par.side_effect = lambda _index: next(statuses)
        controller.recover_safe_field()
        controller._dll.ADB_Stop.assert_called_once_with(1, controller._dev)
        controller.set_dac.assert_called_once_with(1, 0.0)
        controller.set_digout.assert_called_once_with(0)
        controller.boot_board.assert_not_called()

    def test_still_running_process_inhibits_relay_change(self):
        controller = self.controller()
        controller.get_par.return_value = 1
        with self.assertRaises(RuntimeError):
            controller.recover_safe_field()
        controller.set_digout.assert_not_called()
        controller.set_dac.assert_not_called()

    def test_unverified_relay_clear_raises(self):
        controller = self.controller()
        controller.get_digout.return_value = 1
        with self.assertRaises(RuntimeError):
            controller.recover_safe_field()
