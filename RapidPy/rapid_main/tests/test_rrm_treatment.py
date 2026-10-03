import json
from pathlib import Path
import tempfile
from types import SimpleNamespace
import unittest
from unittest.mock import Mock

from rapid_main.config import AppConfig
from rapid_main.legacy_ini import import_vb6_ini
from rapid_main.af_treatment import AfTreatmentError, write_af_treatment_record
from rapid_main.rrm_treatment import plan_rrm_treatment, RrmTreatmentService
from tests.af_fakes import configured_af
from tests.acquisition_fakes import FakeClock


def station_config():
    cfg = AppConfig()
    import_vb6_ini(cfg, Path(__file__).resolve().parents[3] / "VB6/settings/Paleomag_v3.INI")
    cfg.af_demag = configured_af()
    return cfg


class RrmTreatmentTests(unittest.TestCase):
    def fixture(self, label="RRM50/5"):
        self.plan = plan_rrm_treatment(label, station_config())
        self.clock = FakeClock()
        self.log = []
        self.adapter = Mock(simulated=False)
        self.adapter.is_connected.return_value = True
        self.vertical, self.turning = Mock(), Mock()
        self.spinning = False
        def operation(name):
            def run(*args, **kwargs):
                self.log.append(name)
                return SimpleNamespace(ok=True, actual=0., detail="", success=True)
            return run
        self.vertical.home_to_top.side_effect = operation("home")
        self.vertical.move_to.side_effect = operation("move")
        self.turning.rotate_to.side_effect = operation("orient")
        self.turning.restore_spin_reference.side_effect = operation("reference")
        self.adapter.reset_field.side_effect = operation("reset")
        def spin(*args):
            self.log.append("spin")
            self.spinning = True
            return SimpleNamespace(success=True)
        def stop():
            self.log.append("stop")
            self.spinning = False
            return SimpleNamespace(success=True)
        self.turning.spin.side_effect = spin
        self.turning.stop_spin.side_effect = stop
        self.turning.position.side_effect = lambda: round(-self.plan.full_rotation * self.plan.speed_rps * self.clock.monotonic())
        self.check = lambda: False
        self.adapter.set_halt_check.side_effect = lambda check: setattr(self, "check", check)
        def ramp(_ramp):
            self.log.append("ramp")
            self.assertTrue(self.spinning)
            self.clock.sleep(.3)
            if self.check(): raise InterruptedError("cancelled inside ramp")
        self.adapter.apply_calibrated_af.side_effect = ramp
        return RrmTreatmentService(self.adapter, self.vertical, self.turning,
                                   sleep=self.clock.sleep, monotonic=self.clock.monotonic, clock=self.clock.now)

    def test_success_spins_during_ramp_and_stops_before_reference_and_home(self):
        service = self.fixture()
        record = service.execute(self.plan, sample_id="S1", run_id="R1")
        self.assertEqual(self.log, ["home", "move", "orient", "spin", "ramp", "reset", "stop", "reference", "home"])
        self.assertTrue(record.safe_state_confirmed)
        self.assertEqual(record.completed_passes, 1)
        self.assertTrue(any(phase.name == "rotation_readback" for phase in record.phases))
        with tempfile.TemporaryDirectory() as folder:
            path = write_af_treatment_record(Path(folder)/"rrm.json", record)
            evidence = json.loads(path.read_text())
            self.assertEqual(evidence["schema"], "rapidpy.rrm.treatment.v1")
            self.assertEqual(evidence["plan"]["speed_rps"], 5)

    def test_axial_negative_spin_and_native_limits(self):
        cfg = station_config()
        plan = plan_rrm_treatment("RRMZ50/-5 rps", cfg)
        self.assertEqual((plan.ramp.coil, plan.speed_rps), ("axial", -5))
        for label in ("RRM50", "RRM50/0", "RRM50/41", "RRM0/5", "RRM50/nan", "RRM101/5"):
            with self.subTest(label=label), self.assertRaises(ValueError): plan_rrm_treatment(label, cfg)
        cfg.motor_station.controller["turning_motor_1rps"] = 2**31
        with self.assertRaisesRegex(ValueError, "32-bit"): plan_rrm_treatment("RRM50/5", cfg)

    def test_stall_prevents_ramp_but_attempts_both_independent_cleanup_operations(self):
        service = self.fixture()
        self.turning.position.side_effect = lambda: 0
        with self.assertRaises(AfTreatmentError) as result:
            service.execute(self.plan, sample_id="S1", run_id="R1")
        self.assertIn("stalled", result.exception.record.error)
        self.adapter.apply_calibrated_af.assert_not_called()
        self.assertTrue(result.exception.record.safe_state_confirmed)
        self.assertEqual(self.log[-4:], ["reset", "stop", "reference", "home"])

    def test_stop_failure_withholds_lift_and_recovery_never_restarts_spin(self):
        service = self.fixture()
        self.turning.stop_spin.side_effect = RuntimeError("stationary readback unavailable")
        with self.assertRaises(AfTreatmentError) as result:
            service.execute(self.plan, sample_id="S1", run_id="R1")
        self.assertFalse(result.exception.record.safe_state_confirmed)
        self.assertEqual(self.vertical.home_to_top.call_count, 1)
        self.turning.restore_spin_reference.assert_not_called()
        self.turning.stop_spin.side_effect = lambda: SimpleNamespace(success=True)
        recovery = service.recover(self.plan, sample_id="S1", run_id="R1")
        self.assertTrue(recovery.safe_state_confirmed)
        self.assertEqual(recovery.schema, "rapidpy.rrm.recovery.v1")
        self.turning.spin.assert_called_once()
        self.adapter.apply_calibrated_af.assert_called_once()

    def test_field_cleanup_failure_still_stops_and_withholds_lift(self):
        service = self.fixture()
        self.adapter.reset_field.side_effect = RuntimeError("relay reset failed")
        with self.assertRaises(AfTreatmentError) as result:
            service.execute(self.plan, sample_id="S1", run_id="R1")
        self.turning.stop_spin.assert_called_once()
        self.assertFalse(result.exception.record.safe_state_confirmed)
        self.assertEqual(self.vertical.home_to_top.call_count, 1)

    def test_cancel_inside_ramp_retains_evidence_and_verifies_stop(self):
        service = self.fixture()
        start = self.clock.monotonic()
        service.should_cancel = lambda: self.clock.monotonic() - start > .4
        with self.assertRaises(AfTreatmentError) as result:
            service.execute(self.plan, sample_id="S1", run_id="R1")
        self.assertTrue(result.exception.record.safe_state_confirmed)
        self.assertIn("cancelled", result.exception.record.error)
        self.assertFalse(self.spinning)

    def test_failed_recovery_preserves_fault_and_never_moves_lift(self):
        service = self.fixture()
        self.turning.stop_spin.side_effect = RuntimeError("no stop acknowledgement")
        with self.assertRaises(AfTreatmentError) as result: service.recover(self.plan, sample_id="S1", run_id="R1")
        self.assertFalse(result.exception.record.safe_state_confirmed)
        self.vertical.home_to_top.assert_not_called()
        self.turning.spin.assert_not_called()

    def test_optional_bias_uses_independent_cleanup_and_is_retained(self):
        service = self.fixture("RRM50/5@0.05")
        bias = Mock(simulated=False)
        service.bias_adapter = bias
        record = service.execute(self.plan, sample_id="S1", run_id="R1")
        bias.set_bias_mT.assert_called_once_with(.05)
        bias.clear_bias.assert_called_once()
        self.assertEqual(record.bias_mT, .05)
        bias.clear_bias.side_effect = RuntimeError("DAC reset failed")
        with self.assertRaises(AfTreatmentError) as result: service.execute(self.plan, sample_id="S1", run_id="R1")
        self.assertFalse(result.exception.record.safe_state_confirmed)
        self.assertIn("bias_clear", result.exception.record.safe_return_errors[0])

    def test_new_rockmag_labels_retain_field_unit_and_old_artifacts_remain_readable(self):
        from rapid_main.data_model import RockmagStep
        self.assertEqual((RockmagStep.from_label("RRM50/5").value, RockmagStep.from_label("RRM50/5").unit), (50, "mT"))
        self.assertEqual(RockmagStep.from_label("RRM-0.5").unit, "rps")

    def test_queue_preflight_validates_both_parameters_and_inhibits_motion_on_fault(self):
        from rapid_main.hardware_contracts import QueueHardwareBackend, HardwareError
        backend = object.__new__(QueueHardwareBackend)
        backend._config = station_config()
        backend._measurement = None
        backend._af_demag = Mock(simulated=False)
        backend._af_demag.is_connected.return_value = True
        backend._af_treatment_records = []
        backend._pulse_treatment_records = []
        self.assertTrue(backend.validate_treatment_plan(("RRM50/5",)).ok)
        self.assertFalse(backend.validate_treatment_plan(("RRM50",)).ok)
        service = self.fixture()
        self.turning.stop_spin.side_effect = RuntimeError("no stop readback")
        with self.assertRaises(AfTreatmentError) as result: service.execute(self.plan, sample_id="S1", run_id="R1")
        backend._af_treatment_records.append(result.exception.record)
        self.assertTrue(backend.has_unresolved_rotation_fault)
        self.assertFalse(backend.validate_treatment_plan(("AFZ50",)).ok)
        with self.assertRaisesRegex(HardwareError, "inhibited"): backend._ensure_connected()

    def test_worker_publishes_rrm_schema_with_required_digest(self):
        from rapid_main.measurement_worker import MeasurementWorker
        from tests.test_susceptibility_queue import _WorkerBackend, _meta
        service = self.fixture()
        record = service.execute(self.plan, sample_id="S1", run_id="run-9")
        class Backend(_WorkerBackend):
            def __init__(self):
                super().__init__()
                self.af_treatment_records = []
            def set_demag_step(self, label):
                self.af_treatment_records.append(record)
        with tempfile.TemporaryDirectory() as folder:
            out = Path(folder)/"S1"
            worker = MeasurementWorker(meta=_meta("S1"), labels=["RRM50/5"], output_dir=out, backend=Backend(), run_id="run-9")
            errors = []
            worker.error_occurred.connect(errors.append)
            worker.run()
            self.assertEqual(errors, [])
            path = out/"af_treatments"/f"{record.treatment_id}.json"
            self.assertEqual(json.loads(path.read_text())["schema"], "rapidpy.rrm.treatment.v1")
            import hashlib
            index = json.loads((out/"artifact_index.json").read_text())
            row = next(row for row in index["artifacts"] if row["relative_path"] == f"af_treatments/{record.treatment_id}.json")
            self.assertTrue(row["required"])
            self.assertEqual(row["sha256"], hashlib.sha256(path.read_bytes()).hexdigest())

    def test_manual_reset_retains_rrm_recovery_evidence_and_is_available_without_arm_module(self):
        from rapid_main.manual_treatment import ManualArmTreatment
        service = self.fixture()
        recovery = service.recover(self.plan, sample_id="S1", run_id="R1")
        backend = Mock(has_unresolved_rotation_fault=True, _arm_bias=None, _pulse_irm=None)
        backend._af_demag = self.adapter
        backend.pulse_treatment_records = []
        backend.af_treatment_records = []
        backend.return_to_safe_state.side_effect = lambda: backend.af_treatment_records.append(recovery)
        with tempfile.TemporaryDirectory() as folder:
            manual = ManualArmTreatment(backend, folder)
            self.assertTrue(manual.is_connected())
            manual.reset()
            index_path = next(Path(folder).glob("*/artifact_index.json"))
            index = json.loads(index_path.read_text())
            evidence = json.loads((index_path.parent/index["artifacts"][0]["relative_path"]).read_text())
            self.assertEqual(evidence["schema"], "rapidpy.rrm.recovery.v1")
            self.assertTrue(evidence["safe_state_confirmed"])
