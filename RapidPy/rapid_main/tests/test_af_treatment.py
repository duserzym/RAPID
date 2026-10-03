"""AF physical ordering, calibration limits, fault recovery and durable evidence."""
import json
import hashlib
from pathlib import Path
import tempfile
from types import SimpleNamespace
import unittest
from unittest.mock import Mock
from rapidpy_common.adwin_af import AdwinAFController, AdwinBoardConfig, AdwinRampRequest, AdwinError

from rapid_main.af_treatment import (AfTreatmentService, AfTreatmentError,
    plan_af_treatment, plan_calibrated_af_ramp, write_af_treatment_record)
from tests.af_fakes import configured_af
from tests.acquisition_fakes import FakeClock


class AfTreatmentTests(unittest.TestCase):
    def service(self):
        self.journal = []
        self.clock = FakeClock()
        vertical, turning, adapter = Mock(), Mock(), Mock(simulated=False)
        adapter.is_connected.return_value = True
        def operation(name):
            def invoke(*args, **kwargs):
                self.journal.append((name, args))
                return SimpleNamespace(ok=True, actual=0, target=0, detail="")
            return invoke
        vertical.home_to_top.side_effect = operation("home")
        vertical.move_to.side_effect = operation("move")
        turning.rotate_to.side_effect = operation("rotate")
        adapter.apply_calibrated_af.side_effect = operation("ramp")
        adapter.reset_field.side_effect = operation("reset")
        self.vertical, self.turning, self.adapter = vertical, turning, adapter
        return AfTreatmentService(adapter, vertical, turning, sleep=self.clock.sleep, clock=self.clock.now)

    def test_separate_coil_calibration_and_legacy_ramp_mode(self):
        axial = plan_calibrated_af_ramp("AF50", 50, "axial", configured_af())
        transverse = plan_calibrated_af_ramp("AF50", 50, "transverse", configured_af())
        self.assertEqual((axial.monitor_peak_v, axial.ramp_peak_v, axial.frequency_hz), (.5, .5, 100))
        self.assertEqual((transverse.monitor_peak_v, transverse.ramp_peak_v, transverse.frequency_hz), (1, 1.25, 200))
        self.assertEqual((axial.slope_up_vps, axial.slope_down_vps), (1, .5))
        self.assertEqual((transverse.slope_up_vps, transverse.slope_down_vps), (2.5, 1))
        self.assertEqual(axial.ramp_mode, 2)

    def test_pass_order_and_specimen_centre(self):
        plan = plan_af_treatment("AF50", configured_af(), 3000)
        self.assertEqual(plan.target_position, -18500)
        self.assertEqual(plan_af_treatment("AF50", configured_af(), 3001).target_position, -18500)
        with self.assertRaises(ValueError):
            plan_af_treatment("AF50", configured_af(), -3000)
        self.assertEqual([(p.angle_deg, p.ramp.coil) for p in plan.passes], [(0,"axial"),(0,"transverse"),(90,"transverse")])
        maximum = plan_af_treatment("AFMAX", configured_af(), 3000)
        self.assertEqual([(p.angle_deg,p.ramp.field_mT) for p in maximum.passes], [(0,100),(90,100),(360,200)])
        self.assertEqual(len(plan_af_treatment("AFZ50", configured_af(), 3000).passes), 1)

    def test_complete_sequence_confirms_safe_return_and_retains_evidence(self):
        service = self.service()
        record = service.execute(plan_af_treatment("AF50", configured_af(), 3000), sample_id="S1", run_id="R1")
        self.assertEqual([name for name,_ in self.journal], ["home","move","rotate","ramp","rotate","ramp","rotate","ramp","rotate","reset","home"])
        self.assertAlmostEqual(sum(self.clock.slept), 3)
        self.assertEqual(record.completed_passes, 3)
        self.assertTrue(record.safe_state_confirmed)
        with tempfile.TemporaryDirectory() as folder:
            target = Path(folder)/"record.json"
            write_af_treatment_record(target, record)
            self.assertEqual(json.loads(target.read_text())["schema"], "rapidpy.af.treatment.v1")
            with self.assertRaises(FileExistsError):
                write_af_treatment_record(target, record)
            self.assertEqual(list(Path(folder).glob("*.tmp-*")), [])

    def test_failed_ramp_attempts_both_cleanup_operations(self):
        service = self.service()
        self.adapter.apply_calibrated_af.side_effect = RuntimeError("ramp timeout")
        self.adapter.reset_field.side_effect = RuntimeError("relay fault")
        with self.assertRaises(AfTreatmentError) as result:
            service.execute(plan_af_treatment("AF50", configured_af(), 3000), sample_id="S1", run_id="R1")
        record = result.exception.record
        self.assertIn("ramp timeout", record.error)
        self.assertIn("relay fault", record.safe_return_errors[0])
        self.assertFalse(record.safe_state_confirmed)
        self.assertEqual(self.vertical.home_to_top.call_count, 2)

    def test_cancel_during_pause_stops_second_ramp_and_returns_home(self):
        service = self.service()
        service.should_cancel = lambda: bool(self.clock.slept)
        with self.assertRaises(AfTreatmentError) as result:
            service.execute(plan_af_treatment("AF50", configured_af(), 3000), sample_id="S1", run_id="R1")
        self.assertEqual(result.exception.record.completed_passes, 1)
        self.assertTrue(result.exception.record.safe_state_confirmed)
        self.adapter.apply_calibrated_af.assert_called_once()

    def test_invalid_configuration_and_limits_fail_before_motion(self):
        for overrides in ({"enabled":False}, {"axial_calibrated":False},
                          {"axial_calibration":[[1],[2,200]]},
                          {"axial_calibration":[[1,100],[.5,200]]},
                          {"io_rate_hz":100}, {"coil_position":0}):
            with self.subTest(overrides=overrides), self.assertRaises(ValueError):
                plan_af_treatment("AF50", configured_af(**overrides), 3000)
        for field in (-1, float("nan"), .1, 201):
            with self.subTest(field=field), self.assertRaises(ValueError):
                plan_calibrated_af_ramp("AF", field, "axial", configured_af())

    def test_plateau_voltage_and_bounded_extrapolation_are_supported(self):
        cfg = configured_af(axial_calibration=[[1,100],[1,200]], axial_max_mT=250)
        self.assertEqual(plan_calibrated_af_ramp("AF250",250,"axial",cfg).monitor_peak_v, 1)

    def test_worker_publishes_only_current_run_records_with_digest(self):
        from rapid_main.measurement_worker import MeasurementWorker
        from tests.test_susceptibility_queue import _WorkerBackend, _meta
        service = self.service()
        plan = plan_af_treatment("AFZ50", configured_af(), 3000)
        before = service.execute(plan, sample_id="S1", run_id="old-run")
        current = service.execute(plan, sample_id="S1", run_id="run-9")
        class Backend(_WorkerBackend):
            def __init__(self):
                super().__init__()
                self.af_treatment_records = [before]
            def set_demag_step(self, label):
                super().set_demag_step(label)
                self.af_treatment_records.append(current)
        with tempfile.TemporaryDirectory() as folder:
            out = Path(folder)/"S1"
            worker = MeasurementWorker(meta=_meta("S1"), labels=["AFZ50"], output_dir=out, backend=Backend(), run_id="run-9")
            errors = []
            worker.error_occurred.connect(errors.append)
            worker.run()
            self.assertEqual(errors, [])
            artifact = out/"af_treatments"/f"{current.treatment_id}.json"
            self.assertTrue(artifact.exists())
            self.assertFalse((artifact.parent/f"{before.treatment_id}.json").exists())
            index = json.loads((out/"artifact_index.json").read_text())
            row = next(row for row in index["artifacts"] if row["relative_path"] == f"af_treatments/{current.treatment_id}.json")
            self.assertTrue(row["required"])
            self.assertEqual(row["sha256"], hashlib.sha256(artifact.read_bytes()).hexdigest())
            workflow = json.loads((out/"workflow_summary.json").read_text())
            self.assertEqual(workflow["af_treatment_ids"], [current.treatment_id])

    def controller(self):
        controller = object.__new__(AdwinAFController)
        controller.board = AdwinBoardConfig()
        controller._dll = Mock()
        controller._dll.ADB_Start.return_value = 0
        controller._dll.ADB_Stop.return_value = 0
        for name in ("boot_board", "set_af_relays", "clear_all_processes", "load_process", "set_fpar", "set_par"):
            setattr(controller, name, Mock())
        controller._coil_limits = Mock(return_value=(10,10))
        controller.get_par = Mock(return_value=0)
        return controller

    def request(self):
        return AdwinRampRequest(slope_up=1, slope_down=1, peak_monitor_voltage=1,
                                sine_freq_hz=100, ramp_peak_voltage=1, active_coil="axial")

    def test_native_cancellation_stops_started_process(self):
        controller = self.controller()
        cancellation = Mock(side_effect=[False, True])
        with self.assertRaisesRegex(AdwinError, "cancelled while running"):
            controller.run_ramp(self.request(), should_cancel=cancellation)
        controller._dll.ADB_Stop.assert_called_once_with(controller.board.process_num, 1)

    def test_native_cancel_before_initialization_does_not_start_or_boot(self):
        controller = self.controller()
        with self.assertRaisesRegex(AdwinError, "before initialization"):
            controller.run_ramp(self.request(), should_cancel=lambda: True)
        controller.boot_board.assert_not_called()
        controller._dll.ADB_Start.assert_not_called()

    def test_native_status_fault_stops_process_and_preserves_stop_failure(self):
        controller = self.controller()
        controller.get_par.side_effect = RuntimeError("status read lost")
        controller._dll.ADB_Stop.return_value = 9
        with self.assertRaisesRegex(AdwinError, "status read lost; process stop failed with code 9"):
            controller.run_ramp(self.request())

    def test_adapter_rejects_changed_calibration_before_ramp(self):
        from rapid_main.diagnostic_services import AfDemagBackendAdapter
        cfg = configured_af()
        ramp = plan_calibrated_af_ramp("AF50",50,"axial",cfg)
        controller = Mock()
        controller.test_version.return_value = 91
        adapter = AfDemagBackendAdapter(cfg, controller=controller)
        cfg.axial_calibration = [[1.5,100],[2.5,200]]
        with self.assertRaisesRegex(RuntimeError, "changed after treatment planning"):
            adapter.apply_calibrated_af(ramp)
        controller.run_ramp.assert_not_called()

    def test_legacy_station_profile_supplies_calibration_units_and_relays(self):
        from rapid_main.config import AppConfig
        from rapid_main.legacy_ini import import_vb6_ini
        cfg = AppConfig()
        source = Path(__file__).resolve().parents[3]/"VB6"/"settings"/"Paleomag_v3.INI"
        import_vb6_ini(cfg, source)
        plan = plan_af_treatment("AF25", cfg.af_demag, cfg.motion.sample_height)
        self.assertEqual(plan.target_position, -41000)
        self.assertEqual((cfg.af_demag.axial_relay_bit,cfg.af_demag.transverse_relay_bit), (1,2))
        self.assertEqual(cfg.af_demag.axial_max_mT, 282)
        self.assertEqual(len(cfg.af_demag.axial_calibration), 19)
        self.assertEqual([p.ramp.frequency_hz for p in plan.passes], [877,316,316])
        self.assertAlmostEqual(plan.passes[0].ramp.ramp_peak_v, .4022694809115531)
        self.assertAlmostEqual(plan.passes[1].ramp.ramp_peak_v, .7361870767627976)

    def test_legacy_cross_board_relays_cannot_execute(self):
        from rapid_main.config import AppConfig
        from rapid_main.legacy_ini import import_vb6_ini
        cfg = AppConfig()
        cfg.af_demag = configured_af()
        with tempfile.TemporaryDirectory() as folder:
            source = Path(folder)/"relays.ini"
            source.write_text("[Channels]\nAFAxialRelay=DO-1-CH1\nAFTransRelay=DO-2-CH1\n[Boards]\nCommProtocol1=2\nBoardNum1=1\nDO-1-CH1=DIGOUT-1,1\nCommProtocol2=2\nBoardNum2=2\nDO-2-CH1=DIGOUT-2,2\n")
            report = import_vb6_ini(cfg, source)
        self.assertTrue(any("same ADwin board" in warning for warning in report.warnings))
        with self.assertRaisesRegex(ValueError, "relay bits"):
            plan_af_treatment("AF50", cfg.af_demag, 3000)
