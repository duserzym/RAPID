import json
import math
from pathlib import Path
import tempfile
import threading
import time
import unittest
from unittest.mock import Mock, patch

from PySide6 import QtCore, QtWidgets

from rapid_main.app import MainWindow
from rapid_main.calibration_registry import CalibrationRegistry, CalibrationRegistryError
from rapid_main.config import AppConfig
from rapid_main.diagnostic_services import DCMotorBackendAdapter, SusceptibilityBackendAdapter
from rapid_main.legacy_ini import import_vb6_ini
from rapid_main.rod_calibration import RodReading, RodScanJob, RodScanRequest, plan_rod_scan, scan_rod
from rapidpy_common.hardware_safety import HardwareSafetyStore
from tests.test_queue_lift_transfer import LiftSerial


def plan_for(request=None, **settings):
    fields = dict(coil=-22700, travel_limit=-97500, counts_per_cm=961, moment_factor_cgs=9.7914e-6)
    fields.update(settings)
    return plan_rod_scan(request or RodScanRequest('BART-1', 2, .4, .002, 10), **fields)


class RodScanFixture:
    def setUp(self):
        directory = tempfile.TemporaryDirectory()
        self.addCleanup(directory.cleanup)
        self.output = Path(directory.name)


class RodScanTests(RodScanFixture, unittest.TestCase):
    def test_plan_brackets_expected_center_and_rejects_unresolvable_or_unsafe_geometry(self):
        plan = plan_for()
        self.assertEqual(len(plan.positions), 21)
        self.assertEqual(plan.positions[10], -21739)
        self.assertGreaterEqual(plan.positions[0] - plan.positions[-1], 4 * plan.counts_per_cm)
        for changes in (dict(coil=-1000), dict(travel_limit=-23000), dict(counts_per_cm=2),
                        dict(counts_per_cm=float('inf'))):
            with self.subTest(changes=changes), self.assertRaises(ValueError):
                plan_for(**changes)
        with self.assertRaises(ValueError):
            plan_for(RodScanRequest('BART-1', 2, .001, .002, 10))

    def test_failed_independent_read_preserves_samples_and_omits_proposal(self):
        def acquire(plan, readings, evidence, *, cancelled):
            def sample(position):
                raw = 75 * math.exp(-.5 * ((position + 21547) / 500) ** 2)
                if len(readings) == len(plan.positions) + 1:
                    raw /= 2
                return RodReading(position, position, raw, 'Z OK')

            return scan_rod(plan, sample, readings, cancelled=cancelled)

        plan = plan_for()
        result = RodScanJob(plan, self.output / 'rod.json', operator='Operator', config_hash='cfg', scan=acquire
                            ).run(cancelled=lambda: False)
        payload = json.loads(result.artifact.read_text())
        self.assertIsNone(result.fit)
        self.assertIsNone(payload['proposal'])
        self.assertFalse(payload['passed'])
        self.assertIn('independent center', result.error)
        self.assertEqual(len(payload['samples']), len(plan.positions) + 2)
        self.assertEqual(payload['samples'][-2]['target'], payload['samples'][-1]['target'])

    def test_no_comm_artifact_cannot_enter_live_registry(self):
        result = RodScanJob(plan_for(), self.output / 'rod.json', operator='No-Comm', config_hash='cfg'
                            ).run(cancelled=lambda: False)
        self.assertTrue(result.simulated)
        self.assertLessEqual(abs(result.fit.coil + 22700), 1)
        self.assertAlmostEqual(result.fit.moment_factor_cgs, 9.7914e-6)
        registry = CalibrationRegistry(self.output / 'registry')
        with self.assertRaisesRegex(CalibrationRegistryError, 'passed live scan'):
            registry.approve(result.artifact, approved_by='Approver')

    def test_evidence_path_failure_precedes_connection(self):
        acquire = Mock()
        path = self.output / 'exists.json'
        path.write_text('original')
        job = RodScanJob(plan_for(), path, operator='Operator', config_hash='cfg', scan=acquire)
        with self.assertRaises(FileExistsError):
            job.run(cancelled=lambda: False)
        self.assertEqual(path.read_text(), 'original')
        acquire.assert_not_called()


class NativeRodScanTests(RodScanFixture, unittest.TestCase):
    def setUp(self):
        super().setUp()
        config = AppConfig()
        import_vb6_ini(config, Path(__file__).resolve().parents[3] / 'VB6/settings/Paleomag_v3.INI')
        config.susceptibility.enabled = True
        self.motor = DCMotorBackendAdapter(port=config.changer.port, baud=config.changer.baud, config=config)
        self.port = config.changer.port
        self.motor._safety_store = HardwareSafetyStore(self.output / 'safety.json')
        self.bridge = SusceptibilityBackendAdapter(config.susceptibility)
        self.plan = plan_for(coil=config.susceptibility.coil_position,
                             counts_per_cm=config.motor_station.lift_counts_per_cm,
                             moment_factor_cgs=config.susceptibility.moment_factor_cgs)
        self.serials = {}
        self.cancelled = False

    def serial_factory(self, **kwargs):
        device = LiftSerial(**kwargs)
        self.serials[device.port] = device
        return device

    def scan(self, *, zero_callback=None, bad_readback=False):
        def observe():
            pending = self.motor._safety_store.pending()
            self.assertEqual(pending['plan']['action'], 'susceptibility_rod_scan')
            self.assertEqual(pending['plan']['parameters']['mount_confirmed_by'], 'Operator')
            if bad_readback:
                self.serials['COM5'].slop = 100
            if zero_callback is not None:
                zero_callback()
            return 'Z OK'

        def measure():
            position = self.serials['COM5'].position
            return 50 * math.exp(-.5 * ((position + 21547) / 510) ** 2)

        def acquire(plan, readings, evidence, *, cancelled):
            return self.motor.run_rod_scan(plan, readings, evidence, bridge=self.bridge, port=self.port,
                                           operator='Operator', cancelled=cancelled)

        with (patch('rapidpy_common.hardware.serial.Serial', side_effect=self.serial_factory),
              patch('rapidpy_common.hardware.time.sleep'), patch.object(self.bridge, 'test_connection'),
              patch.object(self.bridge, 'zero', side_effect=observe),
              patch.object(self.bridge, 'measure', side_effect=measure) as measured,
              patch.object(self.bridge, 'disconnect') as closed):
            job = RodScanJob(self.plan, self.output / 'live.json', operator='Operator', config_hash='cfg', scan=acquire)
            result = job.run(cancelled=lambda: self.cancelled)
        closed.assert_called_once()
        return result, measured.call_count

    def test_routed_scan_journals_and_verifies_lift_and_calibration_without_lateral_moves(self):
        result, measures = self.scan()
        self.assertEqual(result.error, '')
        self.assertLessEqual(abs(result.fit.coil + 22508), 1)
        self.assertAlmostEqual(result.fit.moment_factor_cgs, 4e-5)
        self.assertEqual(measures, len(self.plan.positions) + 2)
        for port, device in self.serials.items():
            self.assertFalse(device.is_open)
            self.assertEqual(any(b' 134 ' in command for command in device.writes), port == 'COM5')
        self.assertIsNone(self.motor._safety_store.pending())
        payload = json.loads(result.artifact.read_text())
        self.assertEqual(payload['station_evidence']['final_stopped_positions']['Up/Down'], 0)
        self.assertTrue(payload['station_evidence']['bridge_closed'])
        self.assertEqual(len(payload['samples']), measures)
        registry = CalibrationRegistry(self.output / 'registry')
        record = registry.approve(result.artifact, approved_by='Approver')
        self.assertEqual(record.procedure_id, 'calibration/susceptibility-rod')

    def test_failed_position_readback_prevents_bridge_measure_and_leaves_no_proposal(self):
        result, measures = self.scan(bad_readback=True)
        self.assertEqual(measures, 0)
        self.assertIsNone(result.fit)
        self.assertIn('did not verify the planned scan point', result.error)
        self.assertIsNone(self.motor._safety_store.pending())

    def test_halt_after_zero_prevents_the_downward_move(self):
        result, measures = self.scan(zero_callback=lambda: setattr(self, 'cancelled', True))
        self.assertEqual(measures, 0)
        self.assertIsNone(result.fit)
        self.assertFalse(any(b' 134 ' in command for command in self.serials['COM5'].writes))
        self.assertIsNone(self.motor._safety_store.pending())

    def test_unverified_stop_preserves_original_routed_station_for_recovery(self):
        def lose_stationary_readback():
            self.cancelled = True
            self.serials['COM5'].velocity = 1

        result, measures = self.scan(zero_callback=lose_stationary_readback)
        self.assertEqual(measures, 0)
        self.assertTrue(result.recovery_required)
        self.assertIn('Recover Verify Stopped', result.error)
        self.assertTrue(self.motor.can_recover_pending)
        self.assertTrue(self.serials['COM5'].is_open)
        self.serials['COM5'].velocity = 0
        record = self.motor.recover_stop()
        self.assertTrue(record.safe_state_confirmed)
        self.assertIsNone(self.motor._safety_store.pending())
        self.motor.disconnect()


class RodScanDesktopTests(unittest.TestCase):
    def test_menu_runs_no_comm_scan_and_releases_its_device_lease(self):
        app = QtWidgets.QApplication.instance() or QtWidgets.QApplication([])
        with tempfile.TemporaryDirectory() as directory:
            config = AppConfig()
            config.general.nocomm = True
            with patch('rapid_main.app.AppConfig.load', return_value=config):
                window = MainWindow()
            panel = window._calibration_panel
            try:
                window._launch_calibrate_rod()
                self.assertEqual(panel._procedure.currentData(), 'susceptibility_rod')
                self.assertTrue(panel._mode.isHidden())
                panel._rod_standard.setText('BART-1')
                panel._rod_reference.setText('2e-3')
                panel._rod_artifact_dir.setText(directory)
                panel._rod_run.click()
                self.assertTrue(panel.rod_scan_active)
                self.assertFalse(window._queue_start_allowed())
                end = time.monotonic() + 8
                while panel.rod_scan_active and time.monotonic() < end:
                    app.processEvents()
                    threading.Event().wait(.005)
                self.assertFalse(panel.rod_scan_active)
                self.assertIn('Simulated rod scan complete', panel._status.text())
                self.assertIn('Proposed coil position', panel._result.toPlainText())
                self.assertFalse(panel._approve_artifact_btn.isEnabled())
                self.assertEqual(len(list(Path(directory).glob('susceptibility-rod-*.json'))), 1)
                self.assertTrue(window._queue_start_allowed())
            finally:
                if panel.rod_scan_active:
                    panel.stop_rod_scan()
                    panel._rod_worker.wait(5000)
                    panel._rod_scan_finished()
                window.deleteLater()
                QtCore.QCoreApplication.sendPostedEvents(None, QtCore.QEvent.DeferredDelete)
                app.processEvents()
