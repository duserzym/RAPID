from contextlib import nullcontext
import json
import math
from pathlib import Path
import tempfile
import threading
import time
from types import SimpleNamespace
import unittest
from unittest.mock import Mock, patch

from PySide6 import QtCore, QtWidgets

from rapid_main.app import MainWindow
from rapid_main.calibration_registry import CalibrationRegistry, CalibrationRegistryError
from rapid_main.config import AppConfig
from rapid_main.diagnostic_services import SusceptibilityBackendAdapter
from rapid_main.rod_calibration import RodScanJob, RodScanRequest, plan_rod_scan
from rapid_main.squid_rod_calibration import SquidRodRequest, SquidRodScanJob, coil_factors, plan_squid_rod
from rapid_main.tray_rod import TrayRodReservation, TrayRodScan, validate_tray_standard
from tests.test_queue_instruments import ScientificSerial
from tests.test_queue_startup import QueueStartupFixture


class TrayRodNativeTests(QueueStartupFixture, unittest.TestCase):
    def setUp(self):
        super().setUp()
        config = self.backend._config
        config.motor_station.lift_counts_per_cm = 961.
        config.susceptibility.enabled = True
        config.susceptibility.port = 'COM7'
        self.backend._susceptibility = SusceptibilityBackendAdapter(config.susceptibility)
        self.backend._susceptibility_records = []

    def plan(self, *, squid=False, height=2800 / 961, slot=1):
        cfg = self.backend._config
        if squid:
            return plan_squid_rod(SquidRodRequest('STD', height, .3, 1e-6, 10, tray_slot=slot),
                zero_pos=cfg.motion.zero_pos, meas_pos=cfg.motion.meas_pos,
                counts_per_cm=cfg.motor_station.lift_counts_per_cm,
                axis_factors_emu=coil_factors(cfg.calibration), settle_s=cfg.squid.settle_time)
        return plan_rod_scan(RodScanRequest('STD', height, .4, .002, 10, tray_slot=slot),
            coil=cfg.susceptibility.coil_position, travel_limit=cfg.motion.meas_pos,
            counts_per_cm=cfg.motor_station.lift_counts_per_cm,
            moment_factor_cgs=cfg.susceptibility.moment_factor_cgs)

    def run_calibration(self, *, squid=False, chain=False, height=2800 / 961, fail_read=False):
        if chain:
            self.use_chain_station()
        plan = self.plan(squid=squid, height=height)
        scan = TrayRodScan(self.backend, self.vacuum, plan, operator='Operator', chain_slot=1 if chain else None)
        prepare_lifetime = self.backend.prepare_queue_lifetime
        instrument = self.backend._measurement if squid else self.backend._susceptibility
        log = []

        def prepare(*args, **kwargs):
            original = prepare_lifetime(*args, **kwargs)
            self.session = original.session
            self.addCleanup(self.session.release)
            original.table.sleep = self.clock.sleep
            return original

        def connect():
            root = self.session.store._queue(self.session.token)
            self.assertEqual(root['stage']['plan']['purpose'], 'rod_calibration')
            self.assertEqual(root['stage']['plan']['specimen_geometry']['original_slot'], 1)
            self.assertTrue(self.vacuum.is_valve_connected())
            kind = 'squid' if squid else 'susceptibility'
            scan.startup.instruments._client(kind)._serial = ScientificSerial(self, kind)

        def read():
            self.assertTrue(self.vacuum.is_valve_connected())
            position = self.lift_serial().position
            log.append(position)
            if fail_read:
                raise OSError('injected instrument read failure')
            if squid:
                predicted = plan.meas_pos + plan.standard_center_offset
                signal = math.exp(-.5 * ((position - predicted) / 320) ** 2)
                return (2 * signal, -3 * signal, 5 * signal)
            predicted = plan.coil - plan.center_correction
            return .002 / plan.moment_factor_cgs * math.exp(-.5 * ((position - predicted) / 600) ** 2)

        connect_api = 'connect_for_acquisition' if squid else 'test_connection'
        read_api = 'read_raw_xyz' if squid else 'measure'
        job_type = SquidRodScanJob if squid else RodScanJob
        with (patch.object(self.backend, 'prepare_queue_lifetime', side_effect=prepare),
              patch.object(instrument, connect_api, side_effect=connect) as opened,
              patch.object(instrument, read_api, side_effect=read),
              patch('rapid_main.tray_rod.time.monotonic', side_effect=self.clock.monotonic)):
            with (nullcontext() if chain else self.arm_xy_edges()), (
                    nullcontext() if squid else patch.object(instrument, 'zero', return_value='Z OK')):
                job = job_type(plan, self.path.with_name('calibration.json'), operator='Operator', config_hash='cfg', scan=scan)
                result = job.run(cancelled=lambda: False)
        return result, log, opened.call_count

    def test_xy_bartington_loads_measures_and_returns_original_slot_before_approval(self):
        result, readings, opened = self.run_calibration()
        self.assertEqual(result.error, '')
        self.assertIsNotNone(result.fit)
        self.assertTrue(readings)
        self.assertEqual(opened, 1)
        self.assertIsNone(self.store.pending())
        payload = json.loads(result.artifact.read_text())
        evidence = payload['station_evidence']['tray_transfer']
        self.assertEqual((evidence['loaded']['phase'], evidence['returned']['phase']), ('lifted', 'clear'))
        self.assertEqual(evidence['returned']['original_slot'], 1)
        self.assertEqual(evidence['loaded']['sample_height'], 2800)
        self.assertTrue(evidence['shutdown']['safe_state_confirmed'])
        registry = CalibrationRegistry(self.path.parent / 'registry')
        record = registry.approve(result.artifact, approved_by='Approver')
        self.assertEqual(record.procedure_id, 'calibration/susceptibility-rod')
        evidence['returned']['original_slot'] = 2
        bad = result.artifact.with_name('tampered.json')
        bad.write_text(json.dumps(payload))
        with self.assertRaisesRegex(CalibrationRegistryError, 'original-slot return'):
            registry.approve(bad, approved_by='Approver')

    def test_chain_squid_uses_confirmed_reference_then_returns_loaded_standard(self):
        result, receipts, opened = self.run_calibration(squid=True, chain=True)
        self.assertEqual(result.error, '')
        self.assertEqual(opened, 1)
        self.assertTrue(receipts)
        state = self.store.read()
        self.assertEqual(state['plan']['startup']['chain_slot'], 1)
        self.assertEqual(state['status'], 'verified')
        self.assertAlmostEqual(result.fit.measurement_position, -97500, delta=100)

    def test_height_mismatch_returns_standard_without_instrument_io(self):
        result, readings, opened = self.run_calibration(height=2.5)
        self.assertIn('disagrees with its certified height', result.error)
        self.assertEqual((readings, opened), ([], 0))
        self.assertIsNone(result.fit)
        payload = json.loads(result.artifact.read_text())
        self.assertEqual(payload['station_evidence']['tray_transfer']['returned']['phase'], 'clear')
        self.assertIsNone(self.store.pending())

    def test_transport_fault_retains_grip_and_the_original_unfinished_stage(self):
        result, readings, opened = self.run_calibration(fail_read=True)
        self.assertIn('injected instrument read failure', result.error)
        self.assertEqual((len(readings), opened), (1, 1))
        self.assertIsNone(result.fit)
        self.assertTrue(self.vacuum.is_valve_connected())
        pending = self.store.pending()
        self.assertEqual(pending['stage']['family'], 'acquisition')
        self.assertEqual(pending['stage']['status'], 'pending')
        payload = json.loads(result.artifact.read_text())
        self.assertEqual(payload['station_evidence']['pending_queue_token'], pending['token'])

    def test_station_empty_hole_is_rejected_as_a_standard_source(self):
        with self.assertRaisesRegex(ValueError, 'registered nonempty slot'):
            validate_tray_standard(self.backend._config, self.plan(slot=46))


class TrayRodDesktopTests(unittest.TestCase):
    def test_unresolved_scan_transfers_application_lease_to_native_recovery(self):
        startup = SimpleNamespace(session=object())
        lease = Mock()
        scan = SimpleNamespace(startup=startup, backend=SimpleNamespace(_queue_terminal_cleanup=None))
        window = SimpleNamespace(_release_queue_lease=Mock(return_value=False), _shutdown_cleanup_requested=False,
                                 set_flow_state=Mock(), set_status=Mock())
        reservation = TrayRodReservation(scan, lease, lambda scan, lease: MainWindow._settle_tray_rod_scan(window, scan, lease))
        reservation.release()
        reservation.release()
        self.assertIs(window._queue_native_session, startup.session)
        self.assertIs(window._queue_lease, lease)
        window._release_queue_lease.assert_called_once()
        lease.release.assert_not_called()
        window.set_flow_state.assert_called_once_with('error')
        self.assertIn('Shutdown Interrupted Queue', window.set_status.call_args.args[0])

    def test_no_comm_rod_form_records_the_selected_standard_slot(self):
        app = QtWidgets.QApplication.instance() or QtWidgets.QApplication([])
        with tempfile.TemporaryDirectory() as directory:
            config = AppConfig()
            config.general.nocomm = True
            with patch('rapid_main.app.AppConfig.load', return_value=config):
                window = MainWindow()
            panel = window._calibration_panel
            try:
                window._launch_calibrate_rod()
                panel._rod_standard.setText('STD')
                panel._rod_reference.setText('0.002')
                panel._rod_tray_slot.setValue(2)
                panel._rod_artifact_dir.setText(directory)
                panel._rod_run.click()
                end = time.monotonic() + 10
                while panel.rod_scan_active and time.monotonic() < end:
                    app.processEvents()
                    threading.Event().wait(.005)
                self.assertFalse(panel.rod_scan_active)
                self.assertFalse(panel._approve_artifact_btn.isEnabled())
                artifact, = Path(directory).glob('susceptibility-rod-*.json')
                data = json.loads(artifact.read_text())
                self.assertTrue(data['simulated'])
                self.assertEqual(data['plan']['request']['tray_slot'], 2)
            finally:
                if panel.rod_scan_active:
                    panel.stop_rod_scan()
                    panel._rod_worker.wait(5000)
                    panel._rod_scan_finished()
                window.deleteLater()
                QtCore.QCoreApplication.sendPostedEvents(None, QtCore.QEvent.DeferredDelete)
                app.processEvents()
