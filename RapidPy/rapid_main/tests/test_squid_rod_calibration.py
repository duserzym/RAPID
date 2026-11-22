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
from rapid_main.diagnostic_services import DCMotorBackendAdapter, SquidBackendAdapter
from rapid_main.legacy_ini import import_vb6_ini
from rapid_main.squid_rod_calibration import (
    SQUID_ROD_PROCEDURE, SquidRodRequest, SquidRodScanJob, coil_factors, plan_squid_rod, scan_squid_rod,
)
from rapidpy_common.hardware_safety import HardwareSafetyStore
from tests.test_queue_lift_transfer import LiftSerial


def plan_for(**changes):
    settings = dict(zero_pos=-50000, meas_pos=-97500, counts_per_cm=961,
                    axis_factors_emu=(1e-5, 1e-5, 1e-5), settle_s=1.)
    settings.update(changes)
    return plan_squid_rod(SquidRodRequest('MAG-1', 2.4, .3, 1e-6, 10.), **settings)


def response_at(position, center=-96160, width=360):
    signal = math.exp(-.5 * ((position - center) / width) ** 2)
    return 2 + signal * 12, -4 - signal * 40, 1 + signal * 22


class SquidRodTests(unittest.TestCase):
    def setUp(self):
        directory = tempfile.TemporaryDirectory()
        self.addCleanup(directory.cleanup)
        self.output = Path(directory.name)

    def test_coarse_fine_and_independent_center_use_signed_dominant_axis(self):
        plan = plan_for(axis_factors_emu=(1e-5, 2e-5, 1e-5))
        readings = []
        fit = scan_squid_rod(plan, lambda position: (position, response_at(position)), readings,
                             cancelled=lambda: False)
        self.assertEqual(fit.selected_axis, 'Y')
        self.assertLessEqual(abs(fit.fitted_rod_position + 96160), plan.fine_interval / 2)
        self.assertLess(fit.verification_emu, 0)
        self.assertEqual(fit.baseline_drift_emu, (0., 0., 0.))
        self.assertLessEqual(abs((fit.zero_position - fit.measurement_position) - 15 * plan.counts_per_cm), 1)
        self.assertEqual(readings[0].phase, 'baseline')
        self.assertEqual([reading.phase for reading in readings[-2:]], ['verify_center', 'verify_baseline'])
        self.assertGreater(min(reading.target for reading in readings), plan.meas_pos)

    def test_unbracketed_signal_and_bad_plane_are_blocked(self):
        with self.assertRaisesRegex(ValueError, 'accepted negative'):
            plan_for(zero_pos=-100000)
        plan = plan_for()
        readings = []
        with self.assertRaisesRegex(ValueError, 'peak inside'):
            scan_squid_rod(plan, lambda pos: (pos, response_at(pos, center=-99000)), readings,
                           cancelled=lambda: False)
        self.assertFalse(any(reading.phase == 'fine' for reading in readings))

    def test_failed_baseline_drift_is_saved_and_cannot_be_registered(self):
        def acquire(plan, readings, evidence, *, cancelled):
            revisits = 0

            def sample(position):
                nonlocal revisits
                x, y, z = response_at(position)
                if position == plan.baseline_position:
                    revisits += 1
                    if revisits == 3:
                        z += 1.0
                return position, (x, y, z)

            return scan_squid_rod(plan, sample, readings, cancelled=cancelled)

        job = SquidRodScanJob(plan_for(), self.output / 'drift.json', operator='Operator', config_hash='cfg', scan=acquire)
        result = job.run(cancelled=lambda: False)
        artifact = json.loads(result.artifact.read_text())
        self.assertIn('baseline drift exceeded', result.error)
        self.assertIsNone(artifact['proposal'])
        self.assertEqual(artifact['samples'][-1]['phase'], 'verify_baseline')
        registry = CalibrationRegistry(self.output / 'registry')
        with self.assertRaises(CalibrationRegistryError):
            registry.approve(result.artifact, approved_by='Approver')

    def test_native_uses_only_routed_lift_and_reads_inside_the_motion_journal(self):
        config = AppConfig()
        import_vb6_ini(config, Path(__file__).resolve().parents[3] / 'VB6/settings/Paleomag_v3.INI')
        motors = DCMotorBackendAdapter(port=config.changer.port, baud=config.changer.baud, config=config)
        motors._safety_store = HardwareSafetyStore(self.output / 'safety.json')
        squid = SquidBackendAdapter(config.squid)
        plan = plan_for(axis_factors_emu=coil_factors(config.calibration), settle_s=config.squid.settle_time)
        serial_ports = {}
        elapsed = 0.

        def tick():
            nonlocal elapsed
            elapsed += .1
            return elapsed

        def serial_factory(**kwargs):
            device = LiftSerial(**kwargs)
            serial_ports[device.port] = device
            return device

        def read_raw():
            pending = motors._safety_store.pending()
            self.assertEqual(pending['plan']['action'], 'squid_rod_scan')
            position = serial_ports['COM5'].position
            self.assertLess(position, 0)
            self.assertGreaterEqual(position, plan.meas_pos)
            return response_at(position)

        def acquire(plan, readings, evidence, *, cancelled):
            return motors.run_squid_rod_scan(plan, readings, evidence, squid=squid, config=config,
                                            operator='Operator', cancelled=cancelled)

        reader = Mock(is_connected=False)
        reader.raw_client.read_xyz_raw.side_effect = read_raw
        reader.connect.side_effect = lambda *args, **kwargs: setattr(reader, 'is_connected', True)
        reader.disconnect.side_effect = lambda: setattr(reader, 'is_connected', False)
        with (patch('rapidpy_common.hardware.serial.Serial', side_effect=serial_factory),
              patch('rapidpy_common.hardware.time.sleep'), patch('rapidpy_common.hardware.time.monotonic', side_effect=tick),
              patch('rapid_main.diagnostic_services.SquidMomentReader', return_value=reader)):
            job = SquidRodScanJob(plan, self.output / 'squid.json', operator='Operator', config_hash='cfg', scan=acquire)
            result = job.run(cancelled=lambda: False)
        self.assertEqual(result.error, '')
        reader.take_baseline.assert_not_called()
        reader.disconnect.assert_called_once()
        self.assertIsNone(motors._safety_store.pending())
        for port, device in serial_ports.items():
            self.assertFalse(device.is_open)
            self.assertEqual(any(b' 134 ' in command for command in device.writes), port == 'COM5')
        registry = CalibrationRegistry(self.output / 'registry')
        record = registry.approve(result.artifact, approved_by='Approver')
        self.assertEqual(record.procedure_id, SQUID_ROD_PROCEDURE)
        artifact = json.loads(result.artifact.read_text())
        self.assertEqual(artifact['station_evidence']['final_stopped_positions']['Up/Down'], 0)
        self.assertGreater(reader.raw_client.read_xyz_raw.call_count, len(plan.positions))


class SquidRodDesktopTests(unittest.TestCase):
    def test_calibration_center_routes_squid_simulation_and_disallows_approval(self):
        app = QtWidgets.QApplication.instance() or QtWidgets.QApplication([])
        with tempfile.TemporaryDirectory() as directory:
            config = AppConfig()
            config.general.nocomm = True
            with patch('rapid_main.app.AppConfig.load', return_value=config):
                window = MainWindow()
            panel = window._calibration_panel
            try:
                window._launch_calibrate_rod()
                panel.set_procedure('squid_rod')
                self.assertTrue(panel._rod_reference.isHidden())
                self.assertFalse(panel._rod_drift_limit.isHidden())
                panel._rod_standard.setText('MAG-1')
                panel._rod_artifact_dir.setText(directory)
                panel._rod_run.click()
                self.assertTrue(panel.rod_scan_active)
                self.assertEqual(window._ownership.owner_of('squid'), 'rod_calibration')
                end = time.monotonic() + 8
                while panel.rod_scan_active and time.monotonic() < end:
                    app.processEvents()
                    threading.Event().wait(.005)
                self.assertFalse(panel.rod_scan_active)
                self.assertIn('Proposed measurement position', panel._result.toPlainText())
                self.assertFalse(panel._approve_artifact_btn.isEnabled())
                self.assertIsNone(window._ownership.owner_of('squid'))
                artifact, = Path(directory).glob('squid-rod-*.json')
                with self.assertRaisesRegex(CalibrationRegistryError, 'passed live scan'):
                    CalibrationRegistry(Path(directory) / 'registry').approve(artifact, approved_by='Approver')
            finally:
                if panel.rod_scan_active:
                    panel.stop_rod_scan()
                    panel._rod_worker.wait(5000)
                    panel._rod_scan_finished()
                window.deleteLater()
                QtCore.QCoreApplication.sendPostedEvents(None, QtCore.QEvent.DeferredDelete)
                app.processEvents()
