import csv
from dataclasses import replace
import json
import os
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
from rapid_main.uchannel import (
    CoreGeometry, EprTubeGeometry, UChannelJob, UChannelRequest, plan_uchannel, plan_uchannel_from_config, scan_uchannel,
)
from rapidpy_common.hardware_safety import HardwareSafetyStore
from tests.test_queue_lift_transfer import LiftSerial


def request_for(**changes):
    request = UChannelRequest('UC-1', CoreGeometry(10, -78000, -79000, -73000), 12, .5, 'up', False, 2, 1e-6)
    return replace(request, **changes)


def plan_for(request=None, **changes):
    settings = dict(meas_pos=-97500, carriage_limit=-120000, floor_limit=-148955, counts_per_cm=1000,
                    axis_factors_emu=(1e-5, 2e-5, -3e-5), range_factor=1e-5, settle_s=1.)
    settings.update(changes)
    return plan_uchannel(request or request_for(), **settings)


class UChannelTests(unittest.TestCase):
    def setUp(self):
        directory = tempfile.TemporaryDirectory()
        self.addCleanup(directory.cleanup)
        self.output = Path(directory.name)

    def test_epr_positions_follow_tube_geometry_and_add_a_five_cm_zero(self):
        geometry = EprTubeGeometry(10., 2., 7., 1., 1.)
        plan = plan_for(request_for(geometry=geometry, four_passes=True))
        self.assertEqual((plan.positions[0], plan.positions[-1], plan.zero_target), (-88500, -95500, -83500))
        self.assertEqual(plan.angles, (0, 90, 180, 270))
        self.assertEqual(plan.safe_depth, -120000)
        self.assertEqual(plan.reading_budget, 96)

    def test_carriage_floor_extension_and_top_margin_each_block_invalid_plans(self):
        for changes in (dict(carriage_limit=-78800), dict(floor_limit=-90000), dict(floor_limit=None)):
            with self.subTest(changes=changes), self.assertRaises(ValueError):
                plan_for(**changes)
        for request in (request_for(assembly_extension_cm=9),
                        request_for(geometry=CoreGeometry(10., -78000, -79000, -1)),
                        request_for(geometry=EprTubeGeometry(10., 7., 2., 1., 1.))):
            with self.subTest(request=request), self.assertRaises(ValueError):
                plan_for(request)

    def test_four_passes_bracket_each_group_and_remove_linear_drift_in_emu(self):
        plan = plan_for(request_for(four_passes=True, drift_limit_emu=1e-4))
        readings, turns = [], []
        ticks = iter(range(100))

        def sample(position, angle):
            offset = len(readings) / 20
            signal = 0 if position == plan.zero_target else 2
            return position, float(angle), (offset + signal, 2 * offset + signal, -offset + signal)

        result = scan_uchannel(plan, sample, turns.append, readings, cancelled=lambda: False, clock=lambda: next(ticks))
        self.assertEqual(turns, [0, 90, 180, 270, 0])
        self.assertEqual(len(readings), plan.reading_budget)
        self.assertEqual(len(result.profile), len(plan.positions) * 4)
        self.assertEqual([row.phase for row in readings[:6]], ['zero', 'sample', 'sample', 'zero', 'sample', 'zero'])
        for point in result.profile:
            for measured, expected in zip(point.moment_emu, (2e-5, 4e-5, -6e-5)):
                self.assertAlmostEqual(measured, expected, places=12)

    def test_failed_drift_saves_raw_trace_without_a_corrected_profile(self):
        def acquire(plan, readings, evidence, *, cancelled):
            def sample(position, angle):
                return position, float(angle), (float(len(readings)), 0., 0.)

            return scan_uchannel(plan, sample, lambda angle: None, readings, cancelled=cancelled)

        result = UChannelJob(plan_for(), self.output / 'failed', operator='operator',
                             config_hash='cfg', scan=acquire).run(cancelled=lambda: False)
        self.assertIn('baseline drift exceeded', result.error)
        self.assertIsNone(result.fit)
        payload = json.loads(result.artifact.read_text())
        self.assertEqual(payload['status'], 'failed')
        self.assertEqual([row['phase'] for row in payload['readings']], ['zero', 'sample', 'sample', 'zero'])
        self.assertTrue(result.artifact.with_name('readings.csv').is_file())
        self.assertFalse(result.artifact.with_name('profile.csv').exists())

    def test_trace_and_csv_units_are_compatible_with_legacy_output(self):
        result = UChannelJob(plan_for(), self.output / 'simulated', operator='No-Comm',
                             config_hash='cfg').run(cancelled=lambda: False)
        self.assertEqual(result.error, '')
        with result.artifact.with_name('readings.csv').open(newline='') as handle:
            rows = list(csv.reader(handle))
        self.assertEqual(rows[0], ['Long Core Scan'])
        header = rows.index(['Requested Up/Down Pos', 'Actual Up/Down Pos', 'Time (sec)', 'Turn Angle',
                             'SQUID X', 'SQUID Y', 'SQUID Z'])
        self.assertEqual(tuple(map(float, rows[header + 1][4:])), (2., 6., -12.))
        with result.artifact.with_name('profile.csv').open(newline='') as handle:
            self.assertEqual(len(list(csv.DictReader(handle))), len(result.fit.profile))
        with self.assertRaisesRegex(CalibrationRegistryError, 'specimen measurements'):
            CalibrationRegistry(self.output / 'registry').approve(result.artifact, approved_by='Approver')

    def test_interrupted_process_leaves_progress_and_samples_without_publishing(self):
        def acquire(plan, readings, evidence, *, cancelled):
            def sample(position, angle):
                if len(readings) == 3:
                    raise KeyboardInterrupt
                return position, float(angle), (1., 2., 3.)

            return scan_uchannel(plan, sample, lambda angle: None, readings, cancelled=cancelled)

        progress = []
        job = UChannelJob(plan_for(), self.output / 'incomplete', operator='Operator', config_hash='cfg', scan=acquire)
        job.report_progress = progress.append
        with self.assertRaises(KeyboardInterrupt):
            job.run(cancelled=lambda: False)
        self.assertEqual(progress, [1, 2, 3])
        self.assertFalse(job.path.exists())
        stage, = self.output.glob('.uchannel-pending-*')
        manifest = json.loads((stage / 'scan.json').read_text())
        self.assertEqual((manifest['status'], manifest['output_name']), ('running', 'incomplete'))
        trace = [json.loads(line) for line in (stage / 'readings.jsonl').read_text().splitlines()]
        self.assertEqual([reading['phase'] for reading in trace], ['zero', 'sample', 'sample'])
        self.assertFalse((stage / 'profile.csv').exists())

    def test_failed_bundle_promotion_keeps_its_files_and_names_their_location(self):
        job = UChannelJob(plan_for(), self.output / 'complete', operator='No-Comm', config_hash='cfg')
        replace_path = os.replace

        def replace_with_locked_destination(source, destination):
            if destination == job.path:
                raise PermissionError('destination is locked')
            replace_path(source, destination)

        with (patch('rapid_main.uchannel.os.replace', side_effect=replace_with_locked_destination),
              self.assertRaisesRegex(OSError, r'Inspect .*\.uchannel-pending-.*destination is locked')):
            job.run(cancelled=lambda: False)
        self.assertFalse(job.path.exists())
        stage, = self.output.glob('.uchannel-pending-*')
        manifest = json.loads((stage / 'scan.json').read_text())
        self.assertEqual((manifest['status'], manifest['output_name']), ('complete', 'complete'))
        self.assertTrue((stage / 'profile.csv').is_file())
        self.assertTrue((stage / 'readings.csv').is_file())


class UChannelNativeTests(unittest.TestCase):
    def setUp(self):
        directory = tempfile.TemporaryDirectory()
        self.addCleanup(directory.cleanup)
        self.output = Path(directory.name)
        self.config = AppConfig()
        import_vb6_ini(self.config, Path(__file__).resolve().parents[3] / 'VB6/settings/Paleomag_v3.INI')
        self.motors = DCMotorBackendAdapter(port=self.config.changer.port, baud=self.config.changer.baud, config=self.config)
        self.motors._safety_store = HardwareSafetyStore(self.output / 'safety.json')
        self.squid = SquidBackendAdapter(self.config.squid)

    def run_scan(self, *, four_passes, start_angle_counts=0, slop=0):
        plan = plan_uchannel_from_config(request_for(four_passes=four_passes), self.config)
        serial_ports, turns = {}, []
        elapsed = 0.

        def tick():
            nonlocal elapsed
            elapsed += .1
            return elapsed

        def write_turn(command):
            if command.startswith('134 '):
                lift = serial_ports['COM5']
                self.assertEqual((lift.position, lift.status), (0, 16))
                turns.append(int(command.split()[1]))

        def serial_factory(**kwargs):
            device = LiftSerial(**kwargs)
            if device.port == 'COM6':
                device.position = start_angle_counts
                device.on_write = write_turn
            if device.port == 'COM5':
                device.slop = slop
            serial_ports[device.port] = device
            return device

        def read_xyz():
            pending = self.motors._safety_store.pending()
            self.assertEqual(pending['plan']['action'], 'uchannel_scan')
            self.assertEqual(pending['plan']['parameters']['plan']['floor_limit'], -148955)
            position = serial_ports['COM5'].position
            self.assertGreaterEqual(position, plan.safe_depth)
            return (0., 0., 0.) if position == plan.zero_target else (2., 3., 4.)

        def scan(plan, readings, evidence, *, cancelled):
            return self.motors.run_uchannel_scan(plan, readings, evidence, squid=self.squid, config=self.config,
                                                 operator='Operator', cancelled=cancelled)

        reader = Mock(is_connected=False)
        reader.raw_client.read_xyz_raw.side_effect = read_xyz
        reader.connect.side_effect = lambda *args, **kwargs: setattr(reader, 'is_connected', True)
        reader.disconnect.side_effect = lambda: setattr(reader, 'is_connected', False)
        with (patch('rapidpy_common.hardware.serial.Serial', side_effect=serial_factory),
              patch('rapidpy_common.hardware.time.sleep'), patch('rapidpy_common.hardware.time.monotonic', side_effect=tick),
              patch('rapid_main.diagnostic_services.SquidMomentReader', return_value=reader)):
            result = UChannelJob(plan, self.output / 'native', operator='Operator',
                                 config_hash='cfg', scan=scan).run(cancelled=lambda: False)
        return result, reader, serial_ports, turns

    def test_four_live_passes_turn_at_top_and_restore_reference_under_a_journal(self):
        result, reader, ports, turns = self.run_scan(four_passes=True)
        self.assertEqual(result.error, '')
        self.assertEqual(turns, [2000, 4000, 6000, 0])
        self.assertFalse(any(b' 134 ' in command for port in ('COM3', 'COM4') for command in ports[port].writes))
        self.assertIsNone(self.motors._safety_store.pending())
        self.assertTrue(all(not device.is_open for device in ports.values()))
        reader.disconnect.assert_called_once()
        payload = json.loads(result.artifact.read_text())
        self.assertEqual(payload['station_evidence']['final_stopped_positions']['Turning'], 0)
        self.assertEqual(payload['station_evidence']['final_stopped_positions']['Up/Down'], 0)
        self.assertEqual(reader.raw_client.read_xyz_raw.call_count, len(payload['readings']))

    def test_bad_reference_blocks_before_any_motion_or_squid_connection(self):
        result, reader, ports, turns = self.run_scan(four_passes=True, start_angle_counts=1000)
        self.assertIn('0° turning reference', result.error)
        self.assertEqual(turns, [])
        self.assertFalse(any(b' 134 ' in command for device in ports.values() for command in device.writes))
        reader.connect.assert_not_called()
        reader.raw_client.read_xyz_raw.assert_not_called()

    def test_wrong_lift_readback_stops_before_squid_read(self):
        result, reader, ports, turns = self.run_scan(four_passes=False, slop=140)
        self.assertIn('did not verify the planned scan point', result.error)
        self.assertEqual(turns, [])
        reader.raw_client.read_xyz_raw.assert_not_called()
        reader.disconnect.assert_called_once()


class UChannelDesktopTests(unittest.TestCase):
    def test_diagnostics_opens_epr_scan_that_exports_and_releases_squid(self):
        app = QtWidgets.QApplication.instance() or QtWidgets.QApplication([])
        with tempfile.TemporaryDirectory() as directory:
            config = AppConfig()
            config.general.nocomm = True
            with patch('rapid_main.app.AppConfig.load', return_value=config):
                window = MainWindow()
            panel = window._calibration_panel
            try:
                window._launch_uchannel_scan()
                self.assertEqual(panel._procedure.currentData(), 'uchannel')
                panel._channel_kind.setCurrentIndex(1)
                self.assertTrue(panel._core_geometry.isHidden())
                panel._channel_specimen.setText('EPR-1')
                panel._channel_extension.setValue(12)
                panel._epr_length.setValue(10)
                panel._epr_bottom.setValue(2)
                panel._epr_top.setValue(7)
                panel._epr_below.setValue(1)
                panel._epr_above.setValue(1)
                panel._rod_artifact_dir.setText(directory)
                panel._rod_run.click()
                self.assertTrue(panel.rod_scan_active)
                self.assertEqual(window._ownership.owner_of('squid'), 'rod_calibration')
                end = time.monotonic() + 10
                while panel.rod_scan_active and time.monotonic() < end:
                    app.processEvents()
                    threading.Event().wait(.005)
                self.assertFalse(panel.rod_scan_active)
                self.assertIn('Simulated U-channel scan complete', panel._status.text())
                self.assertIn('Profile CSV', panel._result.toPlainText())
                self.assertFalse(panel._approve_artifact_btn.isEnabled())
                self.assertIsNone(window._ownership.owner_of('squid'))
                artifact, = Path(directory).glob('uchannel-*/scan.json')
                self.assertEqual(json.loads(artifact.read_text())['plan']['request']['specimen'], 'EPR-1')
            finally:
                if panel.rod_scan_active:
                    panel.stop_rod_scan()
                    panel._rod_worker.wait(5000)
                    panel._rod_scan_finished()
                window.deleteLater()
                QtCore.QCoreApplication.sendPostedEvents(None, QtCore.QEvent.DeferredDelete)
                app.processEvents()
