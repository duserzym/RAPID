"""Imported station diagnostics through the routed motor client and fake serial."""
from pathlib import Path
import tempfile
import unittest
from unittest.mock import patch

from PySide6 import QtCore, QtWidgets
from rapid_main.config import AppConfig
from rapid_main.diagnostic_services import DCMotorBackendAdapter, HardwareUnavailableError
from rapid_main.dialogs.dc_motors import DCMotorDialog
from rapid_main.legacy_ini import import_vb6_ini
from rapidpy_common.hardware import HardwareError, sample_dropoff_position
from rapidpy_common.hardware_safety import HardwareSafetyStore
from tests.test_queue_lift_transfer import LiftSerial


class StationDiagnosticsTests(unittest.TestCase):
    def setUp(self):
        temporary = tempfile.TemporaryDirectory()
        self.addCleanup(temporary.cleanup)
        self.store = HardwareSafetyStore(Path(temporary.name) / 'diagnostic.json')
        self.config = AppConfig()
        import_vb6_ini(self.config, Path(__file__).resolve().parents[3] / 'VB6/settings/Paleomag_v3.INI')
        self.config.motion.sample_bottom = -5000
        self.config.motion.sample_top = -2000
        self.config.motion.meas_pos = -97500
        for item in (patch('rapidpy_common.hardware.serial.Serial', side_effect=LiftSerial),
                     patch('rapidpy_common.hardware.time.sleep', return_value=None),
                     patch('rapidpy_common.motor_diagnostic_safety.time.sleep', return_value=None)):
            item.start()
            self.addCleanup(item.stop)

    def connect(self, *, xy):
        station = self.config.motor_station
        station.use_xy_table = xy
        if xy is False:
            station.hole_slot = 10
            station.controller['sample_hole_alignment_offset'] = -.02
            station.ports.pop('changer_y')
            station.addresses.pop('changer_y')
        backend = DCMotorBackendAdapter(port='COM3', baud=9600, config=self.config)
        backend._safety_store = self.store
        backend.connect('COM3', 9600)
        self.addCleanup(backend.disconnect)
        self.backend = backend
        return backend

    def serial(self, port):
        return self.backend._client._connections[port]._serial

    def axes_movements(self, port):
        return [command for command in self.serial(port).writes if b' 134 ' in command or b' 135 ' in command]

    def home_xy(self):
        for port, negative, positive in (('COM3', 4, 5), ('COM4', 5, 6)):
            transport = self.serial(port)
            transport.status = (1 << negative) | (1 << positive)

            def edge(command, transport=transport, negative=negative, positive=positive):
                if command.startswith('135 '):
                    distance = int(command.split()[1])
                    transport.position = -1000 if distance < 0 else 1000
                    transport.status = 1 << (positive if distance < 0 else negative)
            transport.on_write = edge
        self.assertEqual(self.backend.home_xy_to_center(empty_rod_confirmed=True), (0, 0))

    def test_chain_reference_and_wrapped_slot_motion_use_three_original_ports(self):
        backend = self.connect(xy=False)
        self.assertEqual(backend.available_axes(), ('Changer (X)', 'Turning', 'Up/Down'))
        self.assertEqual(set(backend._client._connections), {'COM3', 'COM5', 'COM6'})
        chain = self.serial('COM3')
        before = list(chain.writes)
        with self.assertRaisesRegex(HardwareError, 'Establish the changer reference'):
            backend.goto_hole(hole=1)
        with self.assertRaisesRegex(HardwareError, 'Confirm the empty rod'):
            backend.reference_chain_slot(hole=99)
        self.assertEqual(chain.writes, before)
        self.assertIsNone(self.store.read())

        def record_relabel(command):
            if command.startswith('11 10 '):
                pending = self.store.pending()
                self.assertEqual(pending['plan']['action'], 'reference_chain_slot')
                self.assertEqual(pending['plan']['parameters']['slot'], 99)
        chain.on_write = record_relabel
        self.assertEqual(backend.reference_chain_slot(hole=99, empty_rod_confirmed=True), 99)
        self.assertFalse(self.axes_movements('COM3'))
        self.assertEqual(backend.read_hole('Changer (X)'), 99.)
        start = chain.position
        self.assertEqual(backend.goto_hole(hole=1), 1.)
        self.assertEqual(backend.read_hole('Changer (X)'), 1.)
        self.assertLess(abs(chain.position - start), 2100)
        self.assertEqual(len(self.axes_movements('COM3')), 1)
        self.assertFalse(self.axes_movements('COM5'))
        record = self.store.read()['record']
        self.assertFalse(record['operation']['parameters']['geometry']['use_xy_table'])
        self.assertEqual(record['operation']['axes'], ['Changer (X)', 'Up/Down'])
        backend.disconnect()
        backend.connect('COM3', 9600)
        self.assertFalse(backend.can_address_slots)

    def test_xy_slot_commands_use_calibrated_pair_and_reject_y_slop(self):
        backend = self.connect(xy=True)
        x_target, y_target = backend._slot_geometry.xy_target(1)
        with self.assertRaisesRegex(HardwareError, 'before using calibrated slots'):
            backend.read_hole('Changer (X)')
        with self.assertRaisesRegex(HardwareError, 'Confirm the empty rod'):
            backend.home_xy_to_center()
        self.home_xy()
        self.assertTrue(backend.can_address_slots)
        self.assertEqual(backend.move_xy_to_corner(), (-1000, -1000))
        self.assertTrue(backend.can_address_slots)
        self.assertEqual(backend.goto_hole(hole=1), 1.)
        self.assertIn(f' 134 {x_target} '.encode(), self.axes_movements('COM3')[-1])
        self.assertIn(f' 134 {y_target} '.encode(), self.axes_movements('COM4')[-1])
        self.assertEqual(backend.read_hole('Changer (X)'), 1.)
        self.serial('COM4').position += 500
        with self.assertRaises(ValueError):
            backend.read_hole('Changer (X)')
        self.serial('COM4').slop = 500
        with self.assertRaises(ValueError):
            backend.goto_hole(hole=1)
        self.assertFalse(backend.can_address_slots)
        self.assertEqual(self.store.read()['record']['operation']['axes'], ['Changer (X)', 'Changer (Y)', 'Up/Down'])
        self.assertIsNone(self.store.pending())

    def test_lift_clearance_blocks_registered_and_raw_lateral_motion(self):
        backend = self.connect(xy=False)
        backend.reference_chain_slot(hole=99, empty_rod_confirmed=True)
        lift = self.serial('COM5')
        lift.position, lift.status = -2000, 0
        for move in (lambda: backend.goto_hole(hole=1), lambda: backend.move_motor('Changer (X)', target=0, speed=800)):
            with self.assertRaisesRegex(HardwareError, 'lift clearance'):
                move()
        self.assertFalse(self.axes_movements('COM3'))
        self.assertFalse(self.axes_movements('COM5'))
        self.assertIsNone(self.store.pending())

    def test_chain_pickup_uses_contact_and_dropoff_uses_chain_height(self):
        backend = self.connect(xy=False)
        backend.reference_chain_slot(hole=99, empty_rod_confirmed=True)
        self.assertEqual(backend.sample_pickup(), (-5000, -2000, True))
        backend.home_to_top()
        with self.assertRaisesRegex(HardwareError, 'must match the imported station'):
            backend.sample_dropoff(use_xy_table=True)
        self.assertEqual(backend.sample_dropoff(), (-1700, -1700, True))
        self.assertFalse(self.store.read()['record']['operation']['parameters']['use_xy_table'])

    def test_recovery_after_failed_reference_publication_never_relabels_or_moves(self):
        backend = self.connect(xy=False)
        with patch('rapidpy_common.motor_diagnostic_safety.publish_diagnostic_record', side_effect=OSError('disk full')):
            with self.assertRaisesRegex(OSError, 'disk full'):
                backend.reference_chain_slot(hole=99, empty_rod_confirmed=True)
        self.assertFalse(backend.can_address_slots)
        self.assertIsNotNone(self.store.pending())
        backend.disconnect()
        backend.connect('COM3', 9600)
        chain = self.serial('COM3')
        self.assertTrue(backend.recover_stop().safe_state_confirmed)
        self.assertFalse(any(b' 134 ' in command or b' 135 ' in command or b' 145' in command
                             for command in chain.writes))
        self.assertFalse(backend.can_address_slots)

    def test_unknown_imported_mode_preserves_raw_diagnostics_but_blocks_slot_math(self):
        backend = self.connect(xy=None)
        self.assertIsNone(backend.slot_range)
        with self.assertRaisesRegex(HardwareError, 'table mode'):
            backend.goto_hole(hole=1)
        self.assertFalse(self.axes_movements('COM3'))
        self.assertEqual(backend.move_motor('Changer (X)', target=200, speed=800), (200, 200, True))

    def test_invalid_imported_baud_cannot_enable_legacy_slot_addressing(self):
        for invalid in (None, True, '57600', 0):
            self.config.motor_station.baud = invalid
            with self.subTest(invalid=invalid), self.assertRaisesRegex(HardwareUnavailableError, 'station baud rate'):
                DCMotorBackendAdapter(port='COM3', baud=9600, config=self.config)
        self.assertIsNone(self.store.read())

    def test_dropoff_targets_round_the_signed_vb6_count_once(self):
        self.assertEqual(sample_dropoff_position(-5000, 3000, use_xy_table=False), -1700)
        self.assertEqual(sample_dropoff_position(-5000, 5, use_xy_table=False), -4994)
        self.assertEqual(sample_dropoff_position(-4999, 5, use_xy_table=False), -4994)
        self.assertEqual(sample_dropoff_position(-5000, 5, use_xy_table=True), -4996)

    def test_chain_dialog_displays_calibrated_controls_and_confirms_reference(self):
        app = QtWidgets.QApplication.instance() or QtWidgets.QApplication([])
        backend = self.connect(xy=False)
        dialog = DCMotorDialog(backend=backend, port='COM3', baud=9600)
        self.addCleanup(self.close_dialog, app, dialog)
        dialog._connected = True
        dialog._refresh_connection_state()
        self.assertEqual(dialog.axis_combo.count(), 3)
        self.assertTrue(dialog.home_xy_btn.isHidden())
        self.assertTrue(dialog.corner_btn.isHidden())
        self.assertTrue(dialog.reference_chain_btn.isEnabled())
        self.assertFalse(dialog.goto_hole_btn.isEnabled())
        self.assertFalse(dialog.baud_combo.isEnabled())
        self.assertEqual((dialog.hole_spin.decimals(), dialog.hole_spin.maximum()), (0, 100.))

        dialog.hole_spin.setValue(99)
        with patch.object(dialog, '_start_command', side_effect=lambda _label, action: action()) as command:
            with patch.object(QtWidgets.QMessageBox, 'question', return_value=QtWidgets.QMessageBox.StandardButton.No):
                dialog.reference_chain_btn.click()
                command.assert_not_called()
            with patch.object(QtWidgets.QMessageBox, 'question', return_value=QtWidgets.QMessageBox.StandardButton.Yes):
                dialog.reference_chain_btn.click()
        dialog._refresh_connection_state()
        self.assertTrue(dialog.goto_hole_btn.isEnabled())
        self.assertEqual(backend.read_hole('Changer (X)'), 99.)

    def close_dialog(self, app, dialog):
        dialog._plot_timer.stop()
        dialog._telemetry_thread.stop_monitoring()
        self.assertTrue(dialog._telemetry_thread.wait(5000))
        dialog.close()
        dialog.deleteLater()
        QtCore.QCoreApplication.sendPostedEvents(None, QtCore.QEvent.DeferredDelete)
        app.processEvents()
