from pathlib import Path
import tempfile
import unittest
from unittest.mock import patch

from rapid_main.queue_station import QueueStationGeometry
from rapid_main.queue_table_motion import QueueXYTableMotion
from rapidpy_common.hardware import MotorAxisConfig, MotorControllerConfig
from rapidpy_common.hardware_safety import HardwareSafetyError
from rapidpy_common.motor_routing import RoutedMotorSerialClient
from rapidpy_common.queue_safety import QueueSafetyStore, QueueWorkflowSession
from tests.test_motor_routing import FakeSerial


class PositionSerial(FakeSerial):
    def __init__(self, **kwargs):
        super().__init__(**kwargs)
        self.position = 0
        self.velocity = 0
        self.status = 16
        self.command = ''
        self.on_write = None
        self.fail_move = False
        self.slop = 0

    def write(self, payload):
        super().write(payload)
        self.command = payload.decode().strip().split(' ', 1)[1]
        if self.on_write:
            self.on_write(self.command)
        if self.command.startswith('134 '):
            if self.fail_move:
                raise OSError('axis command failed')
            self.position = int(self.command.split()[1]) + self.slop
        elif self.command == '145':
            self.position = 0
        elif self.command.startswith('11 10 '):
            self.relabel_target = -int(self.command.split()[-1])
        elif self.command == '165 1802':
            self.position = self.relabel_target

    def read_until(self, terminator):
        if self.command.startswith('12 '):
            values = [self.position if register == '1' else self.velocity
                      for register in self.command.split()[1:]]
            words = ' '.join(f'{value & 0xffffffff:08X}'[:4] + ' ' + f'{value & 0xffffffff:08X}'[4:]
                             for value in values)
            return f'@16 ACK {words}\r'.encode()
        if self.command == '20':
            return f'@16 ACK {self.status:04X}\r'.encode()
        return super().read_until(terminator)


class QueueTableMotionTests(unittest.TestCase):
    def setUp(self):
        directory = tempfile.TemporaryDirectory()
        self.addCleanup(directory.cleanup)
        self.store = QueueSafetyStore(Path(directory.name) / 'safety.json')
        self.axes = {key: MotorAxisConfig(key, ident, 16, port) for key, ident, port in
                     (('changer_x', 1, 'COM3'), ('turning', 2, 'COM6'),
                      ('updown', 3, 'COM5'), ('changer_y', 4, 'COM4'))}
        self.motor = RoutedMotorSerialClient(MotorControllerConfig(slot_max=100, one_step=-1010.1010101), list(self.axes.values()))
        for item in (patch('rapidpy_common.hardware.serial.Serial', side_effect=PositionSerial),
                     patch('rapidpy_common.hardware.time.sleep', return_value=None)):
            item.start()
            self.addCleanup(item.stop)
        self.motor.connect()
        self.addCleanup(self.motor.disconnect)
        self.geometry = QueueStationGeometry('station.ini', True, 1, 100, 46, -1010.1010101,
                                             ((46, 0, 39), (1, 9590, -11916)), (-3, -2))
        self.cancelled = False
        self.motion = QueueXYTableMotion(self.motor, self.axes, self.geometry,
            sleep=lambda seconds: None, should_cancel=lambda: self.cancelled)
        self.profile = dict(helper='rapid_main_queue', stage_profiles={'motion': self.motion.profile})
        self.session = QueueWorkflowSession.start(self.store, {'samples': ['sample']}, self.profile)
        self.addCleanup(self.session.release)

    def serial(self, key):
        return self.motor._connections[self.axes[key].port]._serial

    def moves(self, key):
        return [payload for payload in self.serial(key).writes if b'134 ' in payload]

    def move(self, slot=1, **overrides):
        checks = dict(reference_verified=True, field_outputs_off_verified=True)
        checks.update(overrides)
        with self.session.claim():
            return self.motion.move_to_slot(self.session, slot, sample_id='sample', **checks)

    def test_actual_native_xy_commands_use_each_registered_port_and_persist_first(self):
        def observe(command):
            self.assertEqual(self.store.pending()['stage']['status'], 'pending')
            self.assertEqual(self.store.pending()['stage']['family'], 'motion')
        for key in self.axes:
            self.serial(key).on_write = observe
        record = self.move()
        self.assertTrue(record.safe_state_confirmed)
        self.assertIn(b'134 9590 483184 8000000 0 0', self.moves('changer_x')[0])
        self.assertIn(b'134 -11916 483184 8000000 0 0', self.moves('changer_y')[0])
        self.assertEqual(self.moves('updown'), [])
        self.assertEqual(self.moves('turning'), [])
        root = self.store.pending()
        self.assertEqual(root['stage']['status'], 'verified')
        self.assertEqual(root['stage']['sample_id'], 'sample')
        self.assertEqual(self.store.verify_history(root), 1)
        self.assertIsNotNone(root)

    def test_empty_xy_location_uses_pair_not_chain_counts(self):
        self.move(46)
        self.assertEqual(self.geometry.verify_empty_xy_readback(46,
            self.serial('changer_x').position, self.serial('changer_y').position), 46)
        self.assertIn(b'134 0 ', self.moves('changer_x')[0])
        self.assertIn(b'134 39 ', self.moves('changer_y')[0])

    def test_unclaimed_and_changed_calibration_cannot_issue_native_commands(self):
        before = {key: len(self.serial(key).writes) for key in self.axes}
        with self.assertRaises(HardwareSafetyError):
            self.motion.move_to_slot(self.session, 1, reference_verified=True, field_outputs_off_verified=True)
        self.motor.config.changer_speed += 1
        with self.assertRaisesRegex(HardwareSafetyError, 'calibration'):
            self.move()
        self.assertEqual({key: len(self.serial(key).writes) for key in self.axes}, before)

    def test_reference_and_fields_require_strict_checks_before_any_io(self):
        before = {key: len(self.serial(key).writes) for key in self.axes}
        for key in ('reference_verified', 'field_outputs_off_verified'):
            for value in (False, 1, 'true'):
                with self.assertRaises(HardwareSafetyError):
                    self.move(**{key: value})
        with self.assertRaises(ValueError):
            self.move(2)
        self.assertEqual({key: len(self.serial(key).writes) for key in self.axes}, before)
        self.assertIsNone(self.store.pending()['stage'])

    def test_lift_clearance_and_switch_are_verified_without_homing(self):
        self.serial('updown').status = 0
        with self.assertRaisesRegex(HardwareSafetyError, 'top-reference'):
            self.move()
        self.assertEqual(self.moves('changer_x'), [])
        self.assertEqual(self.moves('changer_y'), [])
        self.assertFalse(self.motion.last_record.safe_state_confirmed)
        self.assertEqual(self.store.pending()['stage']['status'], 'pending')

    def test_y_command_failure_independently_stops_x_and_all_other_axes(self):
        self.serial('changer_y').fail_move = True
        with self.assertRaises(HardwareSafetyError):
            self.move()
        self.assertTrue(self.moves('changer_x'))
        self.assertEqual(self.store.pending()['stage']['status'], 'pending')
        for key in self.axes:
            self.assertGreaterEqual(sum(b'3 0' in value for value in self.serial(key).writes), 2)

    def test_stationary_slop_and_nonzero_velocity_cannot_verify_a_move(self):
        self.serial('changer_y').slop = 50
        with self.assertRaises(HardwareSafetyError):
            self.move()
        self.assertFalse(self.motion.last_record.safe_state_confirmed)
        self.assertEqual(self.store.pending()['stage']['status'], 'pending')

    def test_running_motor_readback_blocks_motion_even_with_stationary_position(self):
        self.serial('turning').velocity = 1
        with self.assertRaises(HardwareSafetyError):
            self.move()
        self.assertEqual(self.moves('changer_x'), [])
        self.assertIn('turning', self.motion.last_record.cleanup_error)

    def test_cancellation_after_x_command_stops_without_starting_y_or_claiming_success(self):
        def cancel(command):
            if command.startswith('134 '):
                self.cancelled = True
        self.serial('changer_x').on_write = cancel
        with self.assertRaises(HardwareSafetyError):
            self.move()
        self.assertTrue(self.moves('changer_x'))
        self.assertEqual(self.moves('changer_y'), [])
        self.assertFalse(self.motion.last_record.safe_state_confirmed)

    def test_recovery_never_replays_a_completed_xy_transfer(self):
        self.move()
        self.session.release()
        self.session = QueueWorkflowSession.recover(self.store, self.profile)
        self.addCleanup(self.session.release)
        before = len(self.serial('changer_x').writes)
        with self.assertRaisesRegex(HardwareSafetyError, 'never replay'):
            self.move()
        self.assertEqual(len(self.serial('changer_x').writes), before)

    def test_evidence_publication_failure_preserves_pending_original_stage(self):
        with patch.object(self.session.child_store, 'finish', side_effect=OSError('disk full')):
            with self.assertRaises(OSError):
                self.move()
        self.assertTrue(self.motion.last_record.safe_state_confirmed)
        self.assertEqual(self.store.pending()['stage']['status'], 'pending')

    def test_lift_position_outside_clearance_blocks_both_table_axes(self):
        self.serial('updown').position = self.motor.config.sample_bottom
        with self.assertRaisesRegex(HardwareSafetyError, 'clearance'):
            self.move()
        self.assertEqual(self.moves('changer_x'), [])
        self.assertEqual(self.moves('changer_y'), [])

    def test_failed_stop_acknowledgement_remains_unsafe_despite_zero_velocity(self):
        def fail_stop(command):
            if command.startswith('3 '):
                raise OSError('missing stop reply')
        self.serial('changer_x').on_write = fail_stop
        with self.assertRaises(HardwareSafetyError):
            self.move()
        self.assertIn('Stop acknowledgement', self.motion.last_record.cleanup_error)
        self.assertFalse(self.motion.last_record.safe_state_confirmed)
        self.assertEqual(self.store.pending()['stage']['status'], 'pending')
        for key in ('changer_y', 'updown', 'turning'):
            self.assertTrue(any(b'3 0' in payload for payload in self.serial(key).writes))
