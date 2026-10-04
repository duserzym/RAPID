"""Native relative XY commands, switch edges and queue reference provenance."""
from dataclasses import replace
import unittest
from unittest.mock import patch

from rapid_main.queue_xy_reference import QueueXYReference, QueueXYReferenceState
from rapidpy_common.hardware_safety import HardwareSafetyError
from tests.test_queue_lift_transfer import QueueLiftFixture


class QueueXYReferenceTests(QueueLiftFixture, unittest.TestCase):
    def setUp(self):
        super().setUp()
        self.reference = QueueXYReference(self.table, self.vacuum,
            monotonic=self.clock.monotonic, timeout_s=3)
        self.stalled = set()
        self.cancelled = False
        self.table.should_cancel = lambda: self.cancelled
        for key, negative, positive in (('changer_x', 4, 5), ('changer_y', 5, 6)):
            serial = self.serial(key)
            serial.status = (1 << negative) | (1 << positive)
            def observe(command, serial=serial, key=key, neg=negative, pos=positive):
                if command.startswith('135 '):
                    self.assertEqual(self.store.pending()['stage']['status'], 'pending')
                    self.assertEqual(self.store.pending()['stage']['plan']['action'], 'xy_home_reference')
                    distance = int(command.split()[1])
                    if key not in self.stalled:
                        serial.position = -1000 if distance < 0 else 1000
                        serial.status = (1 << pos) if distance < 0 else (1 << neg)
            serial.on_write = observe

    def serial(self, key):
        return self.motor._connections[self.axes[key].port]._serial

    def home(self, **kwargs):
        values = dict(empty_rod_confirmed=True, field_outputs_off_verified=True)
        values.update(kwargs)
        with self.session.claim(): return self.reference.home(self.session, **values)

    def commands(self, key):
        return [value.decode().strip().split(' ', 1)[1] for value in self.serial(key).writes]

    def test_two_edge_commands_are_staged_routed_and_verified_before_zero(self):
        state = self.home()
        self.assertEqual(state.home_counts, (0, 0))
        self.assertEqual(state.connection_id, self.motor._connection_id)
        self.assertEqual(self.store.pending()['stage']['status'], 'verified')
        for key, masks in (('changer_x', (-1, -2)), ('changer_y', (-2, -3))):
            commands = self.commands(key)
            relative = [command for command in commands if command.startswith('135 ')]
            self.assertEqual(len(relative), 2)
            self.assertEqual([int(command.split()[1]) for command in relative], [-30000000, 30000000])
            self.assertEqual([int(command.split()[-2]) for command in relative], list(masks))
            self.assertTrue(all(command.split()[2] == '483184' for command in relative))
            self.assertLess(commands.index(relative[-1]), commands.index('145'))
        self.assertEqual(len(self.reference.last_record.cleanup_observations), 4)
        self.assertFalse(self.vacuum.is_valve_connected())

    def test_reference_can_be_required_after_verified_move_using_same_connection(self):
        state = self.home()
        with self.session.claim():
            self.assertEqual(self.reference.require(self.session), state)
            move = self.table.move_to_slot(self.session, 46, reference_verified=state, field_outputs_off_verified=True)
            self.assertEqual(move.operation['reference_proof']['reference_id'], state.reference_id)
            self.assertEqual(self.reference.require(self.session), state)

    def test_stalled_switch_has_deadline_and_keeps_pending_without_zero(self):
        self.stalled.add('changer_y')
        with self.assertRaisesRegex(HardwareSafetyError, 'deadline'): self.home()
        self.assertEqual(self.store.pending()['stage']['status'], 'pending')
        self.assertFalse(any(command == '145' for command in self.commands('changer_x')))
        self.assertEqual(self.store.latest_reference_context(self.session.token)['phase'], 'unverified')
        self.assertEqual(len(self.reference.last_record.cleanup_observations), 4)

    def test_both_opposing_switches_active_cannot_prove_home(self):
        self.serial('changer_x').status = 0
        with self.assertRaisesRegex(HardwareSafetyError, 'contradictory'): self.home()
        self.assertFalse(any(command.startswith('135 ') for command in self.commands('changer_x')))

    def test_lost_lift_top_clearance_prevents_xy_sweep(self):
        self.lift_serial().status = 0
        with self.assertRaisesRegex(HardwareSafetyError, 'top-reference'): self.home()
        self.assertFalse(any(command.startswith('135 ') for command in self.commands('changer_x')))

    def test_failed_stop_ack_prevents_zero_even_when_velocity_is_zero(self):
        stop = self.motor.stop
        def failed(axis):
            stop(axis)
            if axis.name == 'changer_y': raise RuntimeError('no stop ACK')
        with patch.object(self.motor, 'stop', side_effect=failed):
            with self.assertRaisesRegex(HardwareSafetyError, 'no stop ACK'): self.home()
        self.assertEqual(self.store.pending()['stage']['status'], 'pending')
        self.assertFalse(any(command == '145' for command in self.commands('changer_y')))

    def test_cancel_after_first_xy_command_stops_all_axes_and_never_zeroes(self):
        original = self.serial('changer_x').on_write
        def cancel(command):
            original(command)
            if command.startswith('135 '): self.cancelled = True
        self.serial('changer_x').on_write = cancel
        with self.assertRaisesRegex(HardwareSafetyError, 'cancelled'): self.home()
        self.assertEqual(len(self.reference.last_record.cleanup_observations), 4)
        self.assertFalse(any(command == '145' for command in self.commands('changer_x')))
        self.assertFalse(any(command.startswith('135 ') for command in self.commands('changer_y')))

    def test_publication_failure_does_not_install_live_reference(self):
        with patch.object(self.store, '_publish_event', side_effect=OSError('disk')):
            with self.assertRaises(OSError): self.home()
        with self.session.claim(), self.assertRaises(HardwareSafetyError): self.reference.require(self.session)
        self.assertEqual(self.store.pending()['stage']['status'], 'pending')

    def test_reconnect_invalidates_old_reference_and_transfer_proof(self):
        state = self.home()
        self.motor.connect()
        self.assertNotEqual(state.connection_id, self.motor._connection_id)
        with self.session.claim():
            with self.assertRaises(HardwareSafetyError): self.reference.require(self.session)
            with self.assertRaises(HardwareSafetyError):
                self.table.move_to_slot(self.session, 46, reference_verified=state, field_outputs_off_verified=True)

    def test_original_panel_recovery_cannot_replay_reference_sweep(self):
        self.session.is_recovery = True
        with self.assertRaisesRegex(HardwareSafetyError, 'never replay'): self.home()
        self.assertFalse(any(command.startswith('135 ') for command in self.commands('changer_x')))

    def test_reference_requires_operator_empty_rod_and_field_off_proof_before_stage(self):
        count = self.store.pending()['stage_count']
        for values in (dict(empty_rod_confirmed=False), dict(field_outputs_off_verified=False)):
            with self.assertRaises(HardwareSafetyError): self.home(**values)
        self.assertEqual(self.store.pending()['stage_count'], count)

    def test_gripped_specimen_prevents_home_before_xy_commands(self):
        self.load()
        with self.assertRaises(HardwareSafetyError): self.home()
        self.assertFalse(any(command.startswith('135 ') for command in self.commands('changer_x')))
        self.assertTrue(self.vacuum.is_valve_connected())

    def test_relative_target_overflow_is_rejected_before_command(self):
        self.serial('changer_x').position = -(2**31) + 10
        with self.assertRaisesRegex(HardwareSafetyError, 'overflow'): self.home()
        self.assertFalse(any(command.startswith('135 ') for command in self.commands('changer_x')))

    def test_home_readback_outside_imported_anchor_keeps_pending(self):
        zero = self.motor.zero_target_pos
        def wrong(axis):
            zero(axis)
            if axis.name == 'changer_y': self.serial('changer_y').position = 500
        with patch.object(self.motor, 'zero_target_pos', side_effect=wrong):
            with self.assertRaisesRegex(HardwareSafetyError, 'home counts changed'): self.home()
        self.assertEqual(self.store.pending()['stage']['status'], 'pending')

    def test_changed_reference_id_cannot_authorize_slot_transfer(self):
        state = self.home()
        with self.session.claim(), self.assertRaises(HardwareSafetyError):
            self.table.move_to_slot(self.session, 46, reference_verified=replace(state, reference_id='f' * 32),
                                    field_outputs_off_verified=True)

    def test_pending_foreign_stage_invalidates_reference_until_settlement(self):
        self.home()
        with self.session.claim() as child:
            child.begin('acquisition', {}, self.profile['stage_profiles']['acquisition'])
            with self.assertRaises(HardwareSafetyError): self.reference.require(self.session)

    def test_nonzero_velocity_prevents_sweep_and_retains_pending(self):
        self.serial('changer_x').velocity = 1
        with self.assertRaisesRegex(HardwareSafetyError, 'motion'): self.home()
        self.assertEqual(self.store.pending()['stage']['status'], 'pending')
        self.assertFalse(any(command.startswith('135 ') for command in self.commands('changer_x')))

    def test_failed_disconnect_retains_original_handle_and_invalidates_reference(self):
        self.home()
        original = self.serial('changer_y')
        with patch.object(original, 'close', side_effect=OSError('close failed')):
            with self.assertRaisesRegex(Exception, 'close failed'): self.motor.disconnect()
        self.assertIs(self.motor._connections['COM4']._serial, original)
        self.assertFalse(self.motor.is_connected)
        self.assertIsNone(self.motor._connection_id)
        self.motor.disconnect()
        self.assertFalse(original.is_open)
        self.assertEqual(self.motor._connections, {})

    def test_failed_close_blocks_reconnect_without_replacing_original_handle(self):
        original = self.serial('changer_y')
        with patch.object(original, 'close', side_effect=OSError('close failed')), \
             patch('rapidpy_common.hardware.serial.Serial') as factory:
            with self.assertRaisesRegex(Exception, 'close failed'): self.motor.connect()
            factory.assert_not_called()
        self.assertIs(self.motor._connections['COM4']._serial, original)

    def test_connection_identity_change_during_sweep_prevents_remaining_commands(self):
        original = self.serial('changer_x').on_write
        def changed(command):
            original(command)
            if command.startswith('135 '): self.motor._connection_id = 'f' * 32
        self.serial('changer_x').on_write = changed
        with self.assertRaisesRegex(HardwareSafetyError, 'connection changed'): self.home()
        self.assertFalse(any(command.startswith('135 ') for command in self.commands('changer_y')))
        self.assertIsNone(self.reference._active_session)

    def test_calibration_change_during_sweep_prevents_remaining_commands(self):
        original = self.serial('changer_x').on_write
        def changed(command):
            original(command)
            if command.startswith('135 '): self.motor.config.changer_speed += 1
        self.serial('changer_x').on_write = changed
        with self.assertRaisesRegex(HardwareSafetyError, 'station calibration'): self.home()
        self.assertFalse(any(command.startswith('135 ') for command in self.commands('changer_y')))

    def test_empty_axis_set_never_reports_connected(self):
        from rapidpy_common.motor_routing import RoutedMotorSerialClient
        from rapidpy_common.hardware import MotorControllerConfig
        client = RoutedMotorSerialClient(MotorControllerConfig(), [])
        self.assertFalse(client.is_connected)
