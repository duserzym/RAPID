"""Composed native XY/lift/field/vacuum calls with injected transports."""
import unittest
import json
from unittest.mock import patch

from rapid_main.queue_transfer_coordinator import QueueTransferCoordinator
from rapid_main.queue_xy_reference import QueueXYReference
from rapidpy_common.hardware_safety import HardwareSafetyError
from tests.test_queue_lift_transfer import QueueLiftFixture
from tests import test_queue_field_outputs as field_fixture


class QueueTransferCoordinatorTests(QueueLiftFixture, unittest.TestCase):
    additional_stage_profiles = field_fixture.QueueFieldOutputsTests.additional_stage_profiles

    def setUp(self):
        super().setUp()
        self.reference = QueueXYReference(self.table, self.vacuum, monotonic=self.clock.monotonic, timeout_s=3)
        for key, negative, positive in (('changer_x', 4, 5), ('changer_y', 5, 6)):
            serial = self.motor._connections[self.axes[key].port]._serial
            serial.status = (1 << negative) | (1 << positive)
            def observe(command, serial=serial, neg=negative, pos=positive):
                if command.startswith('135 '):
                    distance = int(command.split()[1])
                    serial.position = -1000 if distance < 0 else 1000
                    serial.status = (1 << pos) if distance < 0 else (1 << neg)
            serial.on_write = observe
        with self.session.claim():
            proof = self.fields.verify_off(self.session)
            self.reference.home(self.session, empty_rod_confirmed=True, field_outputs_off_verified=proof)
        self.coordinator = QueueTransferCoordinator(self.session, self.lift, self.reference, self.fields)

    def call_coordinator(self, method, *args, **kwargs):
        with self.session.claim(): return getattr(self.coordinator, method)(*args, **kwargs)

    def events(self):
        state = self.store._queue(self.session.token)
        self.store.verify_history(state)
        events = []
        link = state['history_head']
        while link:
            event = json.loads(self.store._event_path(self.session.token, link['id']).read_bytes())
            events.append(event['stage'])
            link = event['previous']
        return list(reversed(events))

    def test_composed_load_uses_measured_height_and_empty_table_before_handoff(self):
        state = self.call_coordinator('load', 1, 'S1', file_id='F1')
        self.assertEqual((state.sample_id, state.original_slot, state.file_id, state.sample_height, state.phase),
                         ('S1', 1, 'F1', 2800, 'lifted'))
        self.assertTrue(self.vacuum.is_valve_connected())
        x, y = (self.motor._connections[self.axes[key].port]._serial.position for key in ('changer_x', 'changer_y'))
        self.assertEqual((x, y), (0, 39))
        stages = self.events()
        grip = [stage for stage in stages if stage['family'] == 'vacuum' and stage['plan']['action'] == 'enable'][-1]
        pose = grip['plan']['transfer_pose']
        self.assertEqual(pose['transfer_context']['phase'], 'picked')
        self.assertEqual(pose['operation']['action'], 'verify_specimen_vacuum_pose')
        self.assertTrue(pose['safe_state_confirmed'])
        actions = [stage['plan']['action'] for stage in stages[-7:]]
        self.assertEqual(actions, ['verify_participating_field_outputs_off', 'xy_slot_move', 'specimen_pickup',
                                  'verify_specimen_vacuum_pose', 'enable', 'specimen_home_reference', 'xy_slot_move'])

    def test_composed_return_preserves_original_slot_until_supported_release_and_clearance(self):
        self.call_coordinator('load', 1, 'S1', file_id='F1')
        self.lift_serial().position, self.lift_serial().status = -1200, 0
        started = self.clock.time
        state = self.call_coordinator('return_specimen')
        self.assertEqual((state.phase, state.original_slot, state.sample_height), ('clear', 1, 2800))
        self.assertEqual(self.lift_serial().position, 0)
        self.assertFalse(self.vacuum.is_valve_connected())
        self.assertTrue(self.vacuum.is_pump_on())
        self.assertGreaterEqual(self.clock.time - started, 1.2)
        stages = self.events()
        release = [stage for stage in stages if stage['family'] == 'vacuum' and stage['plan']['action'] == 'pump_ready'][-1]
        self.assertEqual(release['plan']['transfer_pose']['transfer_context']['phase'], 'supported')
        self.assertIsNotNone(self.store.pending())

    def test_another_load_is_rejected_before_any_native_commands(self):
        self.call_coordinator('load', 1, 'S1')
        count = len(self.commands)
        with self.assertRaisesRegex(HardwareSafetyError, 'original specimen'):
            self.call_coordinator('load', 1, 'S2')
        self.assertEqual(len(self.commands), count)
        self.assertTrue(self.vacuum.is_valve_connected())

    def test_cutoff_failure_prevents_table_or_grip_change(self):
        self.cap_voltage = .5
        motor_writes = sum(len(client._serial.writes) for client in self.motor._connections.values())
        vacuum_writes = len(self.vacuum_serial.writes)
        with self.assertRaisesRegex(HardwareSafetyError, 'discharge'):
            self.call_coordinator('load', 1, 'S1')
        self.assertEqual(sum(len(client._serial.writes) for client in self.motor._connections.values()), motor_writes)
        self.assertEqual(len(self.vacuum_serial.writes), vacuum_writes)
        self.assertEqual(self.store.pending()['stage']['family'], 'field_outputs')

    def test_failed_home_preserves_grip_and_cannot_replay_load_or_return(self):
        with patch.object(self.motor, 'home_to_top', side_effect=RuntimeError('home failed')):
            with self.assertRaisesRegex(HardwareSafetyError, 'home failed'):
                self.call_coordinator('load', 1, 'S1')
        self.assertTrue(self.vacuum.is_valve_connected())
        self.assertEqual(self.context().phase, 'unverified')
        count = len(self.vacuum_serial.writes)
        for method, args in (('load', (1, 'S1')), ('return_specimen', ())):
            with self.assertRaises(HardwareSafetyError): self.call_coordinator(method, *args)
        self.assertEqual(len(self.vacuum_serial.writes), count)

    def test_failed_supported_pose_never_releases_grip(self):
        self.call_coordinator('load', 1, 'S1')
        native = self.lift.verify_vacuum_pose
        def disturbed(session, **kwargs):
            self.lift_serial().position = -4000
            return native(session, **kwargs)
        before = len(self.vacuum_serial.writes)
        with patch.object(self.lift, 'verify_vacuum_pose', side_effect=disturbed):
            with self.assertRaisesRegex(HardwareSafetyError, 'pose changed'):
                self.call_coordinator('return_specimen')
        self.assertEqual(len(self.vacuum_serial.writes), before)
        self.assertTrue(self.vacuum.is_valve_connected())
        self.assertEqual(self.context().phase, 'unverified')

    def test_pose_publication_failure_never_connects_grip(self):
        native = self.store._publish_event
        def publish(*args, **kwargs):
            stage = self.store._queue(self.session.token)['stage']
            if stage['plan'].get('action') == 'verify_specimen_vacuum_pose': raise OSError('disk')
            return native(*args, **kwargs)
        before = len(self.vacuum_serial.writes)
        with patch.object(self.store, '_publish_event', side_effect=publish):
            with self.assertRaises(OSError): self.call_coordinator('load', 1, 'S1')
        self.assertEqual(len(self.vacuum_serial.writes), before)
        self.assertFalse(self.vacuum.is_valve_connected())
        self.assertEqual(self.store.pending()['stage']['status'], 'pending')

    def test_intervening_stage_invalidates_a_pose_before_grip(self):
        with self.session.claim():
            proof = self.fields.verify_off(self.session)
            reference = self.reference.require(self.session)
            self.table.move_to_slot(self.session, 1, reference_verified=reference, field_outputs_off_verified=proof)
            self.lift.pickup(self.session, 1, 'S1', field_outputs_off_verified=proof)
            pose = self.lift.verify_vacuum_pose(self.session, field_outputs_off_verified=proof)
            self.fields.verify_off(self.session)
            before = len(self.vacuum_serial.writes)
            with self.assertRaisesRegex(HardwareSafetyError, 'latest original'):
                self.vacuum.queue_set_outputs(self.session, pump_enabled=True, valve_connected=True,
                                              transfer_pose=pose, field_outputs_off_verified=proof)
            self.assertEqual(len(self.vacuum_serial.writes), before)

    def test_reconnect_and_recovery_never_replay_coordinator_motion(self):
        self.motor.connect()
        with self.assertRaises(HardwareSafetyError): self.call_coordinator('load', 1, 'S1')
        self.session.is_recovery = True
        with self.assertRaisesRegex(HardwareSafetyError, 'never replay'):
            self.call_coordinator('load', 1, 'S1')

    def test_unclaimed_caller_cannot_enter_coordinator(self):
        with self.assertRaisesRegex(HardwareSafetyError, 'claim'):
            self.coordinator.load(1, 'S1')

    def test_failed_valve_release_does_not_raise_rod_or_drop_original_identity(self):
        self.call_coordinator('load', 1, 'S1', file_id='F1')
        controller = self.vacuum._pump_only_controller
        with patch.object(controller, 'set_valve_connect', side_effect=RuntimeError('valve ACK missing')):
            with self.assertRaisesRegex(HardwareSafetyError, 'valve ACK missing'):
                self.call_coordinator('return_specimen')
        self.assertEqual(self.lift_serial().position, -2480)
        self.assertTrue(controller._valve_connected)
        self.assertEqual(self.context().sample_id, 'S1')
        self.assertEqual(self.context().phase, 'unverified')
        self.assertEqual(self.store.pending()['stage']['family'], 'vacuum')

    def test_changed_pose_record_is_rejected_before_vacuum_commands(self):
        with self.session.claim():
            proof = self.fields.verify_off(self.session)
            reference = self.reference.require(self.session)
            self.table.move_to_slot(self.session, 1, reference_verified=reference, field_outputs_off_verified=proof)
            self.lift.pickup(self.session, 1, 'S1', field_outputs_off_verified=proof)
            pose = self.lift.verify_vacuum_pose(self.session, field_outputs_off_verified=proof)
            pose.record.transfer_context['sample_id'] = 'S2'
            before = len(self.vacuum_serial.writes)
            with self.assertRaisesRegex(HardwareSafetyError, 'latest original'):
                self.vacuum.queue_set_outputs(self.session, pump_enabled=True, valve_connected=True,
                                              transfer_pose=pose, field_outputs_off_verified=proof)
            self.assertEqual(len(self.vacuum_serial.writes), before)

    def test_changed_calibration_after_load_prevents_return_before_outputs(self):
        self.call_coordinator('load', 1, 'S1')
        self.motor.config.sample_bottom = -6000
        before = len(self.commands), len(self.vacuum_serial.writes)
        with self.assertRaisesRegex(HardwareSafetyError, 'calibration'):
            self.call_coordinator('return_specimen')
        self.assertEqual((len(self.commands), len(self.vacuum_serial.writes)), before)

    def test_new_specimen_can_load_only_after_verified_original_return(self):
        self.call_coordinator('load', 1, 'S1', file_id='F1')
        self.call_coordinator('return_specimen')
        state = self.call_coordinator('load', 1, 'S2', file_id='F2')
        self.assertEqual((state.sample_id, state.file_id, state.phase, state.sample_height),
                         ('S2', 'F2', 'lifted', 2800))

    def test_motor_reconnect_invalidates_recorded_pose_before_grip(self):
        with self.session.claim():
            proof = self.fields.verify_off(self.session)
            reference = self.reference.require(self.session)
            self.table.move_to_slot(self.session, 1, reference_verified=reference, field_outputs_off_verified=proof)
            self.lift.pickup(self.session, 1, 'S1', field_outputs_off_verified=proof)
            pose = self.lift.verify_vacuum_pose(self.session, field_outputs_off_verified=proof)
            self.motor.connect()
            before = len(self.vacuum_serial.writes)
            with self.assertRaisesRegex(HardwareSafetyError, 'latest original'):
                self.vacuum.queue_set_outputs(self.session, pump_enabled=True, valve_connected=True,
                                              transfer_pose=pose, field_outputs_off_verified=proof)
            self.assertEqual(len(self.vacuum_serial.writes), before)
