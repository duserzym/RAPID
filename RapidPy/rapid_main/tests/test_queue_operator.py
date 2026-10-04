"""Native loading corner and durable per-file operator orientation."""
import unittest
from unittest.mock import patch

from rapidpy_common.hardware_safety import HardwareSafetyError
from tests import test_queue_backend_coordinator as backend_fixture


class QueueOperatorTests(backend_fixture.QueueBackendCoordinatorFixture, unittest.TestCase):
    def queue_plan(self):
        return {'commands': {'file_directions': {'F1': True, 'F2': False}}}

    def park(self):
        worker = self.command(self.backend.park_queue_station)
        self.assertTrue(worker.ok, worker.error)

    def confirm(self, kind='load', file_id=''):
        worker = self.command(lambda: self.backend.record_queue_tray_confirmation(kind, file_id, operator='Operator'))
        self.assertTrue(worker.ok, worker.error)

    def test_loading_corner_uses_negative_edges_without_zero_or_rod_flip(self):
        with patch.object(self.motor, 'zero_target_pos') as zero, patch.object(self.motor, 'turning_motor_rotate') as rotate:
            self.park()
        zero.assert_not_called()
        rotate.assert_not_called()
        record = self.store.pending()['stage']['record']
        self.assertEqual(record['schema'], 'rapidpy.queue_loading_corner.v1')
        self.assertEqual(record['operation']['travel_counts'], -3000000)
        self.assertTrue(record['safe_state_confirmed'])
        self.assertTrue(self.vacuum.is_pump_on())
        self.assertFalse(self.vacuum.is_valve_connected())

    def test_confirmed_load_and_file_flip_are_linked_and_independent(self):
        self.park()
        self.confirm()
        with self.session.claim():
            self.assertTrue(self.backend.queue_file_direction('F1'))
            self.assertFalse(self.backend.queue_file_direction('F2'))
        self.park()
        self.confirm('flip', 'F1')
        with self.session.claim():
            self.assertFalse(self.backend.queue_file_direction('F1'))
            self.assertFalse(self.backend.queue_file_direction('F2'))
        context = self.store.latest_operator_context(self.session.token)
        self.assertEqual(context['kind'], 'flip')
        self.assertEqual(context['file_id'], 'F1')
        self.assertEqual(context['operator'], 'Operator')
        self.store.verify_history(self.store.pending())

    def test_unconfirmed_orientation_cannot_authorize_specimen_load(self):
        with self.session.claim(), self.assertRaises(HardwareSafetyError): self.backend.queue_file_direction('F1')

    def test_duplicate_confirmation_never_repeats_field_or_motion_io(self):
        self.park()
        self.confirm()
        before = len(self.commands)
        worker = self.command(lambda: self.backend.record_queue_tray_confirmation('flip', 'F1', operator='Operator'))
        self.assertFalse(worker.ok)
        self.assertEqual(len(self.commands), before)

    def test_loading_again_cannot_reset_a_confirmed_tray(self):
        self.park()
        self.confirm()
        self.park()
        worker = self.command(lambda: self.backend.record_queue_tray_confirmation('load', operator='Operator'))
        self.assertFalse(worker.ok)
        self.assertIn('cannot reset', worker.error)

    def test_loaded_original_specimen_cannot_park_for_tray_handling(self):
        self.load_backend()
        before = len(self.commands)
        worker = self.command(self.backend.park_queue_station)
        self.assertFalse(worker.ok)
        self.assertEqual(len(self.commands), before)
        self.assertTrue(self.vacuum.is_valve_connected())

    def test_changed_switch_after_human_confirmation_keeps_pending_orientation(self):
        self.park()
        self.motor._connections['COM3']._serial.status = 48
        worker = self.command(lambda: self.backend.record_queue_tray_confirmation('load', operator='Operator'))
        self.assertFalse(worker.ok)
        self.assertIn('negative XY edges', worker.error)
        self.assertEqual(self.store.pending()['stage']['status'], 'pending')
        self.assertEqual(self.store.latest_operator_context(self.session.token)['phase'], 'unverified')

    def test_wrong_file_flip_fails_before_fresh_output_commands(self):
        self.park()
        self.confirm()
        self.park()
        before = len(self.commands)
        worker = self.command(lambda: self.backend.record_queue_tray_confirmation('flip', 'foreign', operator='Operator'))
        self.assertFalse(worker.ok)
        self.assertEqual(len(self.commands), before)

    def test_failed_corner_stop_cannot_publish_safe_park(self):
        with patch.object(self.table, '_stopped', return_value=({}, 'stop failure')):
            worker = self.command(self.backend.park_queue_station)
        self.assertFalse(worker.ok)
        self.assertEqual(self.store.pending()['stage']['status'], 'pending')
        self.assertFalse(self.store.pending()['stage']['record']['safe_state_confirmed'])
