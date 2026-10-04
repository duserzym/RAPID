"""Native terminal cutoff/release/close settlement with original injected handles."""
import unittest
from unittest.mock import patch

from rapidpy_common.hardware_safety import HardwareSafetyError
from tests import test_queue_backend_coordinator as backend_fixture


class QueueTerminalTests(backend_fixture.QueueBackendCoordinatorFixture, unittest.TestCase):
    def finish(self):
        return self.command(self.backend.finish_queue_lifetime)

    def test_loaded_specimen_returns_before_terminal_cutoff_release_and_close(self):
        self.load_backend()
        terminal = self.backend._queue_terminal_cleanup
        controllers = tuple(self.motor._connections.values())
        vacuum = self.vacuum._pump_only_controller
        worker = self.finish()
        self.assertTrue(worker.ok, worker.error)
        self.assertIsNone(self.store.pending())
        state = self.store.read()
        self.assertTrue(state['record']['safe_state_confirmed'])
        self.assertEqual(state['record']['settlement']['terminal_record_id'], terminal.last_record.record_id)
        self.assertEqual(state['stage']['record']['schema'], 'rapidpy.queue_terminal_close.v1')
        self.assertTrue(all(controller._serial is None for controller in controllers))
        self.assertIsNone(vacuum._serial)
        self.assertFalse(vacuum._valve_connected)
        self.assertFalse(vacuum._motor_powered)
        self.assertIsNone(self.vacuum._queue_binding)
        self.assertIsNone(self.backend._queue_coordinator)
        self.assertIs(self.backend._safety_store, self.store)
        self.assertFalse(self.backend.has_unresolved_hardware_fault)
        self.assertIsNotNone(self.session._lease)
        self.assertIsNone(self.session._owner)
        self.session.release()
        self.assertIsNone(self.session._lease)

    def test_referenced_empty_station_can_finish_without_fabricated_specimen(self):
        worker = self.finish()
        self.assertTrue(worker.ok, worker.error)
        self.assertIsNone(self.store.pending())
        self.assertIsNone(self.store.read()['stage']['plan']['pose']['operation']['specimen_context'])

    def test_field_cutoff_failure_preserves_grip_and_all_original_transports(self):
        self.load_backend()
        self.cap_voltage = .5
        before = len(self.vacuum_serial.writes)
        worker = self.finish()
        self.assertFalse(worker.ok)
        self.assertIn('discharge', worker.error)
        self.assertTrue(self.vacuum.is_valve_connected())
        self.assertEqual(len(self.vacuum_serial.writes), before)
        self.assertEqual(len(self.motor._connections), 4)
        self.assertIsNone(self.backend._queue_terminal_cleanup._close_token)

    def test_failed_clearance_never_releases_pump_or_closes_transports(self):
        self.lift_serial().status = 0
        before = len(self.vacuum_serial.writes)
        worker = self.finish()
        self.assertFalse(worker.ok)
        self.assertIn('clearance', worker.error)
        self.assertEqual(len(self.vacuum_serial.writes), before)
        self.assertEqual(len(self.motor._connections), 4)
        self.assertEqual(self.store.pending()['stage']['status'], 'pending')

    def test_failed_vacuum_close_still_closes_motors_and_retry_never_replays_outputs(self):
        serial = self.vacuum_serial
        with patch.object(serial, 'close', side_effect=RuntimeError('vacuum close failed')):
            worker = self.finish()
        self.assertFalse(worker.ok)
        terminal = self.backend._queue_terminal_cleanup
        token = terminal._close_token
        self.assertFalse(self.motor._connections)
        self.assertIs(self.vacuum._pump_only_controller._serial, serial)
        native_commands, vacuum_commands = len(self.commands), len(serial.writes)
        worker = self.finish()
        self.assertTrue(worker.ok, worker.error)
        self.assertEqual((len(self.commands), len(serial.writes)), (native_commands, vacuum_commands))
        self.assertEqual(self.store.read()['stage']['token'], token)
        self.assertIsNone(self.store.pending())

    def test_failed_motor_close_retains_original_client_and_retries_only_that_handle(self):
        client = self.motor._connections['COM3']
        serial = client._serial
        with patch.object(serial, 'close', side_effect=RuntimeError('motor close failed')):
            worker = self.finish()
        self.assertFalse(worker.ok)
        self.assertEqual(self.motor._connections, {'COM3': client})
        self.assertIs(client._serial, serial)
        before = len(self.commands), len(self.vacuum_serial.writes), len(serial.writes)
        worker = self.finish()
        self.assertTrue(worker.ok, worker.error)
        self.assertEqual((len(self.commands), len(self.vacuum_serial.writes), len(serial.writes)), before)
        self.assertIsNone(client._serial)

    def test_terminal_stage_is_pending_before_any_handle_close(self):
        original = self.vacuum_serial.close
        observations = []
        def close():
            stage = self.store.pending()['stage']
            observations.append((stage['status'], stage['plan']['action']))
            original()
        with patch.object(self.vacuum_serial, 'close', close):
            worker = self.finish()
        self.assertTrue(worker.ok, worker.error)
        self.assertEqual(observations, [('pending', 'queue_terminal_close')])

    def test_root_publication_failure_keeps_owner_and_retries_without_hardware_commands(self):
        with patch.object(self.store, 'finish_queue', side_effect=OSError('root journal failed')):
            worker = self.finish()
        self.assertFalse(worker.ok)
        self.assertEqual(self.store.pending()['stage']['status'], 'verified')
        self.assertIsNotNone(self.backend._queue_coordinator)
        before = len(self.commands), len(self.vacuum_serial.writes)
        worker = self.finish()
        self.assertTrue(worker.ok, worker.error)
        self.assertEqual((len(self.commands), len(self.vacuum_serial.writes)), before)

    def test_successful_root_write_with_lost_acknowledgement_settles_original_verified_record(self):
        native = self.store.finish_queue
        def lost(*args):
            native(*args)
            raise OSError('acknowledgement lost')
        with patch.object(self.store, 'finish_queue', lost):
            worker = self.finish()
        self.assertFalse(worker.ok)
        self.assertEqual(self.store.read()['status'], 'verified')
        self.assertIsNotNone(self.backend._queue_coordinator)
        terminal = self.backend._queue_terminal_cleanup
        record = terminal.last_root_record
        with self.session.claim_verified_finish(record), self.assertRaises(HardwareSafetyError):
            self.session.child_store.begin('af', {'action': 'must_not_reenergize'}, self.backend._safety_profile())
        before = len(self.commands), len(self.vacuum_serial.writes)
        worker = self.finish()
        self.assertTrue(worker.ok, worker.error)
        self.assertEqual((len(self.commands), len(self.vacuum_serial.writes)), before)

    def test_failed_vacuum_off_never_closes_original_transports(self):
        controller = self.vacuum._pump_only_controller
        with patch.object(controller, 'set_enabled', side_effect=RuntimeError('OFF acknowledgement failed')):
            worker = self.finish()
        self.assertFalse(worker.ok)
        self.assertEqual(len(self.motor._connections), 4)
        self.assertIs(controller._serial, self.vacuum_serial)
        self.assertTrue(controller._motor_powered)
        self.assertEqual(self.store.pending()['stage']['family'], 'vacuum')
        self.assertIsNone(self.backend._queue_terminal_cleanup._close_token)

    def test_close_evidence_publication_failure_retries_only_original_settlement(self):
        native = self.store._publish_event
        def publish(state, stage):
            if stage['plan']['action'] == 'queue_terminal_close': raise OSError('close evidence failed')
            return native(state, stage)
        with patch.object(self.store, '_publish_event', publish):
            worker = self.finish()
        self.assertFalse(worker.ok)
        token = self.store.pending()['stage']['token']
        self.assertEqual(self.store.pending()['stage']['status'], 'pending')
        self.assertFalse(self.motor._connections)
        self.assertIsNone(self.vacuum._pump_only_controller._serial)
        before = len(self.commands), len(self.vacuum_serial.writes)
        worker = self.finish()
        self.assertTrue(worker.ok, worker.error)
        self.assertEqual((len(self.commands), len(self.vacuum_serial.writes)), before)
        self.assertEqual(self.store.read()['stage']['token'], token)

    def test_settlement_claim_rejects_changed_root_record(self):
        self.assertTrue(self.finish().ok)
        state = self.store.read()
        # The session still holds its OS lease until its actual worker exits.
        from rapid_main.queue_terminal import QueueTerminalSafeStateRecord
        wrong = QueueTerminalSafeStateRecord(True, True, True, record_id=state['record']['record_id'])
        with self.assertRaises(HardwareSafetyError), self.session.claim_verified_finish(wrong):
            self.fail('A different terminal record acquired the verified settlement claim.')

    def test_replaced_original_vacuum_handle_is_rejected_before_retry(self):
        with patch.object(self.vacuum_serial, 'close', side_effect=RuntimeError('close failed')):
            self.finish()
        original = self.vacuum._pump_only_controller
        self.vacuum._pump_only_controller = object()
        try:
            worker = self.finish()
            self.assertFalse(worker.ok)
            self.assertIn('original terminal handles', worker.error)
            self.assertIsNotNone(self.store.pending())
        finally:
            self.vacuum._pump_only_controller = original

    def test_recovery_and_unclaimed_callers_cannot_replay_terminal_commands(self):
        with self.assertRaises(HardwareSafetyError): self.backend.finish_queue_lifetime()
        self.session.is_recovery = True
        before = len(self.commands), len(self.vacuum_serial.writes)
        worker = self.finish()
        self.assertFalse(worker.ok)
        self.assertEqual((len(self.commands), len(self.vacuum_serial.writes)), before)
