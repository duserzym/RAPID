"""Only original verified shutdown and interrupted cutoff may be handled on restart."""
import copy
import unittest
from unittest.mock import patch

from rapid_main.queue_restart import QueueRestartRecovery, restart_stage
from rapidpy_common.hardware_safety import HardwareSafetyError
from tests.test_queue_startup import QueueStartupFixture


class QueueRestartTests(QueueStartupFixture, unittest.TestCase):
    def forget_session(self, startup):
        """Keep durable evidence while discarding the original process bindings."""
        startup.session.release()
        self.backend._safety_store = self.store
        self.backend._queue_startup = None
        self.backend._queue_coordinator = None
        self.backend._queue_terminal_cleanup = None
        self.backend._queue_operator_stages = None
        self.backend._queue_instruments = None
        self.vacuum._queue_binding = None

    def ready_to_publish(self, *, close_failed_first=False):
        startup = self.prepare()
        with self.arm_xy_edges():
            worker = self.command(self.backend.start_queue_lifetime)
        self.assertTrue(worker.ok, worker.error)
        if close_failed_first:
            with patch.object(self.vacuum_serial, 'close', side_effect=OSError('retained close')):
                worker = self.command(self.backend.finish_queue_lifetime)
            self.assertFalse(worker.ok)
        with patch.object(startup.session.store, 'finish_queue', side_effect=OSError('root write failed')):
            worker = self.command(self.backend.finish_queue_lifetime)
        self.assertFalse(worker.ok)
        state = self.store.read()
        self.assertEqual(restart_stage(state), 'terminal_close')
        return startup, state

    def recover(self, state):
        recovery = QueueRestartRecovery(self.backend, self.vacuum, state, operator='Recovery Operator')
        self.addCleanup(recovery.release)
        return recovery

    def test_publishes_original_shutdown_even_with_prior_failed_close_event(self):
        startup, state = self.ready_to_publish(close_failed_first=True)
        with patch.object(self.backend, '_safety_store', self.store):
            with self.assertRaisesRegex(HardwareSafetyError, 'Another process'):
                self.recover(state)
        self.forget_session(startup)
        before = list(self.commands), list(self.vacuum_serial.writes)
        recovered = self.recover(state)
        recovered.run()
        committed = self.store.read()
        self.assertEqual(committed['status'], 'verified')
        self.assertEqual(committed['stage'], state['stage'])
        self.assertEqual(committed['history_head'], state['history_head'])
        attestation = committed['record']['settlement']['restart_publication']
        self.assertEqual(attestation['operator'], 'Recovery Operator')
        self.assertEqual(attestation['original_close_event'], state['history_head'])
        self.assertEqual((self.commands, self.vacuum_serial.writes), before)
        self.assertFalse(self.motor._connections)

    def test_field_restart_reuses_pending_stage_and_leaves_queue_latched(self):
        startup = self.prepare()
        self.cap_voltage = .5
        worker = self.command(self.backend.start_queue_lifetime)
        self.assertFalse(worker.ok)
        interrupted = self.store.read()
        self.assertEqual(restart_stage(interrupted), 'field_outputs')
        self.forget_session(startup)
        self.cap_voltage = 0
        before = len(self.commands)
        vacuum_before = list(self.vacuum_serial.writes)
        recovery = self.recover(interrupted)
        recovery.run()
        state = self.store.pending()
        stage = state['stage']
        self.assertEqual(stage['token'], interrupted['stage']['token'])
        self.assertEqual(state['stage_count'], interrupted['stage_count'])
        self.assertEqual(stage['status'], 'verified')
        self.assertEqual(stage['record']['operation']['recovery_operator'], 'Recovery Operator')
        cutoff_commands = self.commands[before:]
        self.assertTrue(stage['record']['af']['safe_state_confirmed'])
        self.assertEqual(len([item for item in stage['record']['af']['observations']
                              if item['action'] == 'process_status_after']), 10)
        self.assertFalse(any(command[0] == 'mcc_dac' and command[2] != 0 for command in cutoff_commands))
        self.assertEqual(self.vacuum_serial.writes, vacuum_before)
        self.assertFalse(self.motor._connections)
        with self.assertRaisesRegex(HardwareSafetyError, 'Station recovery is required'):
            restart_stage(state)

    def test_different_station_and_changed_original_state_are_refused_without_io(self):
        startup = self.prepare()
        self.cap_voltage = .5
        self.assertFalse(self.command(self.backend.start_queue_lifetime).ok)
        state = self.store.read()
        self.forget_session(startup)
        recovery = self.recover(state)
        before = list(self.commands), list(self.vacuum_serial.writes)
        with patch.object(self.backend._config.susceptibility, 'scale_factor', 83.):
            with self.assertRaisesRegex(HardwareSafetyError, 'exact original native station'):
                recovery.run()
        changed = copy.deepcopy(state)
        changed['run_id'] = 'a different run'
        self.store._write(changed)
        with self.assertRaisesRegex(HardwareSafetyError, 'changed after recovery was selected'):
            recovery.run()
        self.assertEqual((self.commands, self.vacuum_serial.writes), before)
        self.assertEqual(self.store.pending(), changed)

    def test_pending_terminal_close_is_not_a_verified_shutdown(self):
        startup = self.prepare()
        with self.arm_xy_edges():
            self.assertTrue(self.command(self.backend.start_queue_lifetime).ok)
        with patch.object(self.vacuum_serial, 'close', side_effect=OSError('retained close')):
            self.assertFalse(self.command(self.backend.finish_queue_lifetime).ok)
        state = self.store.pending()
        self.assertEqual(state['stage']['plan']['action'], 'queue_terminal_close')
        with self.assertRaisesRegex(HardwareSafetyError, 'terminal_close.*pending'):
            restart_stage(state)

    def test_missing_original_vacuum_event_blocks_publication(self):
        startup, state = self.ready_to_publish()
        self.forget_session(startup)
        recovery = self.recover(state)
        head = state['stage']['plan']['vacuum_event']
        self.store._event_path(state['token'], head['id']).unlink()
        with self.assertRaisesRegex(HardwareSafetyError, 'evidence cannot be verified'):
            recovery.run()
        self.assertEqual(self.store.pending(), state)
