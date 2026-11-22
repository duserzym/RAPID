"""Only original verified shutdown and interrupted cutoff may be handled on restart."""
import copy
import unittest
from unittest.mock import patch

from rapid_main.queue_restart import QueueRestartRecovery, restart_stage
from rapidpy_common.hardware import serial
from rapidpy_common.hardware_safety import HardwareSafetyError
from tests.test_queue_lift_transfer import LiftSerial
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
        self.backend._queue_specimen_geometry = None
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

    def interrupted_stage(self, family, *, loaded=False):
        startup = self.prepare()
        with self.arm_xy_edges():
            worker = self.command(self.backend.start_queue_lifetime)
        self.assertTrue(worker.ok, worker.error)
        if loaded:
            self.load_backend()
            self.assertTrue(self.vacuum.is_valve_connected())
        store = startup.session.store
        with startup.session.claim() as child:
            root = store.pending()
            plan = {'action': 'squid' if family == 'acquisition' else 'xy_slot_move', 'sample_id': 'S1'}
            child.begin(family, plan, root['profile']['stage_profiles'][family], sample_id='S1')
        state = store.pending()
        # Simulate process exit closing serial handles. This does not change
        # controller outputs or the original recorded specimen support.
        self.motor.disconnect()
        self.vacuum._pump_only_controller.disconnect()
        self.forget_session(startup)
        return state

    def test_interrupted_acquisition_preserves_grip_identity_and_issues_only_stop_and_read(self):
        original = self.interrupted_stage('acquisition', loaded=True)
        vacuum_before = list(self.vacuum_serial.writes)
        opened = []

        def transport(**settings):
            pending = self.store.pending()
            self.assertEqual(pending['stage']['record']['status'], 'started')
            self.assertEqual(pending['stage']['token'], original['stage']['token'])
            port = LiftSerial(**settings)
            port.position = 2880
            opened.append(port)
            return port

        recovery = self.recover(original)
        with patch.object(serial, 'Serial', side_effect=transport):
            recovery.run()
        state = self.store.pending()
        record = state['stage']['record']
        self.assertEqual(state['stage']['plan'], original['stage']['plan'])
        self.assertEqual(state['stage_count'], original['stage_count'])
        self.assertEqual(state['stage']['status'], 'pending')
        self.assertIs(record['safe_state_confirmed'], False)
        self.assertEqual(record['status'], 'contained')
        self.assertTrue(record['fields_off_verified'])
        self.assertTrue(record['motors_stopped_verified'])
        self.assertTrue(record['motor_handles_closed'])
        self.assertEqual(self.vacuum_serial.writes, vacuum_before)
        context = self.store.latest_transfer_context(state['token'])
        self.assertEqual(context['sample_id'], 'S1')
        self.assertEqual(context['phase'], 'unverified')
        self.assertEqual(len(opened), 4)
        for stream in opened:
            self.assertFalse(stream.is_open)
            self.assertEqual(stream.writes, [b'@16 3 0\r\n', b'@16 12 1 7\r\n', b'@16 12 1 7\r\n'])

    def test_partial_port_failure_still_stops_reachable_axes_and_keeps_pending(self):
        original = self.interrupted_stage('motion')
        self.cap_voltage = .5
        opened = {}

        def transport(**settings):
            if settings['port'] == 'COM4':
                raise OSError('controller absent')
            port = LiftSerial(**settings)
            opened[port.port] = port
            return port

        recovery = self.recover(original)
        with patch.object(serial, 'Serial', side_effect=transport):
            with self.assertRaisesRegex(HardwareSafetyError, 'COM4.*controller absent'):
                recovery.run()
        pending = self.store.pending()
        self.assertEqual(pending['stage']['token'], original['stage']['token'])
        record = pending['stage']['record']
        self.assertEqual(record['status'], 'failed')
        self.assertFalse(record['fields_off_verified'])
        self.assertFalse(record['motors_stopped_verified'])
        self.assertTrue(record['motor_handles_closed'])
        self.assertEqual(record['motor_observations']['changer_y'], [])
        self.assertEqual(set(opened), {'COM3', 'COM5', 'COM6'})
        self.assertTrue(all(port.writes[0] == b'@16 3 0\r\n' and not port.is_open for port in opened.values()))

    def test_failed_original_intent_write_blocks_field_and_serial_io(self):
        original = self.interrupted_stage('motion')
        recovery = self.recover(original)
        before = list(self.commands), list(self.vacuum_serial.writes)
        with patch.object(recovery.session.store, 'finish_stage', side_effect=OSError('disk full')), patch.object(
                serial, 'Serial') as transport, self.assertRaisesRegex(OSError, 'disk full'):
            recovery.run()
        transport.assert_not_called()
        self.assertEqual(self.store.pending(), original)
        self.assertEqual((self.commands, self.vacuum_serial.writes), before)

    def test_failed_close_retains_the_original_port_for_the_same_recovery(self):
        original = self.interrupted_stage('motion')
        first_streams = {}

        def transport(**settings):
            port = LiftSerial(**settings)
            first_streams[port.port] = port
            if port.port == 'COM3':
                def fail_close():
                    raise OSError('close failed')
                port.close = fail_close
            return port

        initial = self.recover(original)
        with patch.object(serial, 'Serial', side_effect=transport), self.assertRaisesRegex(HardwareSafetyError, 'close failed'):
            initial.run()
        pending = self.store.pending()
        self.assertTrue(pending['stage']['record']['motors_stopped_verified'])
        self.assertFalse(pending['stage']['record']['motor_handles_closed'])
        retained = self.motor._connections['COM3']._serial
        self.assertIs(retained, first_streams['COM3'])
        retained.close = LiftSerial.close.__get__(retained)
        initial.release()
        opened = []

        def retry_open(**settings):
            opened.append(settings['port'])
            return LiftSerial(**settings)

        retry = self.recover(pending)
        with patch.object(serial, 'Serial', side_effect=retry_open):
            retry.run()
        self.assertNotIn('COM3', opened)
        self.assertEqual(retained.writes.count(b'@16 3 0\r\n'), 2)
        self.assertFalse(retained.is_open)
        self.assertTrue(self.store.pending()['stage']['record']['motor_handles_closed'])
