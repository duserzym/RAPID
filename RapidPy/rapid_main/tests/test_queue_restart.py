"""Only original verified shutdown and interrupted cutoff may be handled on restart."""
import copy
from types import SimpleNamespace
import unittest
from unittest.mock import patch

from rapid_main.queue_restart import QueueRestartRecovery, RestartShutdownAttestation, restart_stage
from rapidpy_common.hardware import serial
from rapidpy_common.hardware_safety import HardwareSafetyError
from tests.test_queue_lift_transfer import LiftSerial
from tests.test_queue_startup import QueueStartupFixture
from tests.test_queue_vacuum import SerialInterface


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

    def recover(self, state, *, shutdown=None):
        recovery = QueueRestartRecovery(self.backend, self.vacuum, state, operator='Recovery Operator', shutdown=shutdown)
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

    def confirm_shutdown(self, state, **overrides):
        proof = dict(specimen_secured=True, empty_rod=True, work_area_clear=True)
        proof.update(overrides)
        return RestartShutdownAttestation.confirm(state, 'Recovery Operator', **proof)

    def shutdown_transports(self, *, lift_top=True, valve_reply='VALVE CLOSED'):
        vacuum = SerialInterface({'10V00': valve_reply, '10M00': 'MOTOR OFF'})
        opened = []

        def open_serial(**settings):
            stream = vacuum if settings['port'] == 'COM11' else LiftSerial(**settings)
            if settings['port'] == 'COM5' and not lift_top:
                stream.status = 0
            opened.append((settings['port'], stream))
            return stream

        return vacuum, opened, patch.object(serial, 'Serial', side_effect=open_serial)

    def test_shutdown_aborts_original_acquisition_only_after_verified_clearance_and_off(self):
        original = self.interrupted_stage('acquisition', loaded=True)
        recovery = self.recover(original, shutdown=self.confirm_shutdown(original))
        vacuum, opened, transports = self.shutdown_transports()
        write = vacuum.write
        stages = []

        def observe(payload):
            if payload != b'\r':
                stage = self.store.pending()['stage']
                self.assertEqual(stage['token'], original['stage']['token'])
                self.assertTrue(stage['record']['clearance_verified'])
                stages.append(stage['record']['status'])
            return write(payload)

        vacuum.write = observe
        with transports:
            recovery.run()
        state = self.store.read()
        self.assertIsNone(self.store.pending())
        self.assertEqual(state['record']['settlement']['outcome'], 'aborted')
        self.assertEqual(state['stage']['status'], 'aborted')
        self.assertEqual(state['stage']['plan'], original['stage']['plan'])
        self.assertEqual(state['stage']['token'], original['stage']['token'])
        record = state['stage']['record']
        self.assertEqual(record['interrupted_history_head'], original['history_head'])
        self.assertTrue(record['vacuum_handle_closed'])
        self.assertFalse(record['safe_state_confirmed'])
        self.assertEqual(stages, ['release_ready'] * 4)
        self.assertEqual([cmd for cmd in vacuum.writes if cmd != b'\r'], [b'C', b'10V00', b'D', b'10M00'])
        self.assertEqual({port for port, _ in opened}, {'COM3', 'COM4', 'COM5', 'COM6', 'COM11'})
        self.assertTrue(all(not stream.is_open for _, stream in opened))
        self.store.verify_history(state)

    def test_confirmation_is_strict_and_bound_to_current_evidence(self):
        original = self.interrupted_stage('motion')
        with self.assertRaisesRegex(HardwareSafetyError, 'independent specimen support'):
            self.confirm_shutdown(original, specimen_secured=1)
        proof = self.confirm_shutdown(original)
        changed = copy.deepcopy(original)
        changed['run_id'] = 'new run'
        before = list(self.commands), list(self.vacuum_serial.writes)
        with self.assertRaisesRegex(HardwareSafetyError, 'for this queue'):
            self.recover(changed, shutdown=proof)
        self.assertEqual((self.commands, self.vacuum_serial.writes), before)
        self.assertEqual(self.store.pending(), original)

    def test_unverified_lift_clearance_prevents_even_opening_vacuum(self):
        original = self.interrupted_stage('motion')
        recovery = self.recover(original, shutdown=self.confirm_shutdown(original))
        vacuum, opened, transports = self.shutdown_transports(lift_top=False)
        with transports, self.assertRaisesRegex(HardwareSafetyError, 'top-reference switch'):
            recovery.run()
        self.assertNotIn('COM11', {port for port, _ in opened})
        self.assertEqual(vacuum.writes, [])
        self.assertEqual(self.store.pending()['stage']['record']['status'], 'failed')

    def test_negative_valve_reply_still_attempts_pump_off_and_leaves_stage_pending(self):
        original = self.interrupted_stage('motion')
        recovery = self.recover(original, shutdown=self.confirm_shutdown(original))
        vacuum, _, transports = self.shutdown_transports(valve_reply='ERROR')
        with transports, self.assertRaisesRegex(HardwareSafetyError, 'did not positively acknowledge'):
            recovery.run()
        self.assertEqual([cmd for cmd in vacuum.writes if cmd != b'\r'], [b'C', b'10V00', b'D', b'10M00'])
        self.assertFalse(vacuum.is_open)
        state = self.store.pending()
        self.assertEqual(state['stage']['token'], original['stage']['token'])
        self.assertFalse(state['stage']['record']['vacuum_release']['safe_state_confirmed'])

    def test_failed_abort_publication_keeps_original_stage_pending_after_off(self):
        original = self.interrupted_stage('acquisition', loaded=True)
        recovery = self.recover(original, shutdown=self.confirm_shutdown(original))
        vacuum, opened, transports = self.shutdown_transports()
        write = recovery.session.store._write

        def interrupt_commit(state):
            if state['status'] == 'verified':
                raise OSError('root write failed')
            write(state)

        with transports, patch.object(recovery.session.store, '_write', side_effect=interrupt_commit):
            with self.assertRaisesRegex(OSError, 'root write failed'):
                recovery.run()

        pending = self.store.pending()
        self.assertEqual(pending['stage']['token'], original['stage']['token'])
        self.assertEqual(pending['stage']['record']['status'], 'release_ready')
        self.assertEqual([cmd for cmd in vacuum.writes if cmd != b'\r'], [b'C', b'10V00', b'D', b'10M00'])
        self.assertTrue(all(not stream.is_open for _, stream in opened))
        self.store.verify_history(pending)

    def test_shutdown_can_settle_interruption_before_the_first_stage(self):
        startup = self.prepare()
        self.forget_session(startup)
        original = self.store.pending()
        self.assertIsNone(original['stage'])
        recovery = self.recover(original, shutdown=self.confirm_shutdown(original))
        _, _, transports = self.shutdown_transports()
        with transports:
            recovery.run()
        state = self.store.read()
        self.assertEqual(state['stage_count'], 1)
        self.assertEqual(state['stage']['plan']['action'], 'restart_shutdown')
        self.assertEqual(state['stage']['status'], 'aborted')
        self.store.verify_history(state)


class LiveQueueShutdownTests(QueueStartupFixture, unittest.TestCase):
    def interrupt_live_queue(self):
        startup = self.prepare()
        with self.arm_xy_edges():
            worker = self.command(self.backend.start_queue_lifetime)
        self.assertTrue(worker.ok, worker.error)
        self.load_backend()
        instruments = startup.instruments
        cfg = self.backend._config.squid
        handle = SimpleNamespace(port=cfg.port, baudrate=cfg.baud, parity='N', bytesize=8, stopbits=1,
                                 is_open=True)

        def close():
            self.assertEqual(self.store.pending()['stage']['record']['status'], 'release_ready')
            handle.is_open = False

        handle.close = close
        with startup.session.claim():
            child = startup.session.child_store
            token = child.begin('acquisition', {'action': 'squid'}, self.backend._acquisition_safety_profile())
            self.backend._geometry_stage_token = token
            instruments.raw._serial = handle
            instruments.validate()
            self.backend._geometry_stage_token = None
        self.lift_serial().position, self.lift_serial().status = 0, 16
        return startup, handle

    def shutdown_live(self, startup):
        state = self.store.pending()
        attestation = RestartShutdownAttestation.confirm(state, 'Operator', specimen_secured=True,
            empty_rod=True, work_area_clear=True)
        recovery = QueueRestartRecovery(self.backend, self.vacuum, state, operator='Operator',
            shutdown=attestation, live_startup=startup)
        return self.command(recovery.run)

    def test_live_abort_settles_original_handles_without_replaying_motion(self):
        startup, squid = self.interrupt_live_queue()
        original = self.store.pending()
        motors = [(client._serial, len(client._serial.writes)) for client in self.motor._connections.values()]
        vacuum_writes = len(self.vacuum_serial.writes)
        with patch.object(serial, 'Serial', side_effect=AssertionError('Original transports were still open')):
            worker = self.shutdown_live(startup)
        self.assertTrue(worker.ok, worker.error)
        current = self.store.read()
        self.assertEqual(current['token'], original['token'])
        self.assertEqual(current['stage']['token'], original['stage']['token'])
        self.assertEqual(current['stage']['plan'], original['stage']['plan'])
        self.assertEqual(current['stage']['record']['scope'], 'live')
        self.assertEqual(current['stage']['record']['instrument_closures']['closed'],
            dict(squid=True, susceptibility=True))
        self.assertEqual(current['record']['settlement']['outcome'], 'aborted')
        self.assertEqual([cmd for cmd in self.vacuum_serial.writes[vacuum_writes:] if cmd != b'\r'],
            [b'C', b'10V00', b'D', b'10M00'])
        self.assertFalse(squid.is_open or self.vacuum_serial.is_open)
        for stream, before in motors:
            self.assertFalse(stream.is_open)
            expected = [b'@16 3 0\r\n', b'@16 12 1 7\r\n', b'@16 12 1 7\r\n']
            if stream.port == 'COM5':
                expected.append(b'@16 20\r\n')
            self.assertEqual(stream.writes[before:], expected * 2)
        self.assertIsNone(self.backend._queue_startup)
        self.assertIsNone(self.backend._queue_specimen_geometry)
        self.assertIsNotNone(startup.session._lease)
        self.store.verify_history(current)

    def test_unverified_clearance_keeps_original_grip_connection(self):
        startup, squid = self.interrupt_live_queue()
        self.lift_serial().status = 0
        before = list(self.vacuum_serial.writes)
        worker = self.shutdown_live(startup)
        self.assertFalse(worker.ok)
        self.assertIn('top-reference switch', worker.error)
        self.assertEqual(self.vacuum_serial.writes, before)
        self.assertTrue(self.vacuum_serial.is_open and squid.is_open)
        self.assertIs(self.vacuum._queue_binding.session, startup.session)
        self.assertEqual(self.store.pending()['stage']['record']['status'], 'failed')

    def test_failed_science_close_keeps_original_handle_for_verified_retry(self):
        startup, squid = self.interrupt_live_queue()
        with patch.object(squid, 'close', side_effect=OSError('injected SQUID close')):
            worker = self.shutdown_live(startup)
        self.assertFalse(worker.ok)
        self.assertIn('injected SQUID close', worker.error)
        record = self.store.pending()['stage']['record']
        self.assertTrue(record['vacuum_handle_closed'])
        self.assertIs(startup.instruments.raw._serial, squid)
        self.assertFalse(record['instrument_closures']['closed']['squid'])
        self.assertIs(self.backend._safety_store, startup.session.child_store)

        self.vacuum_serial = SerialInterface({'10V00': 'CLOSED', '10M00': 'OFF'})
        worker = self.shutdown_live(startup)
        self.assertTrue(worker.ok, worker.error)
        self.assertFalse(squid.is_open)
        self.assertIsNone(self.store.pending())

    def test_changed_motor_handle_is_refused_before_journal_or_io(self):
        startup, _ = self.interrupt_live_queue()
        pending, commands = self.store.pending(), list(self.commands)
        vacuum_writes = list(self.vacuum_serial.writes)
        with patch.object(self.motor._connections['COM5'], '_serial', LiftSerial(port='COM5')):
            worker = self.shutdown_live(startup)
        self.assertFalse(worker.ok)
        self.assertIn('original live station', worker.error)
        self.assertEqual((self.store.pending(), self.commands, self.vacuum_serial.writes),
            (pending, commands, vacuum_writes))
