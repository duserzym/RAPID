"""Exercise queue borrowing against the actual native vacuum command path."""
from pathlib import Path
import tempfile
import unittest
from unittest.mock import patch

from rapid_main.config import VacuumConfig
from rapid_main.diagnostic_services import VacuumBackendAdapter
from rapidpy_common.hardware_safety import HardwareSafetyError, HardwareSafetyStore
from rapidpy_common.queue_safety import QueueSafetyStore, QueueWorkflowSession, QueueSafeStateRecord
from updown_control.app import VacuumController
from tests.test_vacuum_transport import _FakeVacuumSerial


class SerialInterface(_FakeVacuumSerial):
    def close(self):
        self.is_open = False


class QueueVacuumTests(unittest.TestCase):
    def setUp(self):
        directory = tempfile.TemporaryDirectory()
        self.addCleanup(directory.cleanup)
        self.path = Path(directory.name) / 'safety.json'
        self.store = QueueSafetyStore(self.path)
        self.cfg = VacuumConfig(port='COM11', baud=9600, warn_threshold=0, auto_pump=True)
        self.adapter = VacuumBackendAdapter(self.cfg)
        self.adapter._safety_store = HardwareSafetyStore(self.path)
        self.profile = {'helper': 'rapid_main_queue',
            'resources': {'vacuum': self.adapter._binding()},
            'stage_profiles': {'vacuum': self.adapter.queue_station_binding(), 'af': {'board': 1}}}
        self.session = QueueWorkflowSession.start(self.store, {'commands': ['Holder', 'AF50']}, self.profile)
        self.addCleanup(self.session.release)
        self.serial = SerialInterface({'10MFF': 'MOTOR ON', '10VFF': 'VALVE OPEN',
            '10V00': 'VALVE CLOSED', '10M00': 'MOTOR OFF'})
        self.controller = VacuumController()
        patches = [patch('updown_control.app.serial.Serial', return_value=self.serial),
            patch('rapid_main.diagnostic_services.VacuumController', return_value=self.controller),
            patch('updown_control.app.time.sleep', return_value=None)]
        for item in patches:
            item.start()
            self.addCleanup(item.stop)
        self.addCleanup(self.controller.disconnect)

    def connect(self):
        with self.session.claim():
            self.adapter.queue_connect(self.session)

    def enable(self):
        self.connect()
        with self.session.claim():
            return self.adapter.queue_set_pump(self.session, True)

    def release(self, **overrides):
        checks = dict(motors_stopped_verified=True, specimen_secured=True, field_outputs_off_verified=True)
        checks.update(overrides)
        with self.session.claim():
            return self.adapter.queue_set_pump(self.session, False, **checks)

    def commands(self):
        return [value.decode('ascii') for value in self.serial.writes if value != b'\r']

    def test_connect_borrows_lease_without_enabling_or_guessing_off(self):
        self.connect()
        self.assertTrue(self.adapter.is_connected())
        self.assertEqual(self.commands(), [])
        self.assertFalse(self.adapter.output_state_known)
        self.assertTrue(self.adapter.outputs_held)
        self.assertIn('unverified', self.adapter.status())
        self.assertEqual(self.store.pending()['family'], 'queue')
        with self.assertRaises(HardwareSafetyError):
            with self.store.operation_lease():
                pass

    def test_enable_persists_stage_before_exact_native_output_commands(self):
        original_write = self.serial.write
        observed = []
        def write(payload):
            if payload != b'\r':
                root = self.store.pending()
                observed.append(root['stage']['token'])
                self.assertEqual(root['stage']['family'], 'vacuum')
                self.assertEqual(root['stage']['status'], 'pending')
                with self.assertRaises(HardwareSafetyError):
                    with HardwareSafetyStore(self.path).operation_lease():
                        pass
            return original_write(payload)
        self.serial.write = write
        record = self.enable()
        self.assertTrue(observed)
        self.assertEqual(len(set(observed)), 1)
        self.assertEqual(self.commands(), ['E', '10MFF', 'O', '10VFF'])
        self.assertTrue(record.hold_acknowledged)
        self.assertFalse(record.safe_state_confirmed)
        self.assertTrue(self.adapter.is_pump_on())
        root = self.store.pending()
        self.assertEqual(root['stage']['status'], 'held')
        self.assertEqual(root['stage']['record']['command_evidence'],
            [{'command': '10MFF', 'reply': 'MOTOR ON'}, {'command': '10VFF', 'reply': 'VALVE OPEN'}])
        self.assertEqual(self.store.verify_history(root), 1)

    def test_release_requires_all_three_physical_preconditions_before_any_io(self):
        self.enable()
        before = self.commands()
        for check in ('motors_stopped_verified', 'specimen_secured', 'field_outputs_off_verified'):
            for bad in (False, 'true', 1):
                with self.assertRaises(HardwareSafetyError):
                    self.release(**{check: bad})
                self.assertEqual(self.commands(), before)
                self.assertTrue(self.adapter.outputs_held)
        self.assertEqual(self.store.pending()['stage']['status'], 'held')

    def test_release_proves_both_off_commands_but_leaves_queue_owned(self):
        self.enable()
        record = self.release()
        self.assertTrue(record.safe_state_confirmed)
        self.assertEqual(self.commands()[-4:], ['C', '10V00', 'D', '10M00'])
        self.assertFalse(self.adapter.is_pump_on())
        self.assertFalse(self.adapter.outputs_held)
        self.assertEqual(self.store.pending()['stage']['status'], 'verified')
        self.assertEqual(self.store.pending()['family'], 'queue')
        with self.assertRaisesRegex(RuntimeError, 'original queue'):
            self.adapter.disconnect()
        with self.session.claim():
            self.adapter.queue_disconnect(self.session)
        self.assertFalse(self.adapter.is_connected())
        self.assertIsNone(self.adapter._queue_binding)
        self.assertIsNotNone(self.store.pending())
        self.store.finish_queue(self.session.token, self.profile, QueueSafeStateRecord(True, True, True))
        self.assertIsNone(self.store.pending())

    def test_failed_release_attempts_motor_off_independently_and_retains_pending_stage(self):
        self.enable()
        del self.serial._responses['10V00']
        with self.assertRaisesRegex(HardwareSafetyError, 'unverified'):
            self.release()
        self.assertEqual(self.commands()[-4:], ['C', '10V00', 'D', '10M00'])
        pending = self.store.pending()
        token = pending['stage']['token']
        self.assertEqual(pending['stage']['status'], 'pending')
        self.assertFalse(pending['stage']['record']['safe_state_confirmed'])
        self.assertFalse(self.adapter.output_state_known)
        self.assertTrue(self.adapter.outputs_held)
        with self.session.claim():
            with self.assertRaises(HardwareSafetyError):
                self.adapter.queue_set_pump(self.session, True)
            with self.assertRaises(HardwareSafetyError):
                self.adapter.queue_disconnect(self.session)
        self.serial._responses['10V00'] = 'VALVE CLOSED'
        self.release()
        self.assertEqual(self.store.pending()['stage']['token'], token)
        self.assertEqual(self.store.pending()['stage']['status'], 'verified')

    def test_partial_enable_recovers_off_without_replaying_enable(self):
        self.connect()
        del self.serial._responses['10VFF']
        with self.session.claim():
            with self.assertRaises(HardwareSafetyError):
                self.adapter.queue_set_pump(self.session, True)
        token = self.store.pending()['stage']['token']
        self.assertTrue(self.controller._motor_powered)
        self.assertFalse(self.adapter.output_state_known)
        self.release()
        self.assertEqual(self.store.pending()['stage']['token'], token)
        self.assertEqual(self.commands().count('10MFF'), 1)
        self.assertEqual(self.commands()[-4:], ['C', '10V00', 'D', '10M00'])

    def test_publication_failure_after_off_preserves_unverified_owner_until_explicit_retry(self):
        self.enable()
        with patch.object(self.store, '_publish_event', side_effect=OSError('disk full')):
            with self.assertRaises(OSError):
                self.release()
        self.assertFalse(self.controller.is_enabled)
        self.assertFalse(self.adapter.output_state_known)
        self.assertTrue(self.adapter.outputs_held)
        self.assertEqual(self.store.pending()['stage']['status'], 'pending')
        self.release()
        self.assertFalse(self.adapter.outputs_held)

    def test_bad_owner_or_unclaimed_worker_cannot_connect_or_issue_outputs(self):
        with self.assertRaises(HardwareSafetyError):
            self.adapter.queue_connect(self.session)
        self.assertEqual(self.commands(), [])
        self.enable()
        before = self.commands()
        with self.assertRaises(HardwareSafetyError):
            self.adapter.queue_set_pump(self.session, True)
        self.assertEqual(self.commands(), before)

    def test_original_journal_and_actual_port_baud_are_required_before_outputs(self):
        self.connect()
        self.controller._port = 'COM99'
        with self.session.claim():
            with self.assertRaises(HardwareSafetyError):
                self.adapter.queue_set_pump(self.session, True)
        self.assertEqual(self.commands(), [])
        self.controller._port = 'COM11'
        self.adapter._cfg.baud = 19200
        with self.session.claim():
            with self.assertRaises(HardwareSafetyError):
                self.adapter.queue_set_pump(self.session, True)
        self.assertEqual(self.commands(), [])

    def test_pressure_gate_is_explicit_and_cannot_change_during_owned_lifetime(self):
        self.adapter._cfg.warn_threshold = 20
        with self.session.claim():
            with self.assertRaisesRegex(HardwareSafetyError, 'pressure telemetry'):
                self.adapter.queue_connect(self.session)
        self.assertFalse(self.adapter.is_connected())
        self.assertEqual(self.commands(), [])
        self.adapter._cfg.warn_threshold = 0
        self.enable()
        self.adapter._cfg.warn_threshold = float('nan')
        with self.session.claim():
            with self.assertRaises(HardwareSafetyError):
                self.adapter.queue_set_pump(self.session, True)
        self.assertEqual(self.commands().count('10MFF'), 1)

    def test_diagnostic_methods_cannot_replace_or_clear_queue_ownership(self):
        self.enable()
        before = self.commands()
        with self.assertRaises(RuntimeError):
            self.adapter.set_pump(False)
        with self.assertRaises(RuntimeError):
            self.adapter.connect()
        with self.assertRaises(RuntimeError):
            self.adapter.disconnect()
        self.assertEqual(self.commands(), before)
        self.assertEqual(self.store.pending()['family'], 'queue')

    def test_other_pending_stage_prevents_vacuum_release_or_reenable(self):
        self.enable()
        with self.session.claim() as child:
            child.begin('af', {'peak': 5}, {'board': 1})
            before = self.commands()
            with self.assertRaises(HardwareSafetyError):
                self.adapter.queue_set_pump(self.session, False, motors_stopped_verified=True,
                    specimen_secured=True, field_outputs_off_verified=True)
            with self.assertRaises(HardwareSafetyError):
                self.adapter.queue_set_pump(self.session, True)
            self.assertEqual(self.commands(), before)

    def test_same_station_restart_recovers_release_without_replaying_hold(self):
        self.enable()
        self.session.release()
        recovered = QueueWorkflowSession.recover(QueueSafetyStore(self.path), self.profile)
        self.addCleanup(recovered.release)
        with recovered.claim():
            self.adapter.queue_connect(recovered)
            self.assertFalse(self.adapter.output_state_known)
            before = self.commands()
            with self.assertRaisesRegex(HardwareSafetyError, 'never replay'):
                self.adapter.queue_set_pump(recovered, True)
            self.assertEqual(self.commands(), before)
            released = self.adapter.queue_set_pump(recovered, False, motors_stopped_verified=True,
                specimen_secured=True, field_outputs_off_verified=True)
            self.assertTrue(released.safe_state_confirmed)
            self.adapter.queue_disconnect(recovered)
        self.assertEqual(self.commands().count('10MFF'), 1)
        self.assertIsNotNone(self.store.pending())

    def test_new_adapter_restart_never_treats_cached_off_as_proof(self):
        self.enable()
        self.controller.disconnect()
        self.session.release()
        recovered = QueueWorkflowSession.recover(QueueSafetyStore(self.path), self.profile)
        self.addCleanup(recovered.release)
        fresh = VacuumBackendAdapter(self.cfg)
        fresh._safety_store = HardwareSafetyStore(self.path)
        self.serial.is_open = True
        with recovered.claim():
            fresh.queue_connect(recovered)
            self.assertFalse(self.controller.is_enabled)  # New connection cache only.
            self.assertFalse(fresh.output_state_known)
            self.assertTrue(fresh.outputs_held)
            with self.assertRaises(HardwareSafetyError):
                fresh.queue_disconnect(recovered)
            fresh.queue_set_pump(recovered, False, motors_stopped_verified=True,
                specimen_secured=True, field_outputs_off_verified=True)
            fresh.queue_disconnect(recovered)
        self.assertEqual(self.commands().count('10MFF'), 1)
        self.assertEqual(self.commands()[-4:], ['C', '10V00', 'D', '10M00'])

    def test_failed_transport_close_retains_original_handle_and_owner_for_retry(self):
        self.enable()
        self.release()
        with patch.object(self.serial, 'close', side_effect=OSError('port close failed')):
            with self.session.claim():
                with self.assertRaises(OSError):
                    self.adapter.queue_disconnect(self.session)
        self.assertIs(self.controller._serial, self.serial)
        self.assertIsNotNone(self.adapter._queue_binding)
        self.assertTrue(self.adapter.output_state_known)
        self.assertFalse(self.adapter.operation_active)
        with self.session.claim():
            self.adapter.queue_disconnect(self.session)
        self.assertIsNone(self.controller._serial)
        self.assertIsNone(self.adapter._queue_binding)

    def test_close_error_after_port_closed_can_retry_without_losing_original_handle(self):
        self.enable()
        self.release()
        def incomplete_close():
            self.serial.is_open = False
            raise OSError('close bookkeeping failed')
        with patch.object(self.serial, 'close', side_effect=incomplete_close):
            with self.session.claim():
                with self.assertRaises(OSError):
                    self.adapter.queue_disconnect(self.session)
        self.assertIs(self.controller._serial, self.serial)
        with self.session.claim():
            self.adapter.queue_disconnect(self.session)
        self.assertIsNone(self.controller._serial)

    def test_connection_cleanup_failure_retains_transport_handle_under_queue_owner(self):
        with (patch.object(self.serial, 'reset_input_buffer', side_effect=OSError('setup failed')),
              patch.object(self.serial, 'close', side_effect=OSError('close failed'))):
            with self.session.claim():
                with self.assertRaises(OSError):
                    self.adapter.queue_connect(self.session)
        self.assertIs(self.adapter._pump_only_controller, self.controller)
        self.assertIs(self.controller._serial, self.serial)
        self.assertIsNotNone(self.adapter._queue_binding)
        self.assertFalse(self.adapter.output_state_known)
        self.assertTrue(self.adapter.outputs_held)
        self.assertEqual(self.commands(), [])

    def test_lost_connection_requires_off_recovery_and_never_replays_hold(self):
        self.enable()
        self.serial.is_open = False
        def open_serial(*args, **kwargs):
            self.serial.is_open = True
            return self.serial
        with patch('updown_control.app.serial.Serial', side_effect=open_serial):
            self.connect()
        self.assertFalse(self.adapter.output_state_known)
        with self.session.claim():
            with self.assertRaisesRegex(HardwareSafetyError, 'never replay'):
                self.adapter.queue_set_pump(self.session, True)
        self.release()
        self.assertEqual(self.commands().count('10MFF'), 1)
