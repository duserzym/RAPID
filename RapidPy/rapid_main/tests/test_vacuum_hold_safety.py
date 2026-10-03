from pathlib import Path
import os
import tempfile
import threading
from types import SimpleNamespace
import unittest
from unittest.mock import Mock, patch

from updown_control import app as lift
from rapidpy_common.hardware_safety import HardwareSafetyStore, HardwareSafetyError
from rapidpy_common.vacuum_diagnostic_safety import VacuumHoldSession
from rapidpy_common.hardware import MoveResult
from rapid_main.hardware_contracts import QueueHardwareBackend
from rapid_main.config import AppConfig
from test_vacuum_transport import _FakeVacuumSerial


class FakeVacuum:
    def __init__(self):
        self._port, self._baud = 'COM2', 9600
        self.is_connected = True
        self.is_enabled = False
        self.acknowledgements = []
        self.calls = []
        self.failure = None

    def set_enabled(self, enabled):
        self.calls.append(enabled)
        if self.failure:
            raise RuntimeError(self.failure)
        self.is_enabled = enabled
        self.acknowledgements.append({'command': 'enable' if enabled else 'disable', 'reply': 'ACK'})


class VacuumHoldTests(unittest.TestCase):
    def setUp(self):
        self.directory = tempfile.TemporaryDirectory()
        self.addCleanup(self.directory.cleanup)
        self.store = HardwareSafetyStore(Path(self.directory.name) / 'safety.json')
        self.vacuum = FakeVacuum()
        self.motor = Mock()
        self.motor.is_connected = True
        self.motor.read_registers.return_value = (1000, 0)
        self.motor.move_motor.return_value = MoveResult(1000, 1000, True)
        profile = lift._load_settings_profile(lift.DEFAULT_SETTINGS_PATH)
        with patch.object(lift, 'MotorSerialClient', return_value=self.motor):
            self.controller = lift.UpDownController(profile, safety_store=self.store)
        self.controller._port = 'COM5'
        self.session = VacuumHoldSession(self.vacuum, store=self.store)
        self.controller.held_station = self.session
        self.addCleanup(self.cleanup_session)

    def cleanup_session(self):
        self.motor.stop.side_effect = None
        self.motor.read_registers.side_effect = None
        self.motor.read_registers.return_value = (1000, 0)
        self.vacuum.failure = None
        try:
            self.session.close()
        finally:
            self.session._release()

    def test_vacuum_holds_lease_and_lift_moves_without_releasing_specimen(self):
        self.session.start(self.controller)
        self.assertEqual(self.vacuum.calls, [True])
        self.assertEqual(self.store.pending()['family'], 'station_diagnostic')
        with self.assertRaises(HardwareSafetyError):
            with self.store.operation_lease():
                self.fail('An external owner entered a held station.')
        self.controller.move_to_raw(1000, 100)
        self.assertEqual(self.vacuum.calls, [True])
        self.assertTrue(self.vacuum.is_enabled)
        self.assertIsNotNone(self.store.pending())
        self.assertFalse(self.store.read()['record']['safe_state_confirmed'])
        self.assertEqual(self.store.read()['record']['station_profile']['resources']['lift']['motor_port'], 'COM5')
        self.session.close()
        self.assertEqual(self.vacuum.calls, [True, False])
        self.assertIsNone(self.store.pending())
        self.assertIn('no pressure telemetry', self.store.read()['record']['vacuum_evidence_basis'])

    def test_worker_thread_can_use_lift_under_ui_owned_vacuum_lease(self):
        self.session.start(self.controller)
        failures = []
        def move():
            try:
                self.controller.move_to_raw(1000, 100)
            except Exception as exc:
                failures.append(exc)
        worker = threading.Thread(target=move)
        worker.start()
        worker.join(2)
        self.assertFalse(worker.is_alive())
        self.assertEqual(failures, [])
        self.assertTrue(self.session.active)
        self.assertEqual(self.vacuum.calls, [True])

    def test_vacuum_only_hold_can_join_lift_before_first_connection_output(self):
        self.motor.is_connected = False
        self.session.start()
        self.assertNotIn('lift', self.store.pending()['profile']['resources'])
        def connect(*args, **kwargs):
            self.assertEqual(self.store.pending()['profile']['resources']['lift']['motor_port'], 'COM5')
            self.motor.is_connected = True
        self.motor.connect.side_effect = connect
        self.controller.connect('COM5')
        self.controller.move_to_raw(1000, 100)
        self.assertEqual(self.vacuum.calls, [True])

    def test_existing_binding_cannot_change_or_disconnect_while_held(self):
        self.session.start(self.controller)
        for action in (lambda: self.controller.connect('COM6'), self.controller.disconnect,
                       lambda: self.controller.apply_settings_profile(self.controller.profile)):
            with self.assertRaises(HardwareSafetyError):
                action()
        self.motor.connect.assert_not_called()
        self.motor.disconnect.assert_not_called()

    def test_failed_motor_stop_keeps_vacuum_on_and_blocks_next_motion(self):
        self.session.start(self.controller)
        self.motor.read_registers.return_value = (1000, 1)
        with self.assertRaisesRegex(HardwareSafetyError, 'vacuum is retained'):
            self.session.close()
        self.assertEqual(self.vacuum.calls, [True])
        self.assertTrue(self.session.active)
        self.assertTrue(self.session.faulted)
        with self.assertRaises(HardwareSafetyError):
            self.controller.move_to_raw(1000, 100)
        self.motor.move_motor.assert_not_called()
        self.motor.read_registers.return_value = (1000, 0)
        self.session.close()
        self.assertEqual(self.vacuum.calls, [True, False])

    def test_failed_release_keeps_owner_and_retry_clears_only_after_ack(self):
        self.session.start(self.controller)
        self.vacuum.failure = 'valve acknowledgement missing'
        with self.assertRaises(HardwareSafetyError):
            self.session.close()
        self.assertTrue(self.session.active)
        self.assertIsNotNone(HardwareSafetyStore(self.store.path).pending())
        self.vacuum.failure = None
        self.session.close()
        self.assertFalse(self.session.active)
        self.assertIsNone(self.store.pending())

    def test_restart_requires_both_original_bindings_and_never_enables_vacuum(self):
        self.session.start(self.controller)
        self.session._release()  # Simulate OS releasing a crashed owner's lease.
        recovered = VacuumHoldSession(self.vacuum, store=self.store)
        self.vacuum._port = 'COM4'
        with self.assertRaises(HardwareSafetyError):
            recovered.recover(self.controller)
        self.vacuum._port = 'COM2'
        self.controller._port = 'COM6'
        with self.assertRaises(HardwareSafetyError):
            recovered.recover(self.controller)
        self.assertEqual(self.vacuum.calls, [True])
        self.controller._port = 'COM5'
        recovered.recover(self.controller)
        self.assertEqual(self.vacuum.calls, [True, False])
        self.motor.home_to_top.assert_not_called()
        self.assertIsNone(self.store.pending())

    def test_main_recovery_routes_held_vacuum_to_original_helper(self):
        self.session.start(self.controller)
        backend = object.__new__(QueueHardwareBackend)
        backend._config = AppConfig()
        backend._config.general.nocomm = False
        backend._safety_store = self.store
        backend._client = Mock()
        with self.assertRaises(HardwareSafetyError):
            backend.return_to_safe_state()  # Live owner prevents recovery first.
        self.session._release()
        with self.assertRaisesRegex(RuntimeError, 'Up/Down Control'):
            backend.return_to_safe_state()
        self.assertEqual(backend._client.mock_calls, [])

    def test_journal_failure_prevents_enable(self):
        with patch.object(self.store, 'begin', side_effect=HardwareSafetyError('journal unavailable')):
            with self.assertRaises(HardwareSafetyError):
                self.session.start(self.controller)
        self.assertEqual(self.vacuum.calls, [])
        self.assertFalse(self.session.active)

    def test_faulted_lift_can_reconnect_only_original_binding_before_recovery(self):
        self.session.start(self.controller)
        self.session.faulted = True
        self.motor.is_connected = False
        self.controller.connect('COM5')
        self.motor.connect.assert_called_once_with('COM5', baudrate=57600)
        with self.assertRaises(HardwareSafetyError):
            self.controller.connect('COM6')
        with self.assertRaises(HardwareSafetyError):
            self.controller.move_to_raw(1000, 100)
        self.motor.is_connected = True

    def test_native_main_vacuum_controls_cannot_touch_held_or_crashed_station(self):
        from rapid_main import diagnostic_services as services
        self.session.start(self.controller)
        backend = object.__new__(services.VacuumBackendAdapter)
        backend._pump_only_controller = Mock()
        backend._cfg = SimpleNamespace(port='COM2', baud=9600)
        backend._trace = Mock()
        with patch.dict(os.environ, RAPID_SAFETY_STATE=str(self.store.path)):
            with self.assertRaises(RuntimeError):
                backend.set_pump(False)
            self.session._release()
            with patch.object(services, 'VacuumController') as constructor:
                with self.assertRaises(RuntimeError):
                    backend._connect()
            constructor.assert_not_called()
        backend._pump_only_controller.set_enabled.assert_not_called()

    def test_resource_join_rejects_stale_tokens_other_families_and_rebinding(self):
        self.session.start(self.controller)
        binding = self.controller.safety_profile('COM6')
        with self.assertRaises(HardwareSafetyError):
            self.store.join_diagnostic_resource('0' * 32, 'lift', binding)
        with self.assertRaises(HardwareSafetyError):
            self.store.join_diagnostic_resource(self.session.token, 'lift', binding)
        with self.assertRaises(HardwareSafetyError):
            self.store.join_diagnostic_resource(self.session.token, 'unknown', {})
        self.assertEqual(self.store.pending()['profile']['resources']['lift']['motor_port'], 'COM5')

    def test_publication_failure_after_safe_release_retains_pending_without_live_hold(self):
        self.session.start(self.controller)
        with patch('rapidpy_common.vacuum_diagnostic_safety.publish_diagnostic_record', side_effect=HardwareSafetyError('evidence unavailable')):
            with self.assertRaises(HardwareSafetyError):
                self.session.close()
        self.assertFalse(self.session.active)
        self.assertFalse(self.vacuum.is_enabled)
        self.assertIsNotNone(self.store.pending())


class VacuumReleaseTransportTests(unittest.TestCase):
    def test_valve_failure_still_attempts_motor_off_and_retains_exact_ack(self):
        controller = lift.VacuumController()
        controller._serial = _FakeVacuumSerial({'10M00': 'MOTOR-OFF'})
        with patch.object(lift.time, 'sleep'):
            with self.assertRaises(lift.VacuumCommunicationError):
                controller.set_enabled(False)
        commands = [v for v in controller._serial.writes if v != b'\r']
        self.assertEqual(commands, [b'C', b'10V00', b'D', b'10M00'])
        self.assertEqual(controller.acknowledgements[-1], {'command': '10M00', 'reply': 'MOTOR-OFF'})


class VacuumUiSafetyTests(unittest.TestCase):
    def test_scan_blocks_vacuum_connection_changes_before_any_actuation(self):
        window = SimpleNamespace(_scan_worker=object(), _append=Mock())
        lift.MainWindow._connect_vacuum(window)
        lift.MainWindow._disconnect_vacuum(window)
        self.assertEqual(window._append.call_count, 2)

    def test_failed_release_keeps_window_and_connections_open(self):
        window = SimpleNamespace(_scan_worker=None, _station_outputs_held=lambda: True,
                                 _release_vacuum_station=Mock(side_effect=HardwareSafetyError('lift still moving')),
                                 _append=Mock(), controller=Mock(), vacuum=Mock())
        event = Mock()
        lift.MainWindow.closeEvent(window, event)
        event.ignore.assert_called_once()
        window.controller.disconnect.assert_not_called()
        window.vacuum.disconnect.assert_not_called()
