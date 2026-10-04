from pathlib import Path
import tempfile
import threading
import time
from types import SimpleNamespace
import unittest
from unittest.mock import Mock, patch

from PySide6 import QtWidgets
from rapid_main.config import AppConfig, VacuumConfig
from rapid_main.diagnostic_services import VacuumBackendAdapter, HardwareUnavailableError
from rapid_main.dialogs.vacuum import VacuumDialog
from rapidpy_common.hardware_safety import HardwareSafetyStore, HardwareSafetyError
from rapidpy_common.vacuum_diagnostic_safety import VacuumHoldSession


class VacuumInterface:
    def __init__(self):
        self.is_connected = self.is_enabled = False
        self._port, self._baud = '', 9600
        self.acknowledgements = []
        self.calls = []
        self.failure = False
        self.before_output = None

    def connect(self, port, baudrate):
        self._port, self._baud = port, baudrate
        self.calls.append(('connect', port, baudrate))
        self.is_connected = True

    def disconnect(self):
        self.calls.append(('disconnect',))
        self.is_connected = False

    def set_enabled(self, enabled):
        self.calls.append(('output', enabled))
        if self.before_output:
            self.before_output(enabled)
        if self.failure and not enabled:
            raise RuntimeError('off acknowledgement missing')
        self.is_enabled = enabled
        self.acknowledgements.append({'command': 'enable' if enabled else 'disable', 'reply': 'ACK'})


class VacuumFixture:
    def setUp(self):
        directory = tempfile.TemporaryDirectory()
        self.addCleanup(directory.cleanup)
        self.store = HardwareSafetyStore(Path(directory.name) / 'safety.json')
        self.vacuum = VacuumInterface()
        self.cfg = VacuumConfig(port='COM2', baud=9600, auto_pump=True)
        self.backend = VacuumBackendAdapter(self.cfg)
        self.backend._safety_store = self.store
        self.addCleanup(self.release_owned)

    def connect(self):
        with patch('rapid_main.diagnostic_services.VacuumController', return_value=self.vacuum):
            self.backend.connect()

    def release_owned(self):
        self.vacuum.failure = False
        self.vacuum.before_output = None
        session = self.backend._hold_session
        if session and session.active:
            try:
                session.close()
            finally:
                session._release()


class MainVacuumSafetyTests(VacuumFixture, unittest.TestCase):
    def test_constructor_never_connects_or_enables_saved_autopump(self):
        with patch('rapid_main.diagnostic_services.VacuumController') as factory:
            backend = VacuumBackendAdapter(self.cfg)
        factory.assert_not_called()
        self.assertFalse(backend.is_connected())
        self.assertFalse(backend.output_state_known)
        self.cfg.port = 'COM9'
        self.assertEqual(backend._binding()['port'], 'COM2')

    def test_connect_does_not_guess_off_or_reset_outputs(self):
        self.connect()
        self.assertEqual(self.vacuum.calls, [('connect', 'COM2', 9600)])
        with self.assertRaisesRegex(HardwareUnavailableError, 'not been acknowledged'):
            self.backend.is_pump_on()
        self.assertIsNone(self.store.pending())
        from rapid_main.diagnostic_services import read_vacuum_snapshot
        snapshot = read_vacuum_snapshot(self.backend, warn_threshold=0.)
        self.assertTrue(snapshot.fault)
        self.assertIn('output state is unverified', snapshot.fault_reason)
        self.assertNotIn('pump is off', snapshot.fault_reason)

    def test_hold_claims_before_output_and_releases_only_after_immutable_evidence(self):
        self.connect()
        def observe(enabled):
            self.assertEqual(self.store.pending()['profile']['helper'], 'rapid_main_vacuum')
            with self.assertRaises(HardwareSafetyError):
                with self.store.operation_lease():
                    pass
        self.vacuum.before_output = observe
        self.backend.set_pump(True)
        self.assertTrue(self.backend.outputs_held)
        self.assertTrue(self.backend.is_pump_on())
        self.assertTrue(self.store.pending()['record']['outputs_held'])
        self.backend.set_pump(True)
        self.assertEqual(self.vacuum.calls.count(('output', True)), 1)
        with self.assertRaisesRegex(RuntimeError, 'Release'):
            self.backend.disconnect()
        self.backend.set_pump(False)
        self.assertFalse(self.backend.outputs_held)
        self.assertFalse(self.backend.is_pump_on())
        self.assertIsNone(self.store.pending())
        record = self.store.read()['record']
        self.assertTrue(record['safe_state_confirmed'])
        self.assertIn('no pressure telemetry', record['vacuum_evidence_basis'])
        self.assertTrue(list((self.store.path.parent / 'hardware_diagnostics').glob('station-*/artifact_index.json')))

    def test_failed_release_keeps_lease_and_requires_explicit_retry(self):
        self.connect()
        self.backend.set_pump(True)
        self.vacuum.failure = True
        with self.assertRaisesRegex(HardwareSafetyError, 'release remains unverified'):
            self.backend.set_pump(False)
        self.assertTrue(self.backend.outputs_held)
        self.assertFalse(self.backend.output_state_known)
        with self.assertRaises(HardwareSafetyError):
            with self.store.operation_lease():
                pass
        with self.assertRaisesRegex(RuntimeError, 'failed vacuum release'):
            self.backend.set_pump(True)
        self.vacuum.failure = False
        self.backend.recover_release()
        self.assertIsNone(self.store.pending())

    def test_publication_failure_after_off_retains_restart_latch_without_live_hold(self):
        self.connect()
        self.backend.set_pump(True)
        with patch('rapidpy_common.vacuum_diagnostic_safety.publish_diagnostic_record', side_effect=HardwareSafetyError('disk full')):
            with self.assertRaisesRegex(HardwareSafetyError, 'disk full'):
                self.backend.set_pump(False)
        self.assertFalse(self.backend.outputs_held)
        self.assertIsNotNone(self.store.pending())
        self.backend.recover_release()
        self.assertIsNone(self.store.pending())
        self.assertEqual(self.vacuum.calls.count(('output', True)), 1)

    def test_restart_recovery_matches_original_port_and_never_enables_outputs(self):
        self.connect()
        self.backend.set_pump(True)
        self.backend._hold_session._release()  # Simulate process-owned lease loss.
        wrong = VacuumBackendAdapter(VacuumConfig(port='COM9', baud=9600))
        wrong._safety_store = self.store
        with patch('rapid_main.diagnostic_services.VacuumController') as factory:
            with self.assertRaisesRegex(HardwareSafetyError, 'original wiring'):
                wrong.connect()
        factory.assert_not_called()
        restarted = VacuumBackendAdapter(self.cfg)
        restarted._safety_store = self.store
        restored = VacuumInterface()
        with patch('rapid_main.diagnostic_services.VacuumController', return_value=restored):
            restarted.connect()
        self.assertFalse(restarted.output_state_known)
        restarted.recover_release()
        self.assertEqual(restored.calls, [('connect', 'COM2', 9600), ('output', False)])
        self.assertIsNone(self.store.pending())

    def test_updown_hold_cannot_be_recovered_by_main_panel_on_same_port(self):
        self.store.begin('station_diagnostic', {}, {'helper': 'updown_control', 'resources': {'vacuum': {'port': 'COM2', 'baud': 9600}}})
        with patch('rapid_main.diagnostic_services.VacuumController') as factory:
            with self.assertRaises(HardwareSafetyError):
                self.backend.connect()
        factory.assert_not_called()
        self.assertFalse(self.backend.can_recover_pending)

    def test_reusing_hold_session_does_not_touch_a_previous_unbound_lift(self):
        self.connect()
        session = VacuumHoldSession(self.vacuum, store=self.store)
        session.start()
        session.close()
        old_lift = Mock()
        session._lift = old_lift
        session._axis = Mock()
        session.start()
        session.close()
        self.assertEqual(old_lift.mock_calls, [])
        record = self.store.read()['record']
        self.assertEqual(len(record['observations'][-1]['vacuum_acknowledgements']), 2)

    def test_recovery_panel_exemption_does_not_unlock_other_hardware_controls(self):
        from rapid_main.app import MainWindow
        from rapid_main.device_ownership import DeviceOwnershipManager, DeviceOwnershipError
        self.connect()
        self.backend.set_pump(True)
        window = SimpleNamespace(_measurement_backend=SimpleNamespace(has_unresolved_hardware_fault=True),
            _vacuum_backend=self.backend, _ownership=DeviceOwnershipManager())
        lease = MainWindow.acquire_device(window, 'vacuum', 'vacuum_panel')
        lease.release()
        for resource, owner in [('measurement', 'vacuum_panel'), ('changer', 'dc_motors_panel'), ('vacuum', 'other'), ('squid', 'squid_panel'), ('susceptibility', 'susceptibility_panel')]:
            with self.assertRaises(DeviceOwnershipError):
                MainWindow.acquire_device(window, resource, owner)

    def test_held_checkpoints_preserve_evidence_chain_without_growing_live_history(self):
        self.connect()
        self.backend.set_pump(True)
        session = self.backend._hold_session
        previous_id = self.store.pending()['record']['treatment_id']
        for index in range(12):
            session.observations.append({'checkpoint': index})
            record = session._publish()
            self.assertEqual(record.observations[0]['previous_record_id'], previous_id)
            self.assertEqual(len(session.observations), 1)
            previous_id = record.treatment_id
        self.assertEqual(len(list((self.store.path.parent / 'hardware_diagnostics').glob('station-*/artifact_index.json'))), 13)

    def test_pending_recovery_blocks_simulation_but_allows_hardware_mode(self):
        import os
        from rapid_main.app import MainWindow
        self.connect()
        self.backend.set_pump(True)
        window = SimpleNamespace(_shutdown_cleanup_requested=False, _has_active_automation=lambda: False,
            _owned_dialog_leases={}, _external_process_leases={})
        with patch.dict(os.environ, RAPID_SAFETY_STATE=str(self.store.path)):
            self.assertIn('No-Comm', MainWindow._operating_mode_change_blocker(window, True))
            self.assertEqual(MainWindow._operating_mode_change_blocker(window, False), '')

    def test_corrupt_journal_becomes_a_visible_device_ownership_error(self):
        from rapid_main.app import MainWindow
        from rapid_main.device_ownership import DeviceOwnershipManager, DeviceOwnershipError
        self.store.path.write_text('{broken', encoding='utf-8')
        window = SimpleNamespace(_measurement_backend=None, _vacuum_backend=self.backend,
                                 _ownership=DeviceOwnershipManager())
        with self.assertRaisesRegex(DeviceOwnershipError, 'cannot be verified'):
            MainWindow.acquire_device(window, 'vacuum', 'vacuum_panel')
        self.assertFalse(window._ownership.is_owned('vacuum'))
        self.assertEqual(self.vacuum.calls, [])


class MainVacuumWorkerTests(VacuumFixture, unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.app = QtWidgets.QApplication.instance() or QtWidgets.QApplication([])
        cls.previous_quit = cls.app.quitOnLastWindowClosed()
        cls.app.setQuitOnLastWindowClosed(False)

    @classmethod
    def tearDownClass(cls):
        cls.app.setQuitOnLastWindowClosed(cls.previous_quit)

    def events_until(self, condition, seconds=5):
        deadline = time.monotonic() + seconds
        while not condition() and time.monotonic() < deadline:
            self.app.processEvents()
            time.sleep(.01)
        self.assertTrue(condition())

    def close_dialog(self, dialog):
        self.vacuum.failure = False
        self.vacuum.before_output = None
        dialog.close()
        self.events_until(lambda: not dialog.isVisible() and dialog._command_thread is None)
        dialog.deleteLater()

    def test_corrupt_journal_panel_shows_fault_without_initializing_outputs(self):
        self.store.path.write_text('{broken', encoding='utf-8')
        dialog = VacuumDialog(backend=self.backend)
        try:
            self.assertIn('Hardware safety state is unverified', dialog._status_lbl.text())
            self.assertFalse(dialog._pump_btn.isEnabled())
            self.assertEqual(self.vacuum.calls, [])
        finally:
            dialog.close()
            dialog.deleteLater()

    def test_close_during_enable_waits_then_releases_before_disconnect(self):
        self.connect()
        entered, release = threading.Event(), threading.Event()
        def observe(enabled):
            if enabled:
                entered.set()
                release.wait(5)
        self.vacuum.before_output = observe
        dialog = VacuumDialog(backend=self.backend)
        dialog.show()
        try:
            dialog._on_pump_toggle(True)
            self.assertTrue(entered.wait(2))
            dialog.close()
            self.assertTrue(dialog.isVisible())
            self.assertTrue(dialog._command_thread.isRunning())
            self.assertTrue(self.backend.is_connected())
            release.set()
            self.events_until(lambda: not dialog.isVisible() and dialog._command_thread is None)
            self.assertEqual(self.vacuum.calls, [('connect', 'COM2', 9600), ('output', True), ('output', False), ('disconnect',)])
            self.assertIsNone(self.store.pending())
        finally:
            release.set()
            self.close_dialog(dialog)

    def test_failed_release_aborts_parent_shutdown_without_automatic_command_retries(self):
        from rapid_main.app import MainWindow
        config = AppConfig()
        config.general.nocomm = True
        with patch.object(AppConfig, 'load', return_value=config):
            window = MainWindow()
        window._confirm_shutdown = lambda **kwargs: True
        self.connect()
        self.backend.set_pump(True)
        self.vacuum.failure = True
        window.show()
        window._run_owned_dialog('vacuum', 'vacuum_panel', lambda owner: VacuumDialog(owner, backend=self.backend), modal=False)
        dialog = window._owned_dialogs['vacuum']
        try:
            window.close()
            self.events_until(lambda: not window._shutdown_cleanup_requested)
            self.assertTrue(window.isVisible())
            self.assertTrue(dialog.isVisible())
            self.assertTrue(self.backend.outputs_held)
            self.assertTrue(dialog._release_btn.isEnabled())
            count = self.vacuum.calls.count(('output', False))
            deadline = time.monotonic() + .35
            while time.monotonic() < deadline:
                self.app.processEvents()
                time.sleep(.01)
            self.assertEqual(self.vacuum.calls.count(('output', False)), count)
            old_backend = window._vacuum_backend
            window._on_nocomm_toggled(False)
            self.assertTrue(window.config.general.nocomm)
            self.assertIs(window._vacuum_backend, old_backend)
            self.vacuum.failure = False
            dialog._release()
            self.events_until(lambda: dialog._command_thread is None)
            window.close()
            self.events_until(lambda: not window.isVisible())
            self.assertEqual(window._owned_dialog_leases, {})
        finally:
            self.vacuum.failure = False
            if window._owned_dialogs:
                window.close()
                self.events_until(lambda: not window._owned_dialogs)
            window.close()
            window.deleteLater()
