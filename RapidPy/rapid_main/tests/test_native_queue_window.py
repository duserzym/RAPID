"""MainWindow reserves, starts, loads and settles its original native queue."""
import time
import threading
import os
from pathlib import Path
import unittest
from types import SimpleNamespace
from unittest.mock import patch
from PySide6 import QtCore, QtWidgets

from rapid_main.app import MainWindow
from rapid_main.device_ownership import DeviceOwnershipError
from rapid_main.queue_compiler import QueueSample, QueueOptions
from rapid_main.data_model import SpecimenMeta
from rapid_main.io.specimen_writer import write_header
from tests import test_queue_startup as startup_fixture


RESOURCES = ('measurement', 'changer', 'af_demag', 'vacuum', 'squid', 'susceptibility')


class NativeQueueWindowTests(startup_fixture.QueueStartupFixture, unittest.TestCase):
    def test_invalid_average_count_fails_before_root_operator_prompt_and_outputs(self):
        before = self.store.read(), len(self.commands), len(self.vacuum_serial.writes)
        prompts = []
        with patch.object(QtWidgets.QMessageBox, 'critical', lambda *args: None), patch.object(
                QtWidgets.QMessageBox, 'question', lambda *args: prompts.append(True)):
            for count in (0, True, 32768, 2.5):
                self.assertFalse(self.window.start_queue_run([QueueSample('S1', 'F1', 1, avg_steps=count)], QueueOptions()))
        self.assertEqual(prompts, [])
        self.assertEqual((self.store.read(), len(self.commands), len(self.vacuum_serial.writes)), before)
        self.assertTrue(all(not self.window._ownership.is_owned(item) for item in RESOURCES))

    def test_original_avgsteps_are_frozen_for_holder_and_measurement_handoff(self):
        self.window.config.squid.samples_per_pos = 7
        with self.arm_xy_edges(), patch.object(QtWidgets.QMessageBox, 'question',
                lambda *args: QtWidgets.QMessageBox.StandardButton.Yes):
            self.assertTrue(self.window.start_queue_run([QueueSample('S1', 'F1', 1, avg_steps=3)], QueueOptions()))
            self.wait_for(lambda: bool(self.started), timeout=30)
        self.assertEqual(self.holder_averages, [3])
        self.assertEqual(self.window.queue_measurement_avg_steps('S1'), 3)
        self.assertEqual(self.store.read()['plan']['commands']['samples'][0]['avg_steps'], 3)
        before = len(self.commands), len(self.vacuum_serial.writes)
        command = self.window._queue_current_command
        command.avg_steps = 4
        with self.assertRaisesRegex(ValueError, 'journaled queue'):
            self.window.queue_measurement_avg_steps('S1')
        command.avg_steps = 3
        holder = next(item for item in self.window._queue_plan if item.command_type == 'Holder')
        holder.avg_steps = 4
        with self.assertRaisesRegex(ValueError, 'journaled queue'):
            self.window.queue_measurement_avg_steps('S1')
        self.assertEqual((len(self.commands), len(self.vacuum_serial.writes)), before)
        holder.avg_steps = 3

    def test_invalid_specimen_output_blocks_before_root_operator_prompt_or_io(self):
        before = self.store.read(), len(self.commands), len(self.vacuum_serial.writes)
        prompts = []
        with patch.object(QtWidgets.QMessageBox, 'critical', lambda *args: None), patch.object(
                QtWidgets.QMessageBox, 'question', lambda *args: prompts.append(True)):
            for name in ('../outside', 'C:outside', 'NUL', 'provenance.json'):
                self.assertFalse(self.window.start_queue_run([QueueSample(name, 'manual', 1)], QueueOptions()))
        self.assertEqual(prompts, [])
        self.assertEqual((self.store.read(), len(self.commands), len(self.vacuum_serial.writes)), before)
        self.assertTrue(all(not self.window._ownership.is_owned(item) for item in RESOURCES))

    def test_source_index_snapshot_and_digest_are_required_for_native_handoff(self):
        path = self.store.path.parent / 'index.sam'
        path.write_text('S1 2 OriginalUnit OriginalSite\n', encoding='latin-1')
        identity = os.path.normcase(str(path.resolve()))
        sample = QueueSample('S1', identity, 1, source_file=identity)
        header = path.parent / 'S1'
        write_header(header, SpecimenMeta('S1', comment='Original header', volume=8.2, core_plate_strike=123))
        with self.arm_xy_edges(), patch.object(QtWidgets.QMessageBox, 'question', return_value=QtWidgets.QMessageBox.StandardButton.Yes):
            self.assertTrue(self.window.start_queue_run([sample], QueueOptions()))
            self.wait_for(lambda: bool(self.started), timeout=30)
        source, registrations = self.window.queue_measurement_source('S1')
        self.assertEqual(source, path.resolve())
        self.assertEqual(registrations.entries[0].formation, 'OriginalUnit')
        metadata = self.window.queue_measurement_metadata('S1')
        self.assertEqual((metadata.meta.volume, metadata.meta.core_plate_strike), (8.2, 123))
        provenance = self.window.queue_measurement_provenance('S1')
        self.assertEqual(provenance['source_file'], str(path.resolve()))
        self.assertEqual(provenance['specimen']['meta']['volume'], 8.2)
        provenance['specimen']['meta']['volume'] = 999
        self.assertEqual(self.window.queue_measurement_provenance('S1')['specimen']['meta']['volume'], 8.2)
        before = len(self.commands), len(self.vacuum_serial.writes)
        captured = self.window._queue_source_indexes[identity]
        captured['specimens']['S1']['meta']['volume'] = 99
        with self.assertRaisesRegex(ValueError, 'journaled index metadata'):
            self.window.queue_measurement_metadata('S1')
        captured['specimens']['S1']['meta']['volume'] = 8.2
        original_header = header.read_bytes()
        header.write_bytes(original_header + b'changed\n')
        with self.assertRaisesRegex(ValueError, 'header changed'):
            self.window.queue_measurement_metadata('S1')
        header.write_bytes(original_header)
        captured['entries'][0]['formation'] = 'Changed'
        with self.assertRaisesRegex(ValueError, 'journaled index metadata'):
            self.window.queue_measurement_source('S1')
        captured['entries'][0]['formation'] = 'OriginalUnit'
        path.write_text('S1 2 ChangedUnit ChangedSite\n', encoding='latin-1')
        with self.assertRaisesRegex(ValueError, 'index changed'):
            self.window.queue_measurement_source('S1')
        self.assertEqual((len(self.commands), len(self.vacuum_serial.writes)), before)
        self.assertTrue(all(self.window._ownership.is_owned(item) for item in RESOURCES))

    def test_missing_or_duplicate_source_specimen_blocks_before_root_or_operator_prompt(self):
        path = self.store.path.parent / 'index.sam'
        identity = os.path.normcase(str(path.resolve()))
        before = self.store.read(), len(self.commands), len(self.vacuum_serial.writes)
        for text in ('OTHER\n', 'S1\nS1\n'):
            path.write_text(text, encoding='latin-1')
            with self.subTest(text=text), patch.object(QtWidgets.QMessageBox, 'critical', return_value=None), patch.object(QtWidgets.QMessageBox, 'question') as question:
                self.assertFalse(self.window.start_queue_run([QueueSample('S1', identity, 1, source_file=identity)], QueueOptions()))
                question.assert_not_called()
            self.assertEqual((self.store.read(), len(self.commands), len(self.vacuum_serial.writes)), before)
            self.assertTrue(all(not self.window._ownership.is_owned(item) for item in RESOURCES))

    @classmethod
    def setUpClass(cls):
        cls.app = QtWidgets.QApplication.instance() or QtWidgets.QApplication([])
        cls.quit_policy = cls.app.quitOnLastWindowClosed()
        cls.app.setQuitOnLastWindowClosed(False)

    @classmethod
    def tearDownClass(cls):
        cls.app.setQuitOnLastWindowClosed(cls.quit_policy)

    def setUp(self):
        super().setUp()
        self.window = MainWindow()
        self.window.config = self.backend._config
        self.window.config.general.operator = 'Operator'
        self.window._measurement_backend = self.backend
        self.window._vacuum_backend = self.vacuum
        self.window._sequence_labels = ['NRM']
        self.started = []
        self.holders = []
        self.holder_averages = []
        def measure_holder(marker, *, averaging_cycles=None):
            self.holders.append(marker)
            self.holder_averages.append(averaging_cycles)
        self.backend.measure_queue_holder = measure_holder
        self.window._measurement = SimpleNamespace(is_active=lambda: False, halt_run=lambda: None,
            start_measurement_for_sample=self.start_measurement)
        self.backend._operator, self.backend._treatment_label = '', ''
        self.backend._backend_errors = []
        self.addCleanup(self.cleanup_window)

    def start_measurement(self, sample, **kwargs):
        with self.backend.queue_worker_claim():
            geometry = self.backend._queue_specimen_geometry
            self.started.append((sample, geometry.context.original_slot, geometry.context.file_id,
                self.backend._sample_height(), self.backend._direction_up, kwargs['owner']))
        return True

    def wait_for(self, predicate, *, timeout=15):
        deadline = time.monotonic() + timeout
        while not predicate() and time.monotonic() < deadline:
            self.app.processEvents()
            # Native fixture replaces hardware's shared time.sleep; use a real
            # wait here so Qt polling yields rather than growing Mock call logs.
            threading.Event().wait(.005)
        self.assertTrue(predicate())

    def cleanup_window(self):
        self.window._queue_advance_timer.stop()
        if self.window._queue_command_thread is not None:
            self.window._queue_command_worker.stop()
            self.wait_for(lambda: self.window._queue_command_thread is None)
        # Test-only disposal of retained injected interfaces after the worker exits.
        if self.window._queue_native_session is not None:
            self.window._queue_native_session.release()
            self.window._queue_native_session = None
        if self.window._queue_lease is not None:
            self.window._queue_lease.release()
            self.window._queue_lease = None
        self.window._queue_active = False
        self.window.deleteLater()
        QtCore.QCoreApplication.sendPostedEvents(None, QtCore.QEvent.DeferredDelete)
        self.app.processEvents()

    def start(self, *, answer=QtWidgets.QMessageBox.StandardButton.Yes, do_up=True):
        with self.arm_xy_edges(), patch.object(QtWidgets.QMessageBox, 'question', return_value=answer), patch.object(QtWidgets.QMessageBox, 'critical', return_value=None):
            started = self.window.start_queue_run([QueueSample('S1', 'F1', 1, do_up=do_up)], QueueOptions())
            if started:
                self.wait_for(lambda: bool(self.started) or self.window._queue_active is False)
                self.assertTrue(self.started, self.window._sb_status.text())
        return started

    def test_native_start_holds_all_resources_and_loads_original_slot_file_and_height(self):
        self.assertTrue(self.start(do_up=False), self.window._sb_status.text())
        self.assertEqual(self.started, [('S1', 1, 'F1', 2800, False, 'queue_workflow')])
        self.assertEqual(self.holders, [0])
        self.assertTrue(all(self.window._ownership.owner_of(item) == 'queue_workflow' for item in RESOURCES))
        self.assertTrue(self.vacuum.is_valve_connected())
        self.assertIsNone(self.window._queue_vacuum_fault_reason())
        self.assertIsNotNone(self.window._queue_native_session._lease)

    def test_native_down_without_original_up_stops_before_loading(self):
        self.window.config.general.data_dir = str(self.store.path.parent / 'paired-output')
        with self.arm_xy_edges(), patch.object(QtWidgets.QMessageBox, 'question',
                lambda *args: QtWidgets.QMessageBox.StandardButton.Yes):
            self.assertTrue(self.window.start_queue_run([QueueSample('S1', 'F1', 1, do_both=True)], QueueOptions()))
            self.wait_for(lambda: bool(self.started), timeout=30)
        up = self.window.queue_measurement_side('S1')
        self.assertTrue(up.is_up)
        self.assertEqual(up.pair.queue_run_id, self.store.read()['run_id'])

        self.window._queue_current_command = self.window._queue_plan[up.pair.down_index]
        refused = []
        before = len(self.commands), len(self.vacuum_serial.writes), self.store.read()
        with patch.object(self.window, '_finalize_queue_run', lambda *args, **kw: refused.append(kw['reason'])):
            self.window._run_measure_command()
        self.assertIn('accepted Up acquisition is required', refused[0])
        self.assertEqual((len(self.commands), len(self.vacuum_serial.writes), self.store.read()), before)

    def test_declining_empty_rod_confirmation_never_begins_or_actuates(self):
        before = self.store.read(), len(self.commands), len(self.vacuum_serial.writes)
        self.assertFalse(self.start(answer=QtWidgets.QMessageBox.StandardButton.No))
        self.assertEqual((self.store.read(), len(self.commands), len(self.vacuum_serial.writes)), before)
        self.assertTrue(all(self.window._ownership.owner_of(item) is None for item in RESOURCES))

    def test_native_legacy_down_persists_original_receipt_and_checks_it_before_load(self):
        from tests.test_legacy_up import FIXTURE

        path = self.store.path.parent / 'NRM.UP'
        path.write_text(FIXTURE.replace('A|', 'S1|'), encoding='ascii')
        self.window.config.general.data_dir = str(self.store.path.parent / 'output')
        sample = QueueSample('S1', 'F1', 1, do_up=False, do_both=True)
        before = len(self.commands), list(self.vacuum_serial.writes), self.store.read()
        prompts = []
        choose = lambda *args: (str(path), '')
        with (patch.object(QtWidgets.QFileDialog, 'getOpenFileName', choose),
              patch.object(QtWidgets.QInputDialog, 'getText', lambda *args, **kwargs: ('0.25', True)),
              patch.object(QtWidgets.QMessageBox, 'critical', lambda *args: None),
              patch.object(QtWidgets.QMessageBox, 'question', lambda *args: prompts.append(True))):
            self.assertFalse(self.window.start_queue_run([sample], QueueOptions()))
        self.assertEqual(prompts, [])
        self.assertIn('historical range factor', self.window._sb_status.text())
        self.assertEqual((len(self.commands), self.vacuum_serial.writes, self.store.read()), before)

        with (self.arm_xy_edges(), patch.object(QtWidgets.QFileDialog, 'getOpenFileName', choose),
              patch.object(QtWidgets.QInputDialog, 'getText', lambda *args, **kwargs: (kwargs['text'], True)),
              patch.object(QtWidgets.QMessageBox, 'question', lambda *args: QtWidgets.QMessageBox.StandardButton.Yes)):
            self.assertTrue(self.window.start_queue_run([sample], QueueOptions()), self.window._sb_status.text())
            self.wait_for(lambda: bool(self.started), timeout=30)
        state = self.store.read()
        side = self.window.queue_measurement_side('S1')
        self.assertFalse(side.is_up)
        self.assertEqual(side.pair.queue_run_id, state['run_id'])
        self.assertEqual(state['plan']['commands']['legacy_up_sources']['F1']['sha256'], side.legacy_up.sha256)
        self.assertEqual(self.started[0][4], False)

        command = self.window._queue_current_command
        with patch.object(command, 'legacy_up_required', False):
            with self.assertRaisesRegex(ValueError, 'original queue plan'):
                self.window._queue_measurement_side(command)
        path.write_bytes(b'changed')
        self.window._queue_loaded_command = None
        refusals = []
        before = len(self.commands), list(self.vacuum_serial.writes), self.store.read()
        with patch.object(self.window, '_finalize_queue_run', lambda *args, **kwargs: refusals.append(kwargs['reason'])):
            self.window._run_measure_command()
        self.assertIn('selected legacy Up file changed', refusals[0])
        self.assertEqual((len(self.commands), self.vacuum_serial.writes, self.store.read()), before)

    def test_file_sequence_handoff_requires_original_journaled_command(self):
        self.assertTrue(self.start(), self.window._sb_status.text())
        self.window._sequence_labels = ['IRM100']
        self.assertEqual(self.window.queue_measurement_labels('S1'), ['NRM'])
        command = self.window._queue_current_command
        before = len(self.commands), len(self.vacuum_serial.writes)
        command.measurement_labels = ('AF20',)
        with self.assertRaisesRegex(ValueError, 'journaled queue sequence'):
            self.window.queue_measurement_labels('S1')
        self.assertEqual((len(self.commands), len(self.vacuum_serial.writes)), before)
        self.assertTrue(all(self.window._ownership.is_owned(item) for item in RESOURCES))


    def test_one_busy_resource_blocks_the_whole_start_without_partial_reservations(self):
        lease = self.window.acquire_device('susceptibility', 'diagnostic')
        self.addCleanup(lease.release)
        before = self.store.read(), len(self.commands)
        self.assertFalse(self.start())
        self.assertEqual((self.store.read(), len(self.commands)), before)
        self.assertEqual(self.window._ownership.owner_of('susceptibility'), 'diagnostic')
        self.assertTrue(all(self.window._ownership.owner_of(item) is None for item in RESOURCES[:-1]))

    def test_original_queue_can_reenter_but_other_controls_cannot(self):
        self.assertTrue(self.start(), self.window._sb_status.text())
        lease = self.window.acquire_device('measurement', 'queue_workflow')
        lease.release()
        for resource in RESOURCES:
            with self.assertRaises(DeviceOwnershipError): self.window.acquire_device(resource, 'other_panel')
        self.assertIsNotNone(self.window._queue_native_session._lease)

    def test_replaced_scientific_adapter_cannot_reenter_original_queue(self):
        self.assertTrue(self.start(), self.window._sb_status.text())
        original = self.backend._measurement
        self.backend._measurement = object()
        try:
            with self.assertRaises(DeviceOwnershipError):
                self.window.acquire_device('measurement', 'queue_workflow')
            self.assertTrue(all(self.window._ownership.is_owned(item) for item in RESOURCES))
        finally:
            self.backend._measurement = original

    def test_terminal_worker_finishes_root_before_releasing_all_device_and_os_leases(self):
        self.assertTrue(self.start(), self.window._sb_status.text())
        session = self.window._queue_native_session
        self.window.cancel_queue_run('test stop')
        self.assertTrue(all(self.window._ownership.is_owned(item) for item in RESOURCES))
        self.wait_for(lambda: self.window._queue_native_session is None)
        self.assertIsNone(session._lease)
        self.assertEqual(session.store.read()['status'], 'verified')
        self.assertTrue(all(not self.window._ownership.is_owned(item) for item in RESOURCES))
        self.assertIsNone(self.window._queue_command_thread)

    def test_failed_terminal_cutoff_retains_all_ownership_and_keeps_shutdown_open(self):
        self.assertTrue(self.start(), self.window._sb_status.text())
        session = self.window._queue_native_session
        self.cap_voltage = .5
        self.window.cancel_queue_run('test stop')
        self.wait_for(lambda: self.window._queue_command_thread is None)
        self.assertIs(self.window._queue_native_session, session)
        self.assertIsNotNone(session._lease)
        self.assertEqual(self.window._workflow_state, 'error')
        self.assertTrue(all(self.window._ownership.is_owned(item) for item in RESOURCES))
        self.assertFalse(self.window._release_queue_lease())
        with patch.object(QtWidgets.QMessageBox, 'question', return_value=QtWidgets.QMessageBox.StandardButton.Yes):
            event = QtCore.QEvent(QtCore.QEvent.Close)
            self.window.closeEvent(event)
        self.assertFalse(event.isAccepted())
        self.assertFalse(self.window._shutdown_cleanup_requested)
        self.assertTrue(self.vacuum.is_valve_connected())

    def test_wrong_table_mode_is_rejected_before_startup_or_operator_prompt(self):
        before = self.store.read(), len(self.commands)
        with patch.object(QtWidgets.QMessageBox, 'critical', return_value=None), patch.object(QtWidgets.QMessageBox, 'question') as question:
            self.assertFalse(self.window.start_queue_run([QueueSample('S1', 'F1', 1)], QueueOptions(use_xy_table=False)))
        question.assert_not_called()
        self.assertEqual((self.store.read(), len(self.commands)), before)

    def test_native_chain_slot_confirmation_precedes_start_and_handoff(self):
        self.use_chain_station()
        original = self.store.read(), len(self.commands), list(self.vacuum_serial.writes)
        specimen = QueueSample('S1', 'F1', 1)
        options = QueueOptions(use_xy_table=False)
        with (patch.object(QtWidgets.QMessageBox, 'question', return_value=QtWidgets.QMessageBox.StandardButton.Yes),
              patch.object(QtWidgets.QMessageBox, 'critical', return_value=None) as error,
              patch.object(QtWidgets.QInputDialog, 'getInt', side_effect=[(0, False), (99, True)])):
            self.assertFalse(self.window.start_queue_run([specimen], options))
            self.assertEqual((self.store.read(), len(self.commands), self.vacuum_serial.writes), original)
            self.assertTrue(all(not self.window._ownership.is_owned(item) for item in RESOURCES))
            self.assertTrue(self.window.start_queue_run([specimen], options))
            self.wait_for(lambda: bool(self.started) or not self.window._queue_active, timeout=30)
            self.assertTrue(self.started, self.window._sb_status.text())
            error.assert_not_called()
        self.assertEqual(self.started, [('S1', 1, 'F1', 2800, True, 'queue_workflow')])
        root = self.store.read()
        self.assertEqual(root['plan']['startup']['chain_slot'], 99)
        self.assertFalse(root['plan']['commands']['options']['use_xy_table'])
        self.assertEqual(root['profile']['stage_profiles']['motion']['helper'], 'rapid_main_queue_chain')
        self.assertNotIn('COM4', self.motor._connections)

    def test_ui_shutdown_retry_settles_original_close_without_replaying_outputs(self):
        self.assertTrue(self.start(), self.window._sb_status.text())
        session = self.window._queue_native_session
        with patch.object(self.vacuum_serial, 'close', side_effect=OSError('close failed')):
            self.window.cancel_queue_run('test stop')
            self.wait_for(lambda: self.window._queue_command_thread is None)
        self.assertIsNotNone(self.backend._queue_terminal_cleanup._close_token)
        self.assertTrue(self.window._retry_queue_close_action.isEnabled())
        before = len(self.commands), len(self.vacuum_serial.writes)
        self.assertTrue(self.window.retry_queue_shutdown())
        self.assertTrue(all(self.window._ownership.is_owned(item) for item in RESOURCES))
        self.wait_for(lambda: self.window._queue_native_session is None)
        self.assertEqual((len(self.commands), len(self.vacuum_serial.writes)), before)
        self.assertIsNone(session._lease)
        self.assertEqual(session.store.read()['status'], 'verified')
        self.assertTrue(all(not self.window._ownership.is_owned(item) for item in RESOURCES))
        self.assertFalse(self.window._retry_queue_close_action.isEnabled())

    def test_restart_menu_recovers_cutoff_and_releases_workers_but_keeps_queue_latched(self):
        startup = self.prepare()
        self.cap_voltage = .5
        self.assertFalse(self.command(self.backend.start_queue_lifetime).ok)
        interrupted = self.store.pending()
        startup.session.release()
        self.backend._safety_store = self.store
        self.backend._queue_startup = None
        self.backend._queue_instruments = None
        self.cap_voltage = 0

        prompts = []
        with patch.object(QtWidgets.QMessageBox, 'critical', lambda *args: None), patch.object(
                QtWidgets.QMessageBox, 'question', lambda *args: prompts.append(True)):
            self.assertFalse(self.window.start_queue_run([QueueSample('S1', 'F1', 1)], QueueOptions()))
        self.assertEqual(prompts, [])

        vacuum_before = list(self.vacuum_serial.writes)
        self.assertTrue(self.window.recover_interrupted_queue(), self.window._sb_status.text())
        self.assertTrue(all(self.window._ownership.owner_of(key) == 'queue_restart_recovery' for key in RESOURCES))
        self.wait_for(lambda: self.window._queue_command_thread is None)
        record = self.store.pending()
        self.assertEqual(record['token'], interrupted['token'])
        self.assertEqual(record['stage']['token'], interrupted['stage']['token'])
        self.assertEqual(record['stage']['status'], 'verified')
        self.assertEqual(self.vacuum_serial.writes, vacuum_before)
        self.assertFalse(self.motor._connections)
        self.assertTrue(all(self.window._ownership.owner_of(key) is None for key in RESOURCES))
        with self.store.operation_lease():
            self.assertIn('still requires station recovery', self.window._sb_status.text())

    def test_restart_shutdown_requires_confirmation_and_settles_original_queue(self):
        startup = self.prepare()
        startup.session.release()
        self.backend._safety_store = self.store
        self.backend._queue_startup = None
        self.backend._queue_instruments = None
        original = self.store.pending()
        before = list(self.commands), list(self.vacuum_serial.writes)
        with patch.object(QtWidgets.QMessageBox, 'question', lambda *args: QtWidgets.QMessageBox.StandardButton.No):
            self.assertFalse(self.window.shutdown_interrupted_queue())
        self.assertEqual((self.commands, self.vacuum_serial.writes), before)
        self.assertEqual(self.store.pending(), original)

        with patch.object(QtWidgets.QMessageBox, 'question', lambda *args: QtWidgets.QMessageBox.StandardButton.Yes):
            self.assertTrue(self.window.shutdown_interrupted_queue(), self.window._sb_status.text())
        self.assertTrue(all(self.window._ownership.owner_of(key) == 'queue_restart_recovery' for key in RESOURCES))
        self.wait_for(lambda: self.window._queue_command_thread is None)
        self.assertIsNone(self.store.pending(), self.window._sb_status.text())
        record = self.store.read()
        self.assertEqual(record['token'], original['token'])
        self.assertEqual(record['stage']['status'], 'aborted')
        self.assertTrue(all(self.window._ownership.owner_of(key) is None for key in RESOURCES))

    def test_shutdown_interrupted_live_queue_retains_ownership_until_worker_exits(self):
        self.assertTrue(self.start(), self.window._sb_status.text())
        session = self.window._queue_native_session
        self.lift_serial().position, self.lift_serial().status = 0, 16
        with session.claim():
            token = session.child_store.begin('acquisition', {'action': 'squid'},
                self.backend._acquisition_safety_profile())
        self.window._finalize_queue_run('error', reason='Interrupted test acquisition', return_to_safe=False)
        pending = self.store.pending()
        before = len(self.commands), list(self.vacuum_serial.writes)
        with patch.object(QtWidgets.QMessageBox, 'question', lambda *args: QtWidgets.QMessageBox.StandardButton.No):
            self.assertFalse(self.window.shutdown_interrupted_queue())
        self.assertEqual((len(self.commands), self.vacuum_serial.writes), before)
        self.assertEqual(self.store.pending(), pending)

        with patch.object(QtWidgets.QMessageBox, 'question', lambda *args: QtWidgets.QMessageBox.StandardButton.Yes):
            self.assertTrue(self.window.shutdown_interrupted_queue(), self.window._sb_status.text())
        self.assertIs(self.window._queue_native_session, session)
        self.assertTrue(all(self.window._ownership.owner_of(key) == 'queue_workflow' for key in RESOURCES))
        self.wait_for(lambda: self.window._queue_command_thread is None)
        state = self.store.read()
        self.assertEqual(state['stage']['status'], 'aborted', self.window._sb_status.text())
        self.assertEqual(state['stage']['token'], token)
        self.assertIsNone(self.window._queue_native_session)
        self.assertIsNone(session._lease)
        self.assertTrue(all(self.window._ownership.owner_of(key) is None for key in RESOURCES))

    def test_ui_shutdown_retry_rejects_unfinished_field_stage(self):
        self.assertTrue(self.start(), self.window._sb_status.text())
        self.cap_voltage = .5
        self.window.cancel_queue_run('test stop')
        self.wait_for(lambda: self.window._queue_command_thread is None)
        before = len(self.commands), len(self.vacuum_serial.writes)
        self.assertFalse(self.window.retry_queue_shutdown())
        self.assertEqual((len(self.commands), len(self.vacuum_serial.writes)), before)
        self.assertTrue(all(self.window._ownership.is_owned(item) for item in RESOURCES))

    def test_ui_shutdown_retry_rejects_active_run_without_io(self):
        self.assertTrue(self.start(), self.window._sb_status.text())
        before = len(self.commands), len(self.vacuum_serial.writes)
        self.assertFalse(self.window.retry_queue_shutdown())
        self.assertEqual((len(self.commands), len(self.vacuum_serial.writes)), before)

    def test_pause_after_loaded_pose_defers_measurement_and_resume_does_not_reload(self):
        original = self.backend.load_queue_specimen
        entered, release = threading.Event(), threading.Event()
        calls = []
        def load(*args, **kwargs):
            calls.append(args)
            result = original(*args, **kwargs)
            entered.set()
            if not release.wait(5): raise RuntimeError('test release timeout')
            return result
        self.backend.load_queue_specimen = load
        timer = QtCore.QTimer(self.window)
        def pause_loaded():
            if entered.is_set():
                self.window._queue_paused = True
                release.set()
                timer.stop()
        timer.timeout.connect(pause_loaded)
        timer.start(10)
        try:
            with self.arm_xy_edges(), patch.object(QtWidgets.QMessageBox, 'question', return_value=QtWidgets.QMessageBox.StandardButton.Yes):
                self.assertTrue(self.window.start_queue_run([QueueSample('S1', 'F1', 1)], QueueOptions()))
                # This wait spans startup, holder, load and actual QThread exit.
                # Allow two phase budgets; preserve every pause/no-reload assertion.
                self.wait_for(lambda: entered.is_set() and self.window._queue_command_thread is None, timeout=30)
            self.assertEqual(self.started, [])
            self.assertTrue(self.window._queue_meas_pending_start)
            self.window.toggle_queue_pause()
            self.assertEqual(len(calls), 1)
            self.assertEqual(len(self.started), 1)
            self.assertFalse(self.window._queue_meas_pending_start)
        finally:
            release.set()
            timer.stop()
