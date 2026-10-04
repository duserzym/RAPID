"""Bind native measurement geometry to the original transfer/field owner."""
from types import SimpleNamespace
import unittest
from unittest.mock import patch

from rapid_main.queue_command_worker import QueueCommandWorker
from rapid_main.queue_transfer_coordinator import QueueTransferCoordinator
from rapid_main.queue_xy_reference import QueueXYReference
from rapid_main.acquisition import RecoveryRecord
from tests.test_queue_lift_transfer import QueueLiftFixture
from tests import test_queue_field_outputs as field_fixture
from tests import test_queue_specimen_geometry as geometry_fixture


class QueueBackendCoordinatorFixture(QueueLiftFixture):
    def additional_stage_profiles(self):
        scientific = geometry_fixture.QueueSpecimenGeometryTests.additional_stage_profiles(self)
        fields = field_fixture.QueueFieldOutputsTests.additional_stage_profiles(self)
        self.backend._af_demag = SimpleNamespace(_controller=self.af)
        self.backend._arm_bias = self.arm
        self.backend._pulse_irm = SimpleNamespace(daq=self.daq, relays=self.af)
        scientific['field_outputs'] = fields['field_outputs']
        return scientific

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
        self.coordinator = QueueTransferCoordinator(self.session, self.lift, self.reference, self.fields)
        with self.session.claim():
            proof = self.fields.verify_off(self.session)
            self.reference.home(self.session, empty_rod_confirmed=True, field_outputs_off_verified=proof)
            self.backend.bind_queue_coordinator(self.coordinator)

    def command(self, action):
        worker = QueueCommandWorker(self.backend, action, recover_on_error=False)
        worker.run()
        return worker

    def load_backend(self):
        worker = self.command(lambda: self.backend.load_queue_specimen(1, 'S1', file_id='F1'))
        self.assertTrue(worker.ok, worker.error)


class QueueBackendCoordinatorTests(QueueBackendCoordinatorFixture, unittest.TestCase):
    def test_real_squid_adapter_composes_without_io_then_connects_inside_pending_acquisition(self):
        from rapid_main.diagnostic_services import SquidBackendAdapter
        from updown_control.app import SquidMomentReader
        self.load_backend()
        reader = SquidMomentReader()
        adapter = SquidBackendAdapter(self.backend._config.squid)
        self.backend._measurement = adapter
        self.backend._backend_errors = []
        calls = []
        def connect(*args, **kwargs):
            root = self.session.store._queue(self.session.token)
            self.assertEqual(root['stage']['family'], 'acquisition')
            self.assertEqual(root['stage']['status'], 'pending')
            self.assertEqual(root['stage']['token'], self.backend._geometry_stage_token)
            calls.append((args, kwargs))
            reader.raw_client._serial = SimpleNamespace(is_open=True)
        with patch('rapid_main.diagnostic_services.SquidMomentReader', return_value=reader), \
                patch.object(reader, 'connect', connect), \
                patch.object(reader, 'take_baseline', side_effect=AssertionError('unexpected diagnostic read')):
            with self.backend.queue_worker_claim():
                preflight = self.backend.preflight()
                self.assertTrue(preflight.ok, preflight.blockers)
                self.assertIsNotNone(self.backend._bracketed)
                self.assertEqual(calls, [])
                with self.assertRaisesRegex(Exception, 'pending acquisition'):
                    self.backend._read_squid_unowned()
                self.assertEqual(calls, [])
                # No valid holder is installed: connect is staged, but scientific
                # reading must still fail and retain its acquisition owner.
                with self.assertRaises(Exception):
                    self.backend.read_squid()
            self.assertEqual(len(calls), 1)
            self.assertIs(adapter.raw_client, reader.raw_client)
            self.assertEqual(self.store.pending()['stage']['family'], 'acquisition')

    def test_borrowed_evidence_path_is_the_fixed_parent_journal_and_cannot_be_rebound(self):
        self.assertEqual(self.session.child_store.path, self.session.store.path)
        with self.assertRaises(AttributeError):
            self.session.child_store.path = self.path.parent / 'different.json'

    def test_owned_preflight_never_opens_or_tests_squid_before_acquisition_stage(self):
        self.load_backend()
        self.backend._backend_errors = []
        def compose():
            self.backend._bracketed = SimpleNamespace()
        with patch.object(self.backend, '_ensure_bracketed', compose), self.backend.queue_worker_claim():
            result = self.backend.preflight()
        self.assertTrue(result.ok, result.blockers)
        self.backend._measurement.test_connection.assert_not_called()

    def test_unclaimed_queue_preflight_cannot_test_squid_or_reopen_motors(self):
        self.backend._backend_errors = []
        with patch.object(self.motor, 'connect') as connect:
            result = self.backend.preflight()
        self.assertFalse(result.ok)
        self.backend._measurement.test_connection.assert_not_called()
        connect.assert_not_called()

    def test_worker_load_binds_measured_geometry_and_return_clears_only_after_support_release(self):
        self.load_backend()
        self.assertIs(self.backend._safety_store, self.session.child_store)
        with self.backend.queue_worker_claim():
            self.assertEqual(self.backend._sample_height(), 2800)
            self.assertEqual(self.backend._queue_specimen_geometry.positions('S1'), (-48600, -96100))
        worker = self.command(self.backend.return_to_safe_state)
        self.assertTrue(worker.ok, worker.error)
        self.assertIsNone(self.backend._queue_specimen_geometry)
        self.assertEqual(self.context().phase, 'clear')
        self.assertFalse(self.vacuum.is_valve_connected())
        self.assertTrue(self.vacuum.is_pump_on())
        self.assertIsNotNone(self.store.pending())

    def test_changed_scientific_settings_fail_before_any_transfer_or_cutoff(self):
        self.backend._config.squid.samples_per_pos += 1
        before = len(self.commands), len(self.vacuum_serial.writes)
        worker = self.command(lambda: self.backend.load_queue_specimen(1, 'S1'))
        self.assertFalse(worker.ok)
        self.assertIn('scientific', worker.error)
        self.assertEqual((len(self.commands), len(self.vacuum_serial.writes)), before)

    def test_replaced_native_field_controller_cannot_borrow_original_queue(self):
        self.backend._af_demag._controller = object()
        before = len(self.commands)
        worker = self.command(lambda: self.backend.load_queue_specimen(1, 'S1'))
        self.assertFalse(worker.ok)
        self.assertIn('field circuit', worker.error)
        self.assertEqual(len(self.commands), before)

    def test_loaded_geometry_cannot_be_overwritten_by_another_sample(self):
        self.load_backend()
        worker = self.command(lambda: self.backend.load_queue_specimen(1, 'S2'))
        self.assertFalse(worker.ok)
        self.assertEqual(self.context().sample_id, 'S1')
        self.assertTrue(self.vacuum.is_valve_connected())

    def test_failed_return_keeps_original_binding_and_grip(self):
        self.load_backend()
        with patch.object(self.motor, 'updown_move', side_effect=RuntimeError('motion failed')):
            worker = self.command(self.backend.return_queue_specimen)
        self.assertFalse(worker.ok)
        self.assertIn('motion failed', worker.error)
        self.assertIsNotNone(self.backend._queue_specimen_geometry)
        self.assertTrue(self.vacuum.is_valve_connected())
        self.assertEqual(self.store.pending()['stage']['status'], 'pending')

    def test_cancel_hook_is_shared_with_transfer_and_cleared_after_worker(self):
        cancelled = False
        self.backend.set_halt_check(lambda: cancelled)
        cancelled = True
        with self.backend.queue_worker_claim(), self.assertRaisesRegex(InterruptedError, 'cancelled'):
            self.coordinator.load(1, 'S1')
        self.backend.set_halt_check(None)
        self.load_backend()
        self.assertIsNone(self.backend._halt_check)

    def test_verified_return_retains_transport_recovery_evidence_for_worker_artifacts(self):
        self.load_backend()
        first = RecoveryRecord(1, 'start-1', 'end-1', None, (), 'first recovery')
        second = RecoveryRecord(2, 'start-2', 'end-2', None, (), 'second recovery')
        self.backend._bracketed = SimpleNamespace(transport_recovery_records=(first, second),
                                                 _service=SimpleNamespace(set_cancel_check=lambda check: None))
        worker = self.command(self.backend.return_queue_specimen)
        self.assertTrue(worker.ok, worker.error)
        self.assertIsNone(self.backend._bracketed)
        self.assertEqual(self.backend.transport_recovery_records, (first, second))
        self.load_backend()
        self.assertEqual(self.backend.transport_recovery_records, (first, second))
