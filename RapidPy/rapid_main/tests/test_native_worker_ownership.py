"""Worker-thread claims around real journals, native stops and staged SQUID I/O."""
from dataclasses import replace
from types import SimpleNamespace
import threading
import unittest
from unittest.mock import patch
from PySide6 import QtCore

from rapid_main.acquisition import BlockContext
from rapid_main.data_model import SpecimenMeta
from rapid_main.measurement_worker import MeasurementWorker
from rapid_main.queue_command_worker import QueueCommandWorker
from rapid_main.squid_transport import BracketedSquidBackend
from tests.test_queue_lift_transfer import QueueLiftFixture
from tests import test_queue_specimen_geometry as geometry_fixture
from tests.test_acquisition import _build_service


class NativeWorkerOwnershipTests(QueueLiftFixture, unittest.TestCase):
    additional_stage_profiles = geometry_fixture.QueueSpecimenGeometryTests.additional_stage_profiles

    @classmethod
    def setUpClass(cls):
        cls.app = QtCore.QCoreApplication.instance() or QtCore.QCoreApplication([])

    def setUp(self):
        super().setUp()
        self.load()
        self.backend._safety_store = self.session.child_store
        with self.session.claim(): self.backend.bind_specimen_geometry(self.session)
        self.backend._holder_store = SimpleNamespace(require_valid=lambda **kw: SimpleNamespace(
            require_susceptibility=lambda: 0, susceptibility_evidence_id='holder-1'))
        self.backend._direction_up = True
        self.backend._susceptibility_records = []
        service, transport, vertical, turning, clock = _build_service()
        service._config = replace(service.config, zero_position=-48600, measurement_position=-96100)
        self.squid = BracketedSquidBackend(service, transport_retries=0,
            context_provider=lambda: BlockContext(sample_name='S1', run_id='run-1'))
        def ensure():
            self.session.child_store._owned()
            self.assertEqual(self.store.pending()['stage']['family'], 'acquisition')
            self.backend._bracketed = self.squid
        item = patch.object(self.backend, '_ensure_bracketed', ensure)
        item.start()
        self.addCleanup(item.stop)

    def run_command_thread(self, action, *, recover_on_error=False):
        thread = QtCore.QThread()
        worker = QueueCommandWorker(self.backend, action, recover_on_error=recover_on_error)
        worker.moveToThread(thread)
        thread.started.connect(worker.run)
        worker.settled.connect(thread.quit, QtCore.Qt.DirectConnection)
        worker.settled.connect(worker.deleteLater)
        released = []
        worker.settled.connect(lambda: released.append(self.session._owner is None), QtCore.Qt.DirectConnection)
        thread.start()
        self.assertTrue(thread.wait(10000))
        self.assertEqual(released, [True])
        return worker.ok, worker.error

    def test_command_worker_acquires_original_session_on_its_qthread(self):
        caller = threading.get_ident()
        owners = []
        def action():
            owners.append(self.session._owner)
            self.backend.read_squid()
        ok, error = self.run_command_thread(action)
        self.assertTrue(ok, error)
        self.assertNotEqual(owners, [caller])
        self.assertEqual(self.store.pending()['stage']['status'], 'verified')
        self.assertEqual(self.store.pending()['stage']['sample_id'], 'S1')
        self.assertTrue(self.vacuum.is_valve_connected())
        self.assertIsNone(self.backend._halt_check)

    def test_measurement_worker_holds_claim_for_acquisition_and_artifact_publication(self):
        worker = MeasurementWorker(SpecimenMeta('S1'), [], self.path.parent / 'out', backend=self.backend,
                                   run_id='run-1')
        owners, finished = [], []
        def body():
            owners.append(self.session._owner)
            self.backend.read_squid()
            worker._finish_run(aborted=False)
        def index(aborted):
            self.session.child_store._owned()
            owners.append(self.session._owner)
        worker.run_finished.connect(lambda aborted: finished.append((aborted, self.session._owner)),
                                    QtCore.Qt.DirectConnection)
        with patch.object(worker, '_run_owned', body), patch.object(worker, '_write_artifact_index', index):
            worker.start()
            self.assertTrue(worker.wait(10000))
        self.assertEqual(len(owners), 2)
        self.assertEqual(owners[0], owners[1])
        self.assertNotEqual(owners[0], threading.get_ident())
        self.assertEqual(len(finished), 1)
        self.assertIsNone(finished[0][1])
        self.assertEqual(self.store.pending()['stage']['status'], 'verified')
        self.assertIsNone(self.backend._halt_check)

    def test_competing_worker_does_not_clear_current_halt_hook_or_touch_instruments(self):
        owned = threading.Event()
        release = threading.Event()
        def claim():
            with self.session.claim():
                self.backend._halt_check = sentinel
                owned.set()
                release.wait(10)
        sentinel = lambda: False
        owner = threading.Thread(target=claim)
        owner.start()
        self.assertTrue(owned.wait(5))
        try:
            worker = QueueCommandWorker(self.backend, self.backend.read_squid)
            worker.run()
            self.assertFalse(worker.ok)
            self.assertIn('Another queue worker', worker.error)
            self.assertIs(self.backend._halt_check, sentinel)
            self.assertIsNone(self.squid.last_acquisition)
        finally:
            release.set()
            owner.join(5)
        self.assertFalse(owner.is_alive())

    def test_acquisition_failure_keeps_grip_and_pending_stage_after_worker_exit(self):
        with patch.object(self.squid, 'read_squid', side_effect=RuntimeError('SQUID failure')):
            ok, error = self.run_command_thread(self.backend.read_squid, recover_on_error=True)
        self.assertFalse(ok)
        self.assertIn('SQUID failure', error)
        self.assertIn('generic return cannot replay', error)
        self.assertEqual(self.store.pending()['stage']['status'], 'pending')
        self.assertTrue(self.vacuum.is_valve_connected())

    def test_missing_original_coordinator_blocks_generic_unstaged_motor_return(self):
        with self.backend.queue_worker_claim(), self.assertRaisesRegex(Exception, 'original native queue coordinator'):
            self.backend.return_to_safe_state()
        self.assertTrue(self.vacuum.is_valve_connected())

    def test_wrong_borrowed_store_and_released_session_fail_before_acquisition(self):
        self.backend._safety_store = self.store
        worker = QueueCommandWorker(self.backend, self.backend.read_squid)
        worker.run()
        self.assertFalse(worker.ok)
        self.assertIn('original borrowed', worker.error)
        self.backend._safety_store = self.session.child_store
        self.session.release()
        worker = QueueCommandWorker(self.backend, self.backend.read_squid)
        worker.run()
        self.assertFalse(worker.ok)
        self.assertIn('already been released', worker.error)
        self.assertIsNone(self.squid.last_acquisition)

    def test_hook_cleanup_failure_records_aborted_artifacts_under_original_claim(self):
        worker = MeasurementWorker(SpecimenMeta('S1'), [], self.path.parent / 'out', backend=self.backend,
                                   run_id='run-1')
        owners, finished = [], []
        def body():
            self.backend.read_squid()
            worker._finish_run(aborted=False)
        def index(aborted):
            self.session.child_store._owned()
            owners.append((aborted, self.session._owner))
        native = self.backend.set_halt_check
        def failed_clear(check):
            native(check)
            if check is None: raise RuntimeError('clear failed')
        worker.run_finished.connect(lambda aborted: finished.append((aborted, self.session._owner)),
                                    QtCore.Qt.DirectConnection)
        with patch.object(worker, '_run_owned', body), patch.object(worker, '_write_artifact_index', index), \
                patch.object(self.backend, 'set_halt_check', failed_clear):
            worker.start()
            self.assertTrue(worker.wait(10000))
        self.assertEqual([aborted for aborted, owner in owners], [False, True])
        self.assertEqual(owners[0][1], owners[1][1])
        self.assertEqual(finished, [(True, None)])
