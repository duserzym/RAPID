"""Queue lifetime and subordinate treatment recovery share one durable owner."""
import copy
import json
from pathlib import Path
import tempfile
import threading
import subprocess
import sys
from types import SimpleNamespace
import unittest
from unittest.mock import patch

from rapidpy_common.hardware_safety import HardwareSafetyError, HardwareSafetyStore
from rapidpy_common.queue_safety import QueueSafetyStore, QueueWorkflowSession, QueueSafeStateRecord, QueueVacuumHoldRecord


def record(safe=True, simulated=False, **extra):
    payload = dict(safe_state_confirmed=safe, simulated=simulated, **extra)
    return SimpleNamespace(safe_state_confirmed=safe, simulated=simulated, to_dict=lambda: payload)


def queue_record(safe=True, simulated=False):
    return QueueSafeStateRecord(safe, safe, safe, simulated=simulated)


class QueueSafetyTests(unittest.TestCase):
    def setUp(self):
        self.directory = tempfile.TemporaryDirectory()
        self.addCleanup(self.directory.cleanup)
        self.path = Path(self.directory.name) / 'safety.json'
        self.store = QueueSafetyStore(self.path)
        self.binding = {'motor': {'port': 'COM5'}, 'field': {'board': 1}}
        self.profile = {'stage_profiles': {name: copy.deepcopy(self.binding)
            for name in ('af', 'arm', 'pulse', 'rrm', 'acquisition', 'motion', 'vacuum')},
            'vacuum': {'port': 'COM11', 'baud': 9600}}
        self.session = QueueWorkflowSession.start(self.store, {'commands': ['Holder', 'AF50', 'Return']}, self.profile, run_id='run')
        self.addCleanup(self.session.release)

    def stage(self, family='af', safe=True, simulated=False):
        with self.session.claim() as child:
            with child.operation_lease():
                token = child.begin(family, {'peak': 5}, self.binding, sample_id='specimen', run_id='run')
                child.finish(token, self.binding, record(safe, simulated))
        return token

    def test_stage_success_leaves_root_pending_and_preserves_original_plan(self):
        token = self.stage()
        root = HardwareSafetyStore(self.path).pending()
        self.assertEqual(root['family'], 'queue')
        self.assertEqual(root['token'], self.session.token)
        self.assertEqual(root['stage']['token'], token)
        self.assertEqual(root['stage']['status'], 'verified')
        self.assertEqual(root['plan']['commands'], ['Holder', 'AF50', 'Return'])
        with self.assertRaises(HardwareSafetyError):
            HardwareSafetyStore(self.path).finish(self.session.token, self.profile, record())

    def test_all_treatment_and_acquisition_families_share_one_lifetime_lease(self):
        for family in self.profile['stage_profiles']:
            self.stage(family)
        state = self.store.pending(self.profile)
        self.assertEqual(state['stage_count'], 7)
        self.assertEqual(self.store.verify_history(state), 7)
        with self.assertRaises(HardwareSafetyError):
            with HardwareSafetyStore(self.path).operation_lease():
                self.fail('Second hardware owner acquired the queue lease.')
        self.store.finish_queue(self.session.token, self.profile, queue_record())
        self.assertIsNone(self.store.pending())
        # Verified journal alone does not drop an OS lease before worker exit.
        with self.assertRaises(HardwareSafetyError):
            with HardwareSafetyStore(self.path).operation_lease():
                self.fail('Lease released before session owner settled.')
        self.session.release()
        with HardwareSafetyStore(self.path).operation_lease():
            pass

    def test_recovery_reuses_exact_pending_stage_without_replaying_begin(self):
        token = self.stage('pulse', safe=False)
        self.session.release()
        recovered = QueueWorkflowSession.recover(QueueSafetyStore(self.path), self.profile)
        self.addCleanup(recovered.release)
        with recovered.claim() as child:
            pending = child.pending(self.binding)
            self.assertEqual(pending['token'], token)
            self.assertEqual(pending['plan'], {'peak': 5})
            self.assertEqual(pending['family'], 'pulse')
            self.assertEqual(pending['sample_id'], 'specimen')
            with self.assertRaises(HardwareSafetyError):
                child.begin('af', {}, self.binding)
            child.finish(token, self.binding, record())
            self.assertIsNone(child.pending())
        root = recovered.store.pending()
        self.assertEqual(root['stage_count'], 1)
        self.assertEqual(recovered.store.verify_history(root), 2)
        recovered.store.finish_queue(recovered.token, self.profile, queue_record())
        self.assertIsNone(recovered.store.pending())

    def test_successful_treatment_cannot_clear_an_unverified_queue_release(self):
        self.stage()
        self.store.finish_queue(self.session.token, self.profile, queue_record(False))
        self.assertEqual(self.store.pending()['family'], 'queue')
        self.store.finish_queue(self.session.token, self.profile, queue_record(True, True))
        self.assertIsNotNone(self.store.pending())

    def test_pending_or_simulated_stage_cannot_finish_parent_or_start_next_stage(self):
        for simulated in (False, True):
            if simulated:
                with self.session.claim() as child:
                    prior = child.pending()
                    child.finish(prior['token'], self.binding, record(True, True))
            else:
                self.stage(safe=False)
            with self.assertRaises(HardwareSafetyError):
                self.store.finish_queue(self.session.token, self.profile, queue_record())
            with self.session.claim() as child:
                with self.assertRaises(HardwareSafetyError):
                    child.begin('motion', {}, self.binding)

    def test_changed_station_binding_is_rejected_before_stage_begin(self):
        original = self.path.read_bytes()
        changed = copy.deepcopy(self.binding)
        changed['motor']['port'] = 'COM99'
        with self.session.claim() as child:
            with self.assertRaises(HardwareSafetyError):
                child.begin('af', {}, changed)
        self.assertEqual(self.path.read_bytes(), original)
        self.stage()
        with self.assertRaises(HardwareSafetyError):
            self.store.finish_queue(self.session.token, dict(self.profile, vacuum={'port': 'COM99'}), queue_record())
        self.assertIsNotNone(self.store.pending())

    def test_changed_recovery_profile_releases_failed_attempt_lease(self):
        self.session.release()
        with self.assertRaises(HardwareSafetyError):
            QueueWorkflowSession.recover(self.store, dict(self.profile, vacuum={'port': 'COM99'}))
        recovered = QueueWorkflowSession.recover(self.store, self.profile)
        recovered.release()

    def test_wrong_stage_token_and_profile_cannot_publish_recovery(self):
        token = self.stage(safe=False)
        with self.session.claim() as child:
            for wrong_token, wrong_profile in [('a' * 32, self.binding), (token, {'different': True})]:
                with self.assertRaises(HardwareSafetyError):
                    child.finish(wrong_token, wrong_profile, record())
        self.assertEqual(self.store.pending()['stage']['status'], 'pending')

    def test_stage_adapter_cannot_borrow_without_current_worker_claim(self):
        child = self.session.child_store
        self.assertEqual(child.pending()['family'], 'queue')
        with self.assertRaises(HardwareSafetyError):
            child.begin('af', {}, self.binding)
        with self.assertRaises(HardwareSafetyError):
            with child.operation_lease():
                pass
        self.session.release()
        with self.assertRaises(HardwareSafetyError):
            with self.session.claim():
                pass

    def test_concurrent_workers_cannot_borrow_or_release_live_owner(self):
        entered, release = threading.Event(), threading.Event()
        def owner():
            with self.session.claim():
                entered.set()
                release.wait(5)
        thread = threading.Thread(target=owner)
        thread.start()
        try:
            self.assertTrue(entered.wait(2))
            with self.assertRaises(HardwareSafetyError):
                with self.session.claim():
                    pass
            with self.assertRaises(HardwareSafetyError):
                self.session.release()
            with self.assertRaises(HardwareSafetyError):
                self.session.child_store.begin('af', {}, self.binding)
        finally:
            release.set()
            thread.join(5)
        self.assertFalse(thread.is_alive())
        with self.session.claim():
            with self.assertRaises(HardwareSafetyError):
                self.session.release()

    def test_stage_publication_failure_retains_pending_identity(self):
        with self.session.claim() as child:
            token = child.begin('af', {}, self.binding)
            with patch.object(self.store, '_publish_event', side_effect=OSError('disk full')):
                with self.assertRaises(OSError):
                    child.finish(token, self.binding, record())
        pending = HardwareSafetyStore(self.path).pending()
        self.assertEqual(pending['stage']['token'], token)
        self.assertEqual(pending['stage']['status'], 'pending')

    def test_journal_write_failure_after_event_publication_never_clears_stage(self):
        with self.session.claim() as child:
            token = child.begin('af', {}, self.binding)
            with patch.object(self.store, '_write', side_effect=HardwareSafetyError('journal full')):
                with self.assertRaises(HardwareSafetyError):
                    child.finish(token, self.binding, record())
            self.assertEqual(child.pending()['token'], token)
            child.finish(token, self.binding, record())
        self.assertEqual(len(list(self.path.parent.rglob('*.json'))), 3)
        self.assertEqual(self.store.verify_history(self.store.pending()), 1)

    def test_missing_or_tampered_linked_evidence_prevents_queue_clear(self):
        self.stage()
        state = self.store.pending()
        path = self.store._event_path(state['token'], state['history_head']['id'])
        payload = path.read_bytes()
        for corrupt in (b'{}', payload + b'x'):
            path.write_bytes(corrupt)
            with self.assertRaises(HardwareSafetyError):
                self.store.finish_queue(self.session.token, self.profile, queue_record())
            self.assertIsNotNone(self.store.pending())
        path.unlink()
        with self.assertRaises(HardwareSafetyError):
            self.store.finish_queue(self.session.token, self.profile, queue_record())

    def test_older_event_corruption_prevents_recovery_and_drops_attempt_lease(self):
        self.stage()
        first = self.store.pending()['history_head']
        self.stage('motion')
        path = self.store._event_path(self.session.token, first['id'])
        path.write_bytes(b'corrupt')
        self.session.release()
        with self.assertRaises(HardwareSafetyError):
            QueueWorkflowSession.recover(self.store, self.profile)
        with self.store.operation_lease():
            self.assertIsNotNone(self.store.pending())

    def test_strict_record_flags_cannot_clear_queue_or_stage(self):
        with self.session.claim() as child:
            token = child.begin('af', {}, self.binding)
            for bad in (record('true'), record(True, 'false')):
                with self.assertRaises(HardwareSafetyError):
                    child.finish(token, self.binding, bad)
            contradictory = record()
            contradictory.to_dict = lambda: dict(safe_state_confirmed=False, simulated=False)
            with self.assertRaises(HardwareSafetyError):
                child.finish(token, self.binding, contradictory)
            child.finish(token, self.binding, record())
        with self.assertRaises(HardwareSafetyError):
            self.store.finish_queue(self.session.token, self.profile, queue_record('true'))

    def test_stage_plan_is_detached_and_immutable_events_remain_separate(self):
        plan = {'ramp': {'peak': 5}}
        with self.session.claim() as child:
            token = child.begin('af', plan, self.binding)
            plan['ramp']['peak'] = 900
            child.finish(token, self.binding, record())
        state = self.store.pending()
        self.assertEqual(state['stage']['plan']['ramp']['peak'], 5)
        path = self.store._event_path(state['token'], state['history_head']['id'])
        original = path.read_bytes()
        self.stage('motion')
        self.assertEqual(path.read_bytes(), original)
        self.assertEqual(self.store.verify_history(self.store.pending()), 2)

    def test_shared_reader_rejects_corrupt_queue_extensions_and_unsafe_clear(self):
        self.stage(safe=False)
        original = self.store.pending()
        for mutation in (
            lambda state: state.update(status='verified', record={'safe_state_confirmed': True, 'simulated': False}),
            lambda state: state.update(stage_count=-1),
            lambda state: state['stage'].update(family='unknown'),
            lambda state: state['stage'].update(profile={'port': 'COM99'}),
            lambda state: state['history_head'].update(id='../escape'),
        ):
            state = copy.deepcopy(original)
            mutation(state)
            self.store._write(state)
            with self.assertRaises(HardwareSafetyError):
                HardwareSafetyStore(self.path).pending()
        self.store._write(original)

    def test_begin_failure_does_not_leak_lifetime_lease(self):
        self.session.release()
        with patch.object(self.store, 'begin_queue', side_effect=HardwareSafetyError('cannot persist')):
            with self.assertRaises(HardwareSafetyError):
                QueueWorkflowSession.start(self.store, {}, self.profile)
        with self.store.operation_lease():
            pass

    def test_existing_foreign_latch_cannot_be_replaced_by_queue(self):
        self.session.release()
        with self.assertRaises(HardwareSafetyError):
            QueueWorkflowSession.start(self.store, {}, self.profile)
        self.assertEqual(self.store.pending()['token'], self.session.token)

    def test_acknowledged_vacuum_hold_can_continue_without_claiming_outputs_off(self):
        with self.session.claim() as child:
            token = child.begin('vacuum', {'action': 'enable'}, self.binding)
            child.finish(token, self.binding, QueueVacuumHoldRecord(True, ('VALVE ON: ACK', 'PUMP ON: ACK')))
            self.assertIsNone(child.pending())
        root = HardwareSafetyStore(self.path).pending()
        self.assertEqual(root['stage']['status'], 'held')
        self.assertIs(root['stage']['record']['safe_state_confirmed'], False)
        with self.assertRaises(HardwareSafetyError):
            self.store.finish_queue(self.session.token, self.profile, queue_record())
        self.stage('af')
        self.assertEqual(self.store.verify_history(self.store.pending()), 2)
        self.assertIsNotNone(self.store.pending())

    def test_missing_or_simulated_vacuum_acknowledgements_cannot_authorize_next_stage(self):
        with self.session.claim() as child:
            token = child.begin('vacuum', {'action': 'enable'}, self.binding)
            for bad in (QueueVacuumHoldRecord(True, ()), QueueVacuumHoldRecord(True, ('ACK',), simulated=True),
                        QueueVacuumHoldRecord(False, ('ERROR',))):
                child.finish(token, self.binding, bad)
                self.assertEqual(child.pending()['status'], 'pending')
                with self.assertRaises(HardwareSafetyError):
                    child.begin('motion', {}, self.binding)

    def test_treatment_success_is_not_a_queue_release_record(self):
        self.stage()
        with self.assertRaises(HardwareSafetyError):
            self.store.finish_queue(self.session.token, self.profile, record())
        self.assertIsNotNone(self.store.pending())

    def test_each_independent_output_verification_is_required(self):
        self.stage()
        for checks in ((False, True, True), (True, False, True), (True, True, False)):
            self.store.finish_queue(self.session.token, self.profile, QueueSafeStateRecord(*checks))
            self.assertIsNotNone(self.store.pending())
        self.store.finish_queue(self.session.token, self.profile, QueueSafeStateRecord(True, True, True, ('ack failed',)))
        self.assertIsNotNone(self.store.pending())

    def test_native_treatment_hooks_borrow_the_outer_lease_without_clearing_it(self):
        from dataclasses import dataclass
        from rapid_main.hardware_contracts import QueueHardwareBackend
        from rapid_main.config import AppConfig
        @dataclass
        class Plan:
            peak: float = 5
        backend = object.__new__(QueueHardwareBackend)
        backend._config = AppConfig()
        backend._sample_name, backend._run_id = 'specimen', 'run'
        native_binding = backend._safety_profile()
        # This is a separate initial station snapshot, not a mutation while owned.
        self.store.finish_queue(self.session.token, self.profile, queue_record())
        self.session.release()
        profile = dict(self.profile, stage_profiles={'af': native_binding})
        native = QueueWorkflowSession.start(self.store, {}, profile)
        self.addCleanup(native.release)
        backend._safety_store = native.child_store
        self.assertEqual(backend._durable_fault_family(), 'queue')
        with native.claim():
            self.assertEqual(backend._durable_fault_family(), '')
            with backend._safety_operation('af', Plan()) as token:
                backend._finish_safety_operation(token, record())
            self.assertEqual(backend._durable_fault_family(), '')
        self.assertEqual(HardwareSafetyStore(self.path).pending()['family'], 'queue')

    def test_held_vacuum_cannot_disguise_a_failed_release(self):
        with self.session.claim() as child:
            token = child.begin('vacuum', {'action': 'release'}, self.binding)
            child.finish(token, self.binding, QueueVacuumHoldRecord(True, ('still ON',)))
            self.assertEqual(child.pending()['status'], 'pending')
            with self.assertRaises(HardwareSafetyError):
                child.begin('motion', {}, self.binding)

    def test_corrupt_latest_event_blocks_next_native_stage_before_begin(self):
        self.stage()
        state = self.store.pending()
        self.store._event_path(state['token'], state['history_head']['id']).write_bytes(b'corrupt')
        with self.session.claim() as child:
            with self.assertRaises(HardwareSafetyError):
                child.begin('motion', {}, self.binding)
        self.assertEqual(self.store.pending()['stage_count'], 1)

    def test_native_treatment_recovery_cannot_clear_or_replay_outer_queue(self):
        from rapid_main.hardware_contracts import QueueHardwareBackend
        from rapid_main.config import AppConfig
        backend = object.__new__(QueueHardwareBackend)
        backend._config = AppConfig()
        backend._safety_store = HardwareSafetyStore(self.path)
        with self.assertRaisesRegex(RuntimeError, 'queue coordinator'):
            backend.return_to_safe_state()
        self.assertEqual(self.store.pending()['token'], self.session.token)

    def test_process_crash_drops_os_lease_but_retains_original_stage(self):
        self.store.finish_queue(self.session.token, self.profile, queue_record())
        self.session.release()
        profile_path = self.path.parent / 'profile.json'
        profile_path.write_text(json.dumps(self.profile), encoding='utf-8')
        script = '''
import json, os, sys
from pathlib import Path
sys.path.insert(0, sys.argv[2])
from rapidpy_common.queue_safety import QueueSafetyStore, QueueWorkflowSession
profile = json.loads(Path(sys.argv[3]).read_text(encoding='utf-8'))
session = QueueWorkflowSession.start(QueueSafetyStore(sys.argv[1]), {'command': 'Goto'}, profile)
with session.claim() as child:
    child.begin('motion', {'target': 700}, profile['stage_profiles']['motion'], sample_id='specimen')
    os._exit(17)
'''
        rapidpy = Path(__file__).resolve().parents[2]
        result = subprocess.run([sys.executable, '-c', script, str(self.path), str(rapidpy), str(profile_path)],
            capture_output=True, timeout=15)
        self.assertEqual(result.returncode, 17, result.stderr.decode(errors='replace'))
        recovered = QueueWorkflowSession.recover(self.store, self.profile)
        self.addCleanup(recovered.release)
        with recovered.claim() as child:
            pending = child.pending(self.binding)
            self.assertEqual(pending['family'], 'motion')
            self.assertEqual(pending['plan'], {'target': 700})
            self.assertEqual(pending['sample_id'], 'specimen')
