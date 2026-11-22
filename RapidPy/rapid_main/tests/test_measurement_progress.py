import hashlib
import json
from pathlib import Path
import tempfile
import unittest
from unittest.mock import patch

from rapid_main.acquisition_evidence import canonical
from rapid_main.data_model import SpecimenMeta
from rapid_main.hardware_contracts import PreflightResult
from rapid_main.io.legacy_up import read_up_measurements
from rapid_main.magnetometer import BlockAudit, BracketedMeasurementBlock, reduce_bracketed_measurement
from rapid_main.measurement_progress import MeasurementProgress
from rapid_main.measurement_worker import MeasurementWorker


def block(identity='one', scale=1.):
    return BracketedMeasurementBlock((0., 0., 0.), ((2.*scale, 0., scale),)*4, (0., 0., 0.),
        audit=BlockAudit(block_id=identity, sample_name='A', run_id='run', treatment_label='NRM'))


def capture(store, index=0, blocks=None):
    blocks = blocks or (block(), block('two', 2.))
    store.start_step(index)
    store.capture_step(index, blocks=blocks,
        readings=[reduce_bracketed_measurement(item).moment_emu for item in blocks])


class Backend:
    simulated = True

    def __init__(self, readings=None, *, return_fail=False, treatment_fail=False):
        self.calls = []
        self.readings = iter(readings or (block(),))
        self.return_fail = return_fail
        self.treatment_fail = treatment_fail

    def preflight(self):
        self.calls.append('preflight')
        return PreflightResult.pass_ok()

    def is_available(self):
        return True

    def set_demag_step(self, label):
        self.calls.append('treat')
        if self.treatment_fail:
            raise RuntimeError('uncertain treatment')

    def read_squid(self):
        self.calls.append('read')
        return next(self.readings)

    def return_to_safe_state(self):
        self.calls.append('return')
        if self.return_fail:
            raise RuntimeError('uncertain support')


class MeasurementProgressTests(unittest.TestCase):
    def store(self, directory, **kwargs):
        return MeasurementProgress(directory, run_id='run', specimen='A', labels=['NRM'], **kwargs)

    def test_full_raw_collection_and_legacy_export_survive_reopen(self):
        with tempfile.TemporaryDirectory() as directory:
            store = self.store(directory, source={'file_id':'source-one', 'row_id':'row'})
            store.begin()
            capture(store)
            with self.assertRaises(ValueError):
                store.accepted_steps()
            store.accept()
            restored = self.store(directory, source={'file_id':'source-one', 'row_id':'row'})
            payload, = restored.accepted_steps()
            self.assertEqual(payload['block_statistics']['position_count'], 8)
            self.assertEqual([raw['audit']['block_id'] for raw in payload['blocks']], ['one', 'two'])
            self.assertEqual(len(read_up_measurements((store.directory/'00000.UP').read_text(), 'A').blocks), 2)
            self.assertEqual(restored.reference, store.reference)
            with self.assertRaises(ValueError):
                restored.start_step(0)

    def test_distinct_sources_and_runs_never_replace_raw_acquisitions(self):
        with tempfile.TemporaryDirectory() as directory:
            first = self.store(directory, source={'file_id':'one'})
            second = self.store(directory, source={'file_id':'two'})
            first.begin()
            capture(first)
            original = (first.directory/'00000-captured.json').read_bytes()
            second.begin()
            capture(second, blocks=(block('different', 3.),))
            self.assertNotEqual(first.directory, second.directory)
            self.assertEqual((first.directory/'00000-captured.json').read_bytes(), original)
            with self.assertRaises(FileExistsError):
                self.store(directory, source={'file_id':'one'}).begin()

    def test_torn_intent_and_missing_capture_are_never_accepted_or_overwritten(self):
        with tempfile.TemporaryDirectory() as directory:
            store = self.store(directory)
            store.begin()
            store.start_step(0)
            with self.assertRaises(ValueError):
                store.accept()
            with self.assertRaises(FileExistsError):
                capture(store)
            (store.directory/'00000-started.json').write_text('{')
            with self.assertRaises(ValueError):
                store.capture_step(0, blocks=(), readings=[(1., 2., 3.)])

    def test_modified_and_rehashed_inconsistent_raw_data_are_rejected(self):
        with tempfile.TemporaryDirectory() as directory:
            store = self.store(directory)
            store.begin()
            capture(store)
            store.accept()
            path = store.directory/'00000-captured.json'
            original = path.read_bytes()
            packet = json.loads(original)
            packet['payload']['blocks'][0]['positions'][0][0] += 1
            path.write_text(canonical(packet))
            with self.assertRaisesRegex(ValueError, 'digest'):
                store.accepted_steps()
            packet['sha256'] = hashlib.sha256(canonical(packet['payload']).encode()).hexdigest()
            path.write_text(canonical(packet))
            with self.assertRaisesRegex(ValueError, 'disagree'):
                store.accepted_steps()

    def test_legacy_export_change_blocks_acceptance(self):
        with tempfile.TemporaryDirectory() as directory:
            store = self.store(directory)
            store.begin()
            capture(store)
            (store.directory/'00000.UP').write_text('changed')
            with self.assertRaisesRegex(ValueError, 'UP export'):
                store.accept()

    def test_counter_dvm_observations_and_commands_are_losslessly_retained(self):
        from test_acquisition import _build_service
        service, *_ = _build_service()
        raw = service.acquire().block
        with tempfile.TemporaryDirectory() as directory:
            store = self.store(directory)
            store.begin()
            capture(store, blocks=(raw,))
            store.accept()
            payload, = store.accepted_steps()
            from rapid_main.acquisition_evidence import restore_block
            restored = restore_block(payload['blocks'][0])
            self.assertEqual(len(restored.observations), 6)
            self.assertEqual(restored, raw)

    def test_unstructured_readings_are_retained_without_fabricated_blocks(self):
        with tempfile.TemporaryDirectory() as directory:
            store = self.store(directory)
            store.begin()
            store.start_step(0)
            store.capture_step(0, blocks=(), readings=[(1., 2., 3.)])
            store.accept()
            payload, = store.accepted_steps()
            self.assertEqual(payload['kind'], 'unstructured')
            self.assertEqual(payload['cycle_statistics']['signal_to_drift'], '+infinity')
            self.assertEqual(payload['blocks'], [])
            self.assertIsNone(payload['legacy_up'])
            self.assertFalse((store.directory/'00000.UP').exists())

    def test_repeated_ids_and_mutated_aggregates_cannot_be_sealed(self):
        with tempfile.TemporaryDirectory() as directory:
            store = self.store(directory)
            store.begin()
            with self.assertRaisesRegex(ValueError, 'repeats'):
                capture(store, blocks=(block(), block()))
            self.assertFalse((store.directory/'accepted.json').exists())

    def test_writer_persists_intent_before_treatment_and_keeps_failed_attempt(self):
        with tempfile.TemporaryDirectory() as directory:
            backend = Backend(treatment_fail=True)
            worker = MeasurementWorker(SpecimenMeta(name='A'), ['NRM'], Path(directory), backend, run_id='run')
            original = backend.set_demag_step
            def treat(label):
                self.assertTrue((worker._progress.directory/'00000-started.json').exists())
                original(label)
            backend.set_demag_step = treat
            worker.run()
            self.assertEqual(backend.calls, ['preflight','treat','return'])
            self.assertFalse((worker._progress.directory/'accepted.json').exists())
            self.assertFalse((worker._progress.directory/'00000-captured.json').exists())

    def test_safe_return_failure_preserves_previous_bundle_and_unaccepted_raw_data(self):
        with tempfile.TemporaryDirectory() as directory:
            output = Path(directory)
            good = MeasurementWorker(SpecimenMeta(name='A'), ['NRM'], output, Backend(), run_id='first')
            good.run()
            published = output/'SIMULATED'/'A'
            original = published.read_bytes()
            bad = MeasurementWorker(SpecimenMeta(name='A'), ['NRM'], output,
                Backend((block('later',3.),), return_fail=True), run_id='second')
            bad.run()
            self.assertEqual(published.read_bytes(), original)
            self.assertTrue((bad._progress.directory/'00000-captured.json').exists())
            with self.assertRaises(ValueError):
                bad._progress.accepted_steps()
            self.assertEqual(len(good._progress.accepted_steps()), 1)

    def test_resume_and_duplicate_run_are_rejected_before_backend_io(self):
        with tempfile.TemporaryDirectory() as directory:
            output = Path(directory)
            first = MeasurementWorker(SpecimenMeta(name='A'), ['NRM'], output, Backend(), run_id='same')
            first.run()
            for resume in (False, True):
                backend = Backend()
                worker = MeasurementWorker(SpecimenMeta(name='A'), ['NRM'], output, backend,
                    run_id='same', resume=resume)
                worker.run()
                self.assertEqual(backend.calls, [])

    def test_failed_progress_write_prevents_treatment_io(self):
        with tempfile.TemporaryDirectory() as directory:
            backend = Backend()
            worker = MeasurementWorker(SpecimenMeta(name='A'), ['NRM'], Path(directory), backend)
            with patch.object(MeasurementProgress, 'start_step', side_effect=OSError('disk full')):
                worker.run()
            self.assertEqual(backend.calls, ['preflight','return'])


if __name__ == '__main__':
    unittest.main()
