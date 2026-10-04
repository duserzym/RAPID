"""Published and aborted runs preserve original queue scientific source evidence."""
import hashlib
import json
from pathlib import Path
import tempfile
import unittest

from rapid_main.data_model import SampleIndexRegistration, SampleIndexRegistrations, SpecimenMeta
from rapid_main.hardware_contracts import NoCommBackend
from rapid_main.io.specimen_writer import write_header
from rapid_main.measurement_worker import MeasurementWorker
from rapid_main.specimen_metadata import capture_specimen_metadata, restore_specimen_metadata


class QueueSourceArtifactTests(unittest.TestCase):
    def setUp(self):
        self.directory = tempfile.TemporaryDirectory()
        self.addCleanup(self.directory.cleanup)
        self.root = Path(self.directory.name)
        index = self.root / 'index.sam'
        index.write_text('A 12 Unit Site\n', encoding='latin-1')
        write_header(self.root / 'A', SpecimenMeta('A', comment='Original', volume=8.2, core_plate_strike=123))
        snapshot = capture_specimen_metadata('A', sample_dir=self.root,
            registrations=SampleIndexRegistrations([SampleIndexRegistration('A', formation='Unit', location='Site')]))
        self.meta = restore_specimen_metadata(snapshot).meta
        self.provenance = dict(schema='rapidpy.queue_specimen_source.v1', file_id=str(index),
            source_file=str(index), index_sha256=hashlib.sha256(index.read_bytes()).hexdigest(),
            row_id='a' * 32, specimen=snapshot)

    def worker(self):
        return MeasurementWorker(self.meta, ['NRM'], self.root / 'output', backend=NoCommBackend(),
            specimen_provenance=self.provenance)

    def test_published_bundle_and_index_preserve_detached_source_and_header_digests(self):
        worker = self.worker()
        original = json.loads(json.dumps(self.provenance))
        self.provenance['specimen']['meta']['volume'] = 999
        self.meta.volume = 999
        errors, finished = [], []
        worker.error_occurred.connect(errors.append)
        worker.run_finished.connect(finished.append)
        worker.run()
        self.assertEqual(errors, [])
        self.assertEqual(finished, [False])
        published = worker._publish_dir
        self.assertIn('SIMULATED', published.parts)
        provenance_file = published / 'provenance.json'
        payload = json.loads(provenance_file.read_text(encoding='utf-8'))
        self.assertEqual(payload['specimen_source'], original)
        self.assertEqual(payload['volume_cm3'], 8.2)
        manifest = json.loads((published / 'artifact_index.json').read_text(encoding='utf-8'))
        self.assertEqual(manifest['specimen_source'], original)
        entry = next(item for item in manifest['artifacts'] if Path(item['relative_path']).name == 'provenance.json')
        self.assertEqual((published / entry['relative_path']).resolve(), provenance_file.resolve())
        self.assertEqual(entry['sha256'], hashlib.sha256(provenance_file.read_bytes()).hexdigest())

    def test_aborted_run_preserves_source_evidence_without_publishing_scientific_bundle(self):
        worker = self.worker()
        worker.halt()
        worker.run()
        manifest = json.loads((worker._publish_dir / 'artifact_index.json').read_text(encoding='utf-8'))
        self.assertTrue(manifest['aborted'])
        self.assertEqual(manifest['specimen_source'], self.provenance)
        self.assertFalse((worker._publish_dir / 'provenance.json').exists())

    def test_mismatched_source_metadata_is_rejected_before_any_backend_work(self):
        self.provenance['specimen']['meta']['volume'] = 999
        with self.assertRaisesRegex(ValueError, 'original resolved metadata'):
            self.worker()
