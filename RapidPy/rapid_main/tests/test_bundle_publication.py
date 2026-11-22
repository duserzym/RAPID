"""Fault injection at the public filenames and the final commit point."""

import hashlib
import json
import os
from pathlib import Path
import shutil
import subprocess
import sys
import tempfile
import unittest
from unittest.mock import patch
from zipfile import ZipFile

from data_viewer.data_loading import load_magic_directory
from rapid_main.data_model import SpecimenMeta
from rapid_main.hardware_contracts import NoCommBackend
from rapid_main.io import bundle_history, bundle_publication, measurement_bundle
from rapid_main.io.measurement_bundle import MeasurementBundleWriter, recover_measurement_output
from rapid_main.measurement_worker import MeasurementWorker
from rapid_main.quicklook import write_measurement_quicklook
from rapidpy_common.hardware_safety import HardwareSafetyError
from tests.test_queue_and_bundle import _step


class BundlePublicationTests(unittest.TestCase):
    def setUp(self):
        temp = tempfile.TemporaryDirectory()
        self.addCleanup(temp.cleanup)
        self.root = Path(temp.name)

    def original(self):
        with MeasurementBundleWriter(self.root, SpecimenMeta('A')) as writer:
            writer.append_step(_step('NRM'))
        return self.snapshot()

    def snapshot(self):
        names = ('A', 'A.rmg', 'measurements.txt', 'specimens.txt', 'provenance.json', '.rapidpy-steps.json')
        return {name: (self.root / name).read_bytes() for name in names if (self.root / name).exists()}

    def crash_after_third_replacement(self):
        # A real process exit suppresses finally blocks and releases its OS lease.
        script = '''
from pathlib import Path
import os
import sys
from unittest.mock import patch
sys.path[:0] = sys.argv[2:4]
from rapid_main.io import bundle_publication
from rapid_main.io.measurement_bundle import MeasurementBundleWriter
from rapid_main.data_model import SpecimenMeta
from tests.test_queue_and_bundle import _step
root = Path(sys.argv[1])
replace = bundle_publication._replace_snapshot
def crash(source, target):
    replace(source, target)
    if target == root / 'measurements.txt':
        os._exit(73)
with patch.object(bundle_publication, '_replace_snapshot', crash):
    with MeasurementBundleWriter(root, SpecimenMeta('A')) as writer:
        writer.append_step(_step('AF20'))
'''
        rapidpy = Path(__file__).resolve().parents[2]
        result = subprocess.run([sys.executable, '-c', script, str(self.root), str(rapidpy), str(rapidpy / 'rapid_main')],
            capture_output=True, text=True, timeout=30)
        self.assertEqual(result.returncode, 73, result.stderr)

    def test_failed_replacement_restores_the_complete_previous_bundle(self):
        before = self.original()
        writer = MeasurementBundleWriter(self.root, SpecimenMeta('A'))
        writer.append_step(_step('AF20'))
        replace = bundle_publication._replace_snapshot

        def interruption(source, target):
            if target == self.root / 'measurements.txt':
                raise OSError('injected MagIC replacement failure')
            replace(source, target)

        with patch.object(bundle_publication, '_replace_snapshot', interruption):
            with self.assertRaisesRegex(OSError, 'MagIC replacement failure'):
                writer.commit()
        self.assertEqual(self.snapshot(), before)
        self.assertFalse((self.root / bundle_publication.JOURNAL).exists())

    def test_next_writer_rolls_back_an_interrupted_process(self):
        before = self.original()
        self.crash_after_third_replacement()
        self.assertNotEqual(self.snapshot(), before)
        writer = MeasurementBundleWriter(self.root, SpecimenMeta('A'), resume=True)
        self.addCleanup(writer.abort)
        self.assertEqual(self.snapshot(), before)
        self.assertFalse(writer.append_step(_step('NRM')))
        self.assertFalse((self.root / bundle_publication.JOURNAL).exists())
        archives = list((self.root / bundle_history.HISTORY_DIRECTORY).iterdir())
        self.assertEqual(len(archives), 1)
        with bundle_history.open_bundle_archive(archives[0]) as archive:
            for name, data in before.items():
                self.assertEqual(archive.read(name), data)

    def test_replacement_archives_the_old_science_and_acquisition_for_data_review(self):
        first = MeasurementWorker(SpecimenMeta('A'), ['NRM'], self.root, NoCommBackend(), run_id='original')
        first.run()
        output = self.root / 'SIMULATED'
        old_index = (output / 'artifact_index.json').read_bytes()
        old_rmg = (output / 'A.rmg').read_bytes()
        files = [entry for entry in json.loads(old_index)['artifacts'] if entry['exists']]
        second = MeasurementWorker(SpecimenMeta('A'), ['AF20'], self.root, NoCommBackend(), run_id='replacement')
        second.run()
        self.assertIs(second._pending_run_result, False)
        provenance = json.loads((output / 'provenance.json').read_bytes())
        ref = provenance['previous_publication']
        manifest = output / ref['relative_path']
        self.assertEqual(ref['sha256'], hashlib.sha256(manifest.read_bytes()).hexdigest())
        with bundle_history.open_bundle_archive(manifest.parent, expected_id=ref['archive_id']) as archive:
            self.assertEqual(archive.read('artifact_index.json'), old_index)
            self.assertEqual(archive.read('A.rmg'), old_rmg)
            for entry in files:
                self.assertEqual(hashlib.sha256(archive.read(entry['relative_path'])).hexdigest(),
                                 entry['sha256'], entry['name'])
            members = {name: archive.read(name) for name in archive.namelist()}
        previous = load_magic_directory(manifest.parent)
        self.assertEqual(previous.source_kind, 'measurement_archive')
        self.assertEqual([step.demag_label for step in previous.specimens[0].steps], ['NRM'])
        self.assertEqual(load_magic_directory(output).specimens[0].steps[0].demag_label, 'AF20')
        members['measurements.txt'] = b'changed archive'
        with ZipFile(manifest.parent / bundle_history.ARCHIVE_PAYLOAD, 'w') as archive:
            for name, data in members.items():
                archive.writestr(name, data)
        with self.assertRaisesRegex(ValueError, 'Archived measurement file changed'):
            load_magic_directory(manifest.parent)

    @unittest.skipUnless(os.name == 'nt', 'Windows legacy path limit')
    def test_archive_manifest_can_exceed_the_windows_legacy_path_limit(self):
        output = self.root / ('x' * (196 - len(str(self.root))))
        output.mkdir()
        writable = bundle_history._archive_io_path(output)
        self.addCleanup(shutil.rmtree, writable)
        for label in ('NRM', 'AF20'):
            with MeasurementBundleWriter(writable, SpecimenMeta('A')) as writer:
                writer.append_step(_step(label))
        previous = json.loads((output / 'provenance.json').read_bytes())['previous_publication']
        manifest = output / previous['relative_path']
        self.assertGreater(len(str(manifest)), 260)
        self.assertRegex(previous['sha256'], r'^[0-9a-f]{64}$')
        prior = load_magic_directory(manifest.parent)
        self.assertEqual([step.demag_label for step in prior.specimens[0].steps], ['NRM'])

    def test_archive_failure_keeps_the_previous_publication(self):
        before = self.original()
        writer = MeasurementBundleWriter(self.root, SpecimenMeta('A'))
        writer.append_step(_step('AF20'))
        with patch.object(ZipFile, 'write', side_effect=OSError('archive disk full')):
            with self.assertRaisesRegex(OSError, 'archive disk full'):
                writer.commit()
        self.assertEqual(self.snapshot(), before)
        self.assertFalse((self.root / bundle_publication.JOURNAL).exists())

    def test_edit_after_archive_is_preserved_and_blocks_scientific_replacement(self):
        before = self.original()
        writer = MeasurementBundleWriter(self.root, SpecimenMeta('A'))
        writer.append_step(_step('AF20'))
        capture = measurement_bundle.archive_previous_bundle

        def edit_after_capture(*args):
            result = capture(*args)
            (self.root / 'A.rmg').write_bytes(b'operator edit')
            return result

        with patch.object(measurement_bundle, 'archive_previous_bundle', side_effect=edit_after_capture):
            with self.assertRaisesRegex(ValueError, 'changed after its archive was written'):
                writer.commit()
        self.assertEqual(self.snapshot(), {**before, 'A.rmg': b'operator edit'})
        only_archive, = (self.root / bundle_history.HISTORY_DIRECTORY).iterdir()
        with bundle_history.open_bundle_archive(only_archive) as archived:
            self.assertEqual(archived.read('A.rmg'), before['A.rmg'])

    def test_conflicting_edit_retains_the_recovery_journal_and_all_current_bytes(self):
        self.original()
        self.crash_after_third_replacement()
        (self.root / 'A.rmg').write_bytes(b'Unrelated operator edit')
        current = self.snapshot()
        with self.assertRaisesRegex(ValueError, 'Output changed during measurement publication'):
            recover_measurement_output(self.root, 'A')
        self.assertEqual(self.snapshot(), current)
        self.assertTrue((self.root / bundle_publication.JOURNAL / 'manifest.json').is_file())

    def test_unrelated_files_in_the_journal_directory_are_preserved(self):
        journal = self.root / bundle_publication.JOURNAL
        journal.mkdir()
        other = journal / 'operator-data.txt'
        other.write_bytes(b'original')
        with self.assertRaisesRegex(ValueError, 'unexpected files'):
            recover_measurement_output(self.root, 'A')
        self.assertEqual(other.read_bytes(), b'original')

    def test_one_active_writer_cannot_reset_another_writers_staging(self):
        writer = MeasurementBundleWriter(self.root, SpecimenMeta('A'))
        self.addCleanup(writer.abort)
        writer.append_step(_step('NRM'))
        before = writer._staged.specimen_file.read_bytes()
        with self.assertRaises(HardwareSafetyError):
            MeasurementBundleWriter(self.root, SpecimenMeta('A'))
        self.assertEqual(writer._staged.specimen_file.read_bytes(), before)
        writer.commit()

    def test_late_commit_failure_restores_science_summary_and_keeps_index(self):
        output = self.root / 'SIMULATED'
        original = MeasurementWorker(SpecimenMeta('A'), ['NRM'], self.root, NoCommBackend())
        original.run()
        names = ('A', 'A.rmg', 'measurements.txt', 'provenance.json', 'susceptibility.json', 'artifact_index.json')
        before = {name: (output / name).read_bytes() for name in names}
        delayed = MeasurementWorker(SpecimenMeta('A'), ['AF20'], self.root, NoCommBackend())
        write = bundle_publication.BundlePublication._write_manifest

        def fail_commit(publication, manifest):
            if manifest['status'] == 'committed':
                self.assertTrue(delayed._progress._path('accepted.json').is_file())
                raise OSError('injected commit marker failure')
            write(publication, manifest)

        with patch.object(bundle_publication.BundlePublication, '_write_manifest', fail_commit):
            delayed.run()
        self.assertEqual({name: (output / name).read_bytes() for name in names}, before)
        self.assertFalse(delayed._progress._path('accepted.json').exists())
        self.assertIn('.rapidpy-attempts', delayed._publish_dir.parts)
        self.assertTrue((delayed._publish_dir / 'artifact_index.json').is_file())

    def test_quicklook_is_written_under_the_publication_lease_and_included_in_final_index(self):
        specimen = SpecimenMeta('A')
        output = self.root / 'SIMULATED'
        worker = MeasurementWorker(specimen, ['NRM'], self.root, NoCommBackend(), run_id='original')
        writes = []

        def publish_quicklook(path, steps, *, simulated):
            with self.assertRaises(HardwareSafetyError):
                MeasurementBundleWriter(self.root, specimen, simulated=True)
            provisional_index = json.loads((output / 'artifact_index.json').read_bytes())
            self.assertTrue(provisional_index['diagnostics_pending'])
            writes.append(path)
            return write_measurement_quicklook(path, steps, simulated=simulated)

        with patch('rapid_main.measurement_worker.write_measurement_quicklook', side_effect=publish_quicklook):
            worker.run()

        self.assertFalse(worker.quicklook_error)
        self.assertEqual(writes, [output / 'quicklook.json'])
        final_index = json.loads((output / 'artifact_index.json').read_bytes())
        self.assertNotIn('diagnostics_pending', final_index)
        entry, = [item for item in final_index['artifacts'] if item['name'] == 'quicklook_summary']
        self.assertTrue(entry['exists'])
        self.assertEqual(entry['sha256'], hashlib.sha256(writes[0].read_bytes()).hexdigest())

    def test_scientific_index_is_current_even_if_final_diagnostics_are_interrupted(self):
        class Interrupted(BaseException):
            pass

        first = MeasurementWorker(SpecimenMeta('A'), ['NRM'], self.root, NoCommBackend(), run_id='first')
        first.run()
        second = MeasurementWorker(SpecimenMeta('A'), ['AF20'], self.root, NoCommBackend(), run_id='second')
        with patch.object(second, '_finish_run', side_effect=Interrupted), self.assertRaises(Interrupted):
            second.run()

        output = self.root / 'SIMULATED'
        index = json.loads((output / 'artifact_index.json').read_text(encoding='utf-8'))
        self.assertEqual(index['run_id'], 'second')
        self.assertTrue(index['diagnostics_pending'])
        entries = {item['name']: item for item in index['artifacts']}
        self.assertFalse(entries['workflow_summary']['exists'])
        for name in ('rmg_file', 'susceptibility_summary', 'measurement_provenance', 'acquisition_progress:accepted.json'):
            artifact = entries[name]
            self.assertTrue(artifact['exists'])
            published = output / artifact['relative_path']
            self.assertEqual(artifact['sha256'], hashlib.sha256(published.read_bytes()).hexdigest())
