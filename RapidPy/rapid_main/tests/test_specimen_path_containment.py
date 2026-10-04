"""Real source/output paths stay within their selected dataset directories."""
import os
from pathlib import Path
import subprocess
import tempfile
import unittest

from rapid_main.data_model import SampleIndexRegistrations, SpecimenMeta
from tests import test_queue_and_bundle as bundle_fixtures
from rapid_main.io.measurement_bundle import MeasurementBundleWriter
from rapid_main.io.specimen_writer import write_header
from rapid_main.measurement_worker import MeasurementWorker
from rapid_main.specimen_metadata import capture_specimen_metadata, restore_specimen_metadata
from rapid_main.specimen_paths import contained_path


class SpecimenPathContainmentTests(unittest.TestCase):
    def setUp(self):
        self.directory = tempfile.TemporaryDirectory()
        self.addCleanup(self.directory.cleanup)
        self.root = Path(self.directory.name)
        self.source = self.root / 'source'
        self.source.mkdir()

    def link_directory(self, link, target):
        if os.name == 'nt':
            result = subprocess.run(['cmd', '/c', 'mklink', '/J', str(link), str(target)],
                capture_output=True, text=True, check=False)
            if result.returncode:
                self.skipTest('Junction creation unavailable: ' + result.stderr)
        else:
            link.symlink_to(target, target_is_directory=True)

    def test_source_rejects_escape_absolute_drive_device_and_stream_names(self):
        for name in ('../outside', '..\\outside', str(self.root / 'outside'), 'C:outside',
                     '\\server\\share\\A', 'NUL', 'folder/COM1.txt', 'A:stream', 'A\nB', '.'):
            with self.subTest(name=name), self.assertRaises(ValueError):
                capture_specimen_metadata(name, sample_dir=self.source, registrations=SampleIndexRegistrations([]))

    def test_nested_relative_headers_and_bundles_remain_supported(self):
        name = 'group/A'
        meta = SpecimenMeta(name, comment='Nested', volume=8.2)
        write_header(self.source / name, meta)
        snapshot = capture_specimen_metadata(name, sample_dir=self.source, registrations=SampleIndexRegistrations([]))
        self.assertEqual(snapshot['meta']['volume'], 8.2)
        output = self.root / 'output'
        bundle = MeasurementBundleWriter(output, meta)
        bundle.append_step(bundle_fixtures._step('NRM'))
        paths = bundle.commit()
        self.assertEqual(paths.specimen_file, output / 'group' / 'A')
        self.assertTrue(paths.specimen_file.is_file())
        self.assertTrue(paths.rmg_file.is_file())
        self.assertEqual(contained_path(self.source, 'group\\A'), self.source / 'group' / 'A')

    def test_invalid_output_names_fail_before_creating_or_resetting_any_directory(self):
        output = self.root / 'output'
        for name in ('../outside', str(self.root / 'outside'), 'CON', 'A:stream',
                     'measurements.txt', 'PROVENANCE.JSON', 'sub/../artifact_index.json', '.rapidpy-pending-1',
                     'specimens.txt/A', 'pulse_treatments/A'):
            with self.subTest(name=name), self.assertRaises(ValueError):
                MeasurementBundleWriter(output, SpecimenMeta(name))
            self.assertFalse(output.exists())

    def test_source_directory_link_cannot_read_an_external_header(self):
        outside = self.root / 'outside'
        outside.mkdir()
        write_header(outside / 'A', SpecimenMeta('A', comment='Outside'))
        self.link_directory(self.source / 'linked', outside)
        with self.assertRaisesRegex(ValueError, 'escapes'):
            capture_specimen_metadata('linked/A', sample_dir=self.source, registrations=SampleIndexRegistrations([]))

    def test_simulated_output_link_cannot_publish_outside_selected_directory(self):
        output, outside = self.root / 'output', self.root / 'outside'
        output.mkdir()
        outside.mkdir()
        self.link_directory(output / 'SIMULATED', outside)
        with self.assertRaisesRegex(ValueError, 'escapes|link'):
            MeasurementBundleWriter(output, SpecimenMeta('A'), simulated=True)
        self.assertEqual(list(outside.iterdir()), [])

    def test_publication_rejects_new_parent_link_and_preserves_external_data(self):
        output, outside = self.root / 'output', self.root / 'outside'
        outside.mkdir()
        sentinel = outside / 'A'
        sentinel.write_bytes(b'Original outside data')
        bundle = MeasurementBundleWriter(output, SpecimenMeta('group/A'))
        bundle.append_step(bundle_fixtures._step('NRM'))
        self.link_directory(output / 'group', outside)
        try:
            with self.assertRaisesRegex(ValueError, 'escapes|link'):
                bundle.commit()
            self.assertEqual(sentinel.read_bytes(), b'Original outside data')
        finally:
            bundle.abort()

    def test_shared_output_file_is_rejected_without_modifying_other_link(self):
        output = self.root / 'output'
        output.mkdir()
        sentinel = self.root / 'outside.txt'
        sentinel.write_bytes(b'Original outside data')
        os.link(sentinel, output / 'SIMULATION.txt')
        with self.assertRaisesRegex(ValueError, 'hard links'):
            MeasurementBundleWriter(output, SpecimenMeta('A'), simulated=True,
                allow_simulated_production_output=True)
        self.assertEqual(sentinel.read_bytes(), b'Original outside data')

    def test_atomic_json_temporary_file_cannot_overwrite_a_shared_external_file(self):
        output = self.root / 'output'
        bundle = MeasurementBundleWriter(output, SpecimenMeta('A'))
        bundle.append_step(bundle_fixtures._step('NRM'))
        sentinel = self.root / 'outside.txt'
        sentinel.write_bytes(b'Original outside data')
        os.link(sentinel, output / f'.rapidpy-steps.json.tmp-{os.getpid()}')
        try:
            with self.assertRaisesRegex(ValueError, 'hard links'):
                bundle.commit()
            self.assertEqual(sentinel.read_bytes(), b'Original outside data')
        finally:
            bundle.abort()

    def test_caller_metadata_edits_cannot_redirect_the_admitted_bundle(self):
        output = self.root / 'output'
        meta = SpecimenMeta('A', comment='Original')
        bundle = MeasurementBundleWriter(output, meta)
        meta.name = '../outside'
        meta.comment = 'Changed'
        bundle.append_step(bundle_fixtures._step('NRM'))
        paths = bundle.commit()
        self.assertEqual(paths.specimen_file, output / 'A')
        self.assertTrue(paths.specimen_file.read_text(encoding='latin-1').startswith('Original'))
        self.assertFalse((self.root / 'outside').exists())

    def test_worker_rejects_invalid_output_before_backend_preflight(self):
        class Backend:
            simulated = False
            calls = 0
            def preflight(self):
                self.calls += 1
        backend = Backend()
        with self.assertRaises(ValueError):
            MeasurementWorker(SpecimenMeta('../outside'), ['NRM'], self.root / 'output', backend=backend)
        self.assertEqual(backend.calls, 0)
        self.assertFalse((self.root / 'output').exists())

    def test_changed_source_parent_is_rejected_even_when_header_bytes_match(self):
        original = self.source / 'group'
        original.mkdir()
        write_header(original / 'A', SpecimenMeta('group/A', comment='Original'))
        snapshot = capture_specimen_metadata('group/A', sample_dir=self.source, registrations=SampleIndexRegistrations([]))
        outside = self.root / 'outside'
        original.rename(outside)
        self.link_directory(original, outside)
        with self.assertRaisesRegex(ValueError, 'escapes'):
            restore_specimen_metadata(snapshot)
