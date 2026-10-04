"""Queue scientific metadata is frozen before hardware startup."""
import json
from pathlib import Path
import tempfile
import unittest
from unittest.mock import patch

from rapid_main.data_model import SampleIndexRegistration, SampleIndexRegistrations, SpecimenMeta
from rapid_main.io.specimen_writer import write_header
from rapid_main.specimen_metadata import capture_specimen_metadata, restore_specimen_metadata
from rapid_main import specimen_metadata


class QueueSpecimenMetadataTests(unittest.TestCase):
    def setUp(self):
        self.directory = tempfile.TemporaryDirectory()
        self.addCleanup(self.directory.cleanup)
        self.root = Path(self.directory.name)
        self.registrations = SampleIndexRegistrations([SampleIndexRegistration('A',
            formation='Unit', location='Site', depth_cm='12')])

    def capture(self):
        return capture_specimen_metadata('A', sample_dir=self.root, registrations=self.registrations)

    def test_header_orientation_volume_and_provenance_survive_journal_round_trip(self):
        write_header(self.root / 'A', SpecimenMeta('A', comment='Original', core_plate_strike=123,
            core_plate_dip=45, bedding_strike=67, bedding_dip=8, volume=10.2, fold_axis=30, fold_plunge=15))
        snapshot = json.loads(json.dumps(self.capture()))
        resolution = restore_specimen_metadata(snapshot)
        self.assertEqual(resolution.source, 'specimen-header')
        self.assertEqual(resolution.header_path, (self.root / 'A').resolve())
        self.assertEqual((resolution.meta.core_plate_strike, resolution.meta.volume,
            resolution.meta.fold_axis, resolution.meta.site), (123, 10.2, 30, 'Unit'))
        resolution.meta.volume = 999
        self.assertEqual(restore_specimen_metadata(snapshot).meta.volume, 10.2)

    def test_changed_or_removed_header_is_rejected(self):
        path = self.root / 'A'
        write_header(path, SpecimenMeta('A', comment='Original'))
        snapshot = self.capture()
        original = path.read_bytes()
        path.write_bytes(original + b'changed\n')
        with self.assertRaisesRegex(ValueError, 'header changed'):
            restore_specimen_metadata(snapshot)
        path.unlink()
        with self.assertRaisesRegex(ValueError, 'header changed'):
            restore_specimen_metadata(snapshot)

    def test_missing_header_is_frozen_and_new_header_cannot_override_index(self):
        snapshot = self.capture()
        self.assertEqual(restore_specimen_metadata(snapshot).meta.comment, 'depth 12 cm')
        self.assertEqual(snapshot['sources'], {str((self.root / 'A').resolve()): None})
        write_header(self.root / 'A', SpecimenMeta('A', comment='Unexpected'))
        with self.assertRaisesRegex(ValueError, 'header changed'):
            restore_specimen_metadata(snapshot)

    def test_header_change_during_capture_is_rejected(self):
        original_resolver = specimen_metadata.resolve_specimen_meta
        def changing_resolver(*args, **kwargs):
            resolution = original_resolver(*args, **kwargs)
            write_header(self.root / 'A', SpecimenMeta('A', comment='Changed'))
            return resolution
        with patch.object(specimen_metadata, 'resolve_specimen_meta', changing_resolver):
            with self.assertRaisesRegex(ValueError, 'header changed while preparing'):
                self.capture()

    def test_nonfinite_header_metadata_cannot_enter_a_queue_snapshot(self):
        write_header(self.root / 'A', SpecimenMeta('A', volume=float('nan')))
        with self.assertRaises(ValueError):
            self.capture()
