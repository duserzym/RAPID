"""Compare the measurement pipeline with the VB6 orientation and export equations."""

import csv
from dataclasses import replace
import json
import math
from pathlib import Path
import tempfile
import unittest

from rapid_main.data_model import SpecimenMeta
from rapid_main.hardware_contracts import NoCommBackend
from rapid_main.io.sample_index import read_sample_index_registrations
from rapid_main.io.sam_reader import read_sam
from rapid_main.io.specimen_writer import write_header
from rapid_main.magnetometer import BracketedMeasurementBlock
from rapid_main.measurement_worker import MeasurementWorker
from rapid_main.specimen_metadata import capture_specimen_metadata, restore_specimen_metadata
from rapid_main.specimen_orientation import specimen_directions


class BracketedStation(NoCommBackend):
    def read_squid(self):
        zero = (0., 0., 0.)
        return BracketedMeasurementBlock(zero, ((3.2, 2., 1.), (-2., 2.8, 1.),
            (-3.2, -2., 1.), (2., -2.8, 1.)), zero, is_up=True, range_factor=1e-5)


class SpecimenOrientationTests(unittest.TestCase):
    def setUp(self):
        self.meta = SpecimenMeta('A', core_plate_strike=142., core_plate_dip=21.,
            bedding_strike=88., bedding_dip=32., magnetic_declination_e=11., volume=8.2)

    def assert_direction(self, direction, dec, inc):
        self.assertAlmostEqual(direction.dec, dec, places=8)
        self.assertAlmostEqual(direction.inc, inc, places=8)

    def test_core_geographic_and_bedding_reference(self):
        # Expected values evaluated from the original Measure_Unfold and
        # Measure_Rotate component equations for this non-orthogonal fixture.
        directions = specimen_directions((3., -2., 1.), self.meta, raw_squid=True)
        self.assert_direction(directions.core, 326.3099324740, 15.5013186415)
        self.assert_direction(directions.geographic, 30.6625142291, -2.1676569038)
        self.assert_direction(directions.structural, 33.5478452529, 27.3912748718)

        folded = specimen_directions((3., -2., 1.), replace(self.meta, fold_axis=43., fold_plunge=24.))
        self.assert_direction(folded.structural, 28.3902450876, 27.3912748718)

    def test_ataf_axis_mapping_and_missing_orientation(self):
        meta = SpecimenMeta('A')
        for label, vector in (('aTAFX20', (-1., 3., 2.)), ('aTAFY20', (3., 1., 2.))):
            with self.subTest(label=label):
                directions = specimen_directions((3., -2., 1.), meta, label=label)
                expected = specimen_directions(vector, meta).core
                self.assertEqual((directions.core, directions.geographic, directions.structural), (expected,) * 3)

        # Zero angles explicitly recorded in a specimen header are usable.
        directions = specimen_directions((1., 0., 0.), replace(meta, core_orientation_available=True))
        self.assert_direction(directions.geographic, 270., 0.)

    def test_blank_magnetic_declination_keeps_fold_column_in_its_place(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / 'site.sam'
            path.write_text(f'CIT\nTest site\n{47.:5.1f} {-90.:5.1f} {"":5} {43.:5.1f} {24.:5.1f}\nA\n', encoding='latin-1')
            self.assertEqual(read_sam(path), ['A'])
            registration, = read_sample_index_registrations(path).entries
            self.assertIsNone(registration.magnetic_declination_e)

    def test_queue_metadata_and_worker_publish_cit_intensity_and_magic_moment(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            index = root / 'site.sam'
            index.write_text('CIT\nTest site\n 47.0 -90.0  11.0\nA\n', encoding='latin-1')
            write_header(root / 'A', self.meta)
            snapshot = capture_specimen_metadata('A', sample_dir=root,
                registrations=read_sample_index_registrations(index))
            meta = restore_specimen_metadata(snapshot).meta
            self.assertEqual(meta.magnetic_declination_e, 11.)
            self.assertTrue(meta.has_core_orientation)

            worker = MeasurementWorker(meta, ['NRM'], root / 'output', BracketedStation())
            finished, results = [], []
            worker.run_finished.connect(finished.append)
            worker.step_complete.connect(results.append)
            worker.run()
            self.assertEqual(finished, [False], worker._error_messages)
            result, = results
            moment = math.sqrt(14.) * 1e-5
            self.assertAlmostEqual(result.step.moment, moment)
            self.assertAlmostEqual(result.step.intensity_emu_per_cm3, moment / 8.2)

            output = root / 'output' / 'SIMULATED'
            fields = (output / 'A').read_text(encoding='latin-1').splitlines()[2].split()
            self.assertEqual(tuple(map(float, fields[1:5])), (30.7, -2.2, 33.5, 27.4))
            self.assertEqual(tuple(map(float, fields[7:9])), (326.3, 15.5))
            self.assertEqual(fields[5], f'{moment / 8.2:.2E}')
            self.assertEqual(fields[9], f'{math.sqrt(.16 / 3):.4E}')
            self.assertEqual(tuple(map(float, fields[10:12])), (0., 0.))

            rows = (output / 'measurements.txt').read_text().splitlines()
            row, = csv.DictReader(rows[1:], delimiter='\t')
            x, y, z = (float(row[f'magn_{axis}']) for axis in 'xyz')
            self.assertTrue(math.isclose(float(row['magn_moment']), moment * 1e-3, rel_tol=1e-5))
            self.assertAlmostEqual(math.degrees(math.atan2(y, x)), float(row['dir_dec']), places=3)
            self.assertAlmostEqual(math.degrees(math.atan2(z, math.hypot(x, y))), float(row['dir_inc']), places=3)
            payload = json.loads((output / 'provenance.json').read_text())
            self.assertEqual(payload['coordinate_correction']['volume_cm3_used'], 8.2)
            self.assertEqual(payload['coordinate_correction']['magnetic_declination_e'], 11.)
