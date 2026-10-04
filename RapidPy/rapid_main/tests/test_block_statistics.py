from dataclasses import replace
import json
import math
from pathlib import Path
import tempfile
import unittest

from rapid_main.block_statistics import block_collection_statistics
from rapid_main.data_model import SpecimenMeta
from rapid_main.hardware_contracts import PreflightResult
from rapid_main.magnetometer import BlockAudit, BracketedMeasurementBlock
from rapid_main.measurement_worker import MeasurementWorker


def block(scale=1., *, up=True):
    # Four corrected positions are (2,0,1),(0,2,1),(-2,0,1),(0,-2,1),
    # multiplied by scale. Invert each position rotation explicitly.
    return BracketedMeasurementBlock(
        (0., 0., 0.), ((2*scale, 0., scale if up else -scale),) * 4,
        (0., 0., 0.), is_up=up, range_factor=1e-5)


class SequenceBackend:
    simulated = True

    def __init__(self, readings):
        self.readings = iter(readings)
        self.returned = 0

    def preflight(self):
        return PreflightResult.pass_ok()

    def is_available(self):
        return True

    def set_demag_step(self, label):
        pass

    def read_squid(self):
        return next(self.readings)

    def return_to_safe_state(self):
        self.returned += 1


class BlockStatisticsTests(unittest.TestCase):
    def test_single_block_position_statistics_match_hand_calculation(self):
        result = block_collection_statistics((block(),))
        self.assertEqual(result.mean_raw, (0, 0, 1))
        self.assertEqual(result.moment_emu, (0, 0, 1e-5))
        self.assertAlmostEqual(result.axis_sd_raw[0], math.sqrt(8/3))
        self.assertAlmostEqual(result.axis_sd_raw[1], math.sqrt(8/3))
        self.assertEqual(result.axis_sd_raw[2], 0)
        self.assertAlmostEqual(result.fischer_sd_deg, 81*math.sqrt((4-4/math.sqrt(5))/3))
        self.assertAlmostEqual(result.sig_induced, .5)
        self.assertAlmostEqual(result.sig_noise, math.sqrt(3/16))
        self.assertAlmostEqual(result.sig_drift, 1e9, delta=1e-6)
        self.assertEqual(result.up_to_down, 0)

    def test_up_down_collection_uses_all_eight_positions_and_subset_magnitudes(self):
        result = block_collection_statistics((block(), block(2, up=False)))
        self.assertEqual(result.block_count, 2)
        self.assertEqual(result.position_count, 8)
        self.assertEqual(result.mean_raw, (0, 0, 1.5))
        self.assertAlmostEqual(result.axis_sd_raw[2], math.sqrt(2/7))
        self.assertAlmostEqual(result.axis_sd_raw[0], math.sqrt(40/7))
        self.assertEqual(result.up_to_down, .5)
        self.assertEqual(result.error_horizontal_deg, 0)
        self.assertAlmostEqual(result.fischer_sd_deg, 81*math.sqrt((8-8/math.sqrt(5))/7))

    def test_calibration_applies_to_mean_and_absolute_component_sd(self):
        result = block_collection_statistics((replace(block(), axis_calibration=(-2, 3, 4)),))
        self.assertEqual(result.moment_emu, (0, 0, 4e-5))
        self.assertAlmostEqual(result.axis_sd_emu[0], math.sqrt(8/3)*2e-5)

    def test_signal_drift_and_holder_use_mean_block_magnitudes(self):
        blocks = []
        for offset in (.01, -.01):
            original = block()
            # Baseline varies linearly; holder contributes signed Z. Add both
            # to S so correction restores exactly the original four readings.
            positions = tuple((x + offset*(index+1)/5, y, z + offset)
                              for index, (x, y, z) in enumerate(original.positions))
            blocks.append(replace(original, positions=positions, zero_after=(offset, 0, 0),
                                  holder_positions=((0., 0., offset),) * 4))
        result = block_collection_statistics(blocks)
        self.assertAlmostEqual(result.sig_drift, 100.)
        self.assertAlmostEqual(result.sig_holder, 100.)

    def test_horizontal_error_preserves_vb6_angle_difference(self):
        up = replace(block(), positions=((1., -1., 1.), (1., 1., 1.),
                                         (-1., 1., 1.), (-1., -1., 1.)))
        down = replace(block(up=False), positions=((1., -1., -1.), (1., 1., -1.),
                                                   (-1., 1., -1.), (-1., -1., -1.)))
        result = block_collection_statistics((up, down))
        self.assertAlmostEqual(result.error_horizontal_deg, -270.)

    def test_unequal_subset_counts_weight_positions_instead_of_two_direction_means(self):
        result = block_collection_statistics((block(), block(), block(4, up=False)))
        self.assertEqual(result.mean_raw, (0, 0, 2))
        self.assertEqual(result.up_to_down, .25)

    def test_invalid_block_rejects_whole_collection_before_combination(self):
        with self.assertRaises(ValueError):
            block_collection_statistics((block(), replace(block(), zero_after=(10, 0, 0))))
        with self.assertRaises(ValueError):
            block_collection_statistics(())

    def test_changed_calibration_direction_type_and_audit_context_rejected(self):
        for bad in (replace(block(), range_factor=2e-5),
                    replace(block(), axis_calibration=(1, 2, 1)), replace(block(), is_up=1)):
            with self.subTest(bad=bad), self.assertRaises(ValueError):
                block_collection_statistics((block(), bad))
        audited = replace(block(), audit=BlockAudit(sample_name='A', run_id='run1'))
        with self.assertRaises(ValueError):
            block_collection_statistics((audited, block()))
        with self.assertRaises(ValueError):
            block_collection_statistics((audited, replace(audited, audit=BlockAudit(sample_name='B', run_id='run1'))))
        self.assertEqual(block_collection_statistics((audited, audited)).block_count, 2)

    def test_worker_publishes_position_csd_and_collection_provenance(self):
        backend = SequenceBackend((block(), block(2)))
        with tempfile.TemporaryDirectory() as directory:
            worker = MeasurementWorker(SpecimenMeta('A'), ['NRM'], directory,
                                       backend=backend, samples_per_position=2)
            completed, finished = [], []
            worker.step_complete.connect(completed.append)
            worker.run_finished.connect(finished.append)
            worker.run()
            self.assertEqual(finished, [False])
            expected = block_collection_statistics((block(), block(2)))
            self.assertEqual(completed[0].collection_stats, expected)
            self.assertEqual(completed[0].step.error_angle, expected.fischer_sd_deg)
            published = Path(directory) / 'SIMULATED'
            provenance = json.loads((published / 'provenance.json').read_text())
            self.assertEqual(provenance['block_collection_statistics'][0]['position_count'], 8)
            # Header occupies two lines; error-angle is the sixth step column.
            step_columns = (published / 'A').read_text().splitlines()[-1].split()
            self.assertIn(f'{expected.fischer_sd_deg:.1f}', step_columns)
            self.assertEqual(backend.returned, 1)

    def test_worker_rejects_mixed_payload_cycle_and_leaves_no_scientific_bundle(self):
        with tempfile.TemporaryDirectory() as directory:
            worker = MeasurementWorker(SpecimenMeta('A'), ['NRM'], directory,
                                       backend=SequenceBackend((block(), (1., 2., 3.))), samples_per_position=2)
            finished = []
            worker.run_finished.connect(finished.append)
            worker.run()
            self.assertEqual(finished, [True])
            self.assertFalse((Path(directory) / 'SIMULATED' / 'A').exists())

    def test_tuple_step_cannot_display_prior_block_quality(self):
        with tempfile.TemporaryDirectory() as directory:
            worker = MeasurementWorker(SpecimenMeta('A'), ['NRM', 'AF10'], directory,
                                       backend=SequenceBackend((block(), (1., 2., 3.))))
            completed = []
            worker.step_complete.connect(completed.append)
            worker.run()
            self.assertEqual(len(completed), 2)
            self.assertIsNotNone(completed[0].block_result)
            self.assertIsNone(completed[1].block_result)
            self.assertIsNone(completed[1].collection_stats)
            self.assertEqual(completed[1].step.error_angle, 0)


if __name__ == '__main__':
    unittest.main()
