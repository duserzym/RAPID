"""VB6 Sample.cls raw interchange fixtures, independent of the encoder."""
from dataclasses import replace
import unittest

from rapid_main.io.legacy_up import LegacyUpBlock, encode_up_file, parse_up_file, read_up_measurements
from rapid_main.magnetometer import reduce_bracketed_measurement

FIXTURE = """Sample|Direction|Blocks|MsmtType|Block|MsmtNum|X,Y,Z
A|U|1|Z|1|1|0.0000000E+0,0.0000000E+0,0.0000000E+0|2026-10-04 10:00:00
A|U|1|Z|1|2|0,0,0|2026-10-04 10:00:01
A|U|1|S|1|1|2,-4,6|2026-10-04 10:00:02
A|U|1|S|1|2|4,2,6|2026-10-04 10:00:03
A|U|1|S|1|3|-2,4,6|2026-10-04 10:00:04
A|U|1|S|1|4|-4,-2,6|2026-10-04 10:00:05
A|U|1|H|1|1|0,0,0|2026-10-04 10:00:06
A|U|1|H|1|2|0,0,0|2026-10-04 10:00:07
A|U|1|H|1|3|0,0,0|2026-10-04 10:00:08
A|U|1|H|1|4|0,0,0|2026-10-04 10:00:09
"""


class LegacyUpTests(unittest.TestCase):
    def test_reads_vb6_ten_row_block_and_timestamps(self):
        run = read_up_measurements(FIXTURE, 'A')
        self.assertEqual(run.sample, 'A')
        self.assertEqual(len(run.blocks), 1)
        block = run.blocks[0]
        self.assertEqual(block.zero_before, (0, 0, 0))
        self.assertEqual(block.positions[2], (-2, 4, 6))
        self.assertEqual(block.holder_positions, ((0, 0, 0),) * 4)
        self.assertEqual(block.timestamps[0], '2026-10-04 10:00:00')
        self.assertEqual(block.timestamps[-1], '2026-10-04 10:00:09')

    def test_explicit_calibration_reduces_without_fabricating_hardware_evidence(self):
        block = read_up_measurements(FIXTURE, 'A').blocks[0].to_measurement_block(
            range_factor=1e-5, axis_calibration=(2, 3, 4))
        self.assertIsNone(block.audit)
        self.assertEqual(block.observations, ())
        result = reduce_bracketed_measurement(block)
        for actual, expected in zip(result.moment_emu, (4e-5, 12e-5, 24e-5)):
            self.assertAlmostEqual(actual, expected)

    def test_encoder_round_trip_retains_two_blocks_directions_and_row_times(self):
        up = read_up_measurements(FIXTURE, 'A')
        up = replace(up, blocks=up.blocks * 2)
        down = replace(up, sample='B', blocks=tuple(replace(b, is_up=False) for b in up.blocks))
        encoded = encode_up_file((up, down))
        self.assertEqual(len(encoded.splitlines()), 41)
        self.assertEqual(encoded.count('\r\n'), 41)
        self.assertIn('A|U|2|Z|2|1|', encoded)
        self.assertIn('B|D|2|H|2|4|', encoded)
        self.assertEqual(parse_up_file(encoded), (up, down))

    def test_last_exact_up_run_selected_over_down_and_other_samples(self):
        first = read_up_measurements(FIXTURE, 'A')
        other = replace(first, sample='AA')
        latest = replace(first, blocks=(replace(first.blocks[0], zero_after=(0.1, 0, 0)),))
        down = replace(first, blocks=(replace(first.blocks[0], is_up=False),))
        text = encode_up_file((first, other, latest, down))
        self.assertEqual(read_up_measurements(text, 'A'), latest)
        self.assertEqual(read_up_measurements(text, 'AA'), other)
        with self.assertRaises(ValueError):
            read_up_measurements(text, 'a')

    def test_seven_column_older_rows_preserve_missing_time(self):
        text = '\n'.join(line.rsplit('|', 1)[0] if index else line
                         for index, line in enumerate(FIXTURE.splitlines()))
        run = read_up_measurements(text, 'A')
        self.assertEqual(run.blocks[0].timestamps, (None,) * 10)
        self.assertEqual(parse_up_file(encode_up_file((run,))), (run,))

    def test_empty_header_has_no_sample_and_zero_block_runs_rejected(self):
        header = FIXTURE.splitlines()[0] + '\r\n'
        self.assertEqual(parse_up_file(header), ())
        with self.assertRaises(ValueError):
            read_up_measurements(header, 'A')
        with self.assertRaises(ValueError):
            parse_up_file(FIXTURE.replace('|U|1|', '|U|0|'))

    def test_truncated_latest_run_never_falls_back_to_older_up(self):
        partial = '\n'.join(FIXTURE.splitlines()[1:-1])
        with self.assertRaisesRegex(ValueError, 'truncated'):
            read_up_measurements(FIXTURE + partial, 'A')

    def test_corrupt_identity_counters_roles_order_and_direction_fail_closed(self):
        for before, after in (
            ('A|U|1|Z|1|2', 'B|U|1|Z|1|2'),
            ('A|U|1|Z|1|2', 'A|D|1|Z|1|2'),
            ('A|U|1|Z|1|2', 'A|U|1|Z|1|1'),
            ('A|U|1|S|1|2', 'A|U|1|H|1|2'),
            ('|U|1|', '|Q|1|'), ('|U|1|', '|U|1garbage|'),
            ('|U|1|', '|U|-1|'), ('|U|1|', '|U|99999999999999999|'),
        ):
            with self.subTest(after=after), self.assertRaises(ValueError):
                parse_up_file(FIXTURE.replace(before, after))

    def test_nonfinite_incomplete_vectors_extra_columns_and_invalid_dates_rejected(self):
        for before, after in (
            ('|2,-4,6|', '|nan,-4,6|'), ('|2,-4,6|', '|2,-4,inf|'),
            ('|2,-4,6|', '|2,-4|'), ('|2,-4,6|', '|2,-4,6,8|'),
            ('|2,-4,6|', '|2,no,6|'), ('10:00:00', '25:00:00'),
            ('2026-10-04', '2026-02-30'), ('10:00:00', '10:00:00|extra'),
        ):
            with self.subTest(after=after), self.assertRaises(ValueError):
                parse_up_file(FIXTURE.replace(before, after))

    def test_missing_or_invalid_header_and_interior_blank_rows_rejected(self):
        for text in ('', FIXTURE.replace('Sample|Direction', 'Specimen|Direction'),
                     FIXTURE.replace('A|U|1|S|1|1', '\nA|U|1|S|1|1')):
            with self.subTest(text=text[:30]), self.assertRaises(ValueError):
                parse_up_file(text)

    def test_encoder_rejects_ambiguous_names_mixed_directions_or_bad_shape(self):
        run = read_up_measurements(FIXTURE, 'A')
        for bad in (replace(run, sample='A|B'), replace(run, sample='A\nB'),
                    replace(run, blocks=()), replace(run, blocks=(
                        run.blocks[0], replace(run.blocks[0], is_up=False))),
                    replace(run, blocks=(replace(run.blocks[0], positions=((1, 2, 3),)),))):
            with self.subTest(bad=bad), self.assertRaises(ValueError):
                encode_up_file((bad,))

    def test_external_calibration_must_be_explicit_and_finite(self):
        raw = read_up_measurements(FIXTURE, 'A').blocks[0]
        for factor, calibration in ((0, (1, 1, 1)), (float('nan'), (1, 1, 1)),
                                    (True, (1, 1, 1)), (1e-5, (1, float('inf'), 1))):
            with self.subTest(factor=factor), self.assertRaises(ValueError):
                raw.to_measurement_block(range_factor=factor, axis_calibration=calibration)

    def test_raw_block_export_preserves_readings_and_requires_row_times(self):
        raw = read_up_measurements(FIXTURE, 'A').blocks[0]
        block = raw.to_measurement_block(range_factor=1e-5, axis_calibration=(1, 1, 1))
        self.assertEqual(LegacyUpBlock.from_measurement_block(block, timestamps=raw.timestamps), raw)
        with self.assertRaises(ValueError):
            LegacyUpBlock.from_measurement_block(block, timestamps=())


if __name__ == '__main__':
    unittest.main()
