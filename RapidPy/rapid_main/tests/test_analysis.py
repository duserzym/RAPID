import math
import unittest

from rapid_main.analysis import (
    extract_measurement_vectors,
    fit_principal_axis,
    moment_decay_summary,
    moment_statistics,
    moment_value_statistics,
    reading_cycle_statistics,
    vector_centroid,
    vector_orientation,
)
from rapid_main.data_model import MeasurementStep


def _step(label: str, x: float, y: float, z: float, moment: float) -> MeasurementStep:
    return MeasurementStep(
        demag_label=label,
        gdec=0.0,
        ginc=0.0,
        sdec=0.0,
        sinc=0.0,
        moment=moment,
        error_angle=0.0,
        crdec=0.0,
        crinc=0.0,
        sdx=x,
        sdy=y,
        sdz=z,
    )


class AnalysisTests(unittest.TestCase):
    def assertTupleAlmostEqual(
        self,
        actual: tuple[float, ...],
        expected: tuple[float, ...],
        places: int = 7,
    ) -> None:
        self.assertEqual(len(actual), len(expected))
        for actual_value, expected_value in zip(actual, expected):
            self.assertAlmostEqual(actual_value, expected_value, places=places)

    def test_extract_measurement_vectors_uses_specimen_cartesian_fields(self) -> None:
        steps = [
            _step("NRM", 1.0, 2.0, 3.0, 4.0),
            _step("AF20", -1.5, 0.25, 2.5, 3.0),
        ]

        vectors = extract_measurement_vectors(steps)

        self.assertEqual([vector.label for vector in vectors], ["NRM", "AF20"])
        self.assertEqual(
            [vector.vector for vector in vectors],
            [(1.0, 2.0, 3.0), (-1.5, 0.25, 2.5)],
        )
        self.assertAlmostEqual(vectors[0].magnitude, math.sqrt(14.0))

    def test_centroid_and_principal_axis_report_orientation_and_variance_evidence(self) -> None:
        steps = [
            _step("NRM", 1.0, 1.0, 0.0, 10.0),
            _step("AF10", 2.0, 2.0, 0.0, 8.0),
            _step("AF20", 3.0, 3.0, 0.0, 6.0),
            _step("AF30", 4.0, 4.0, 0.0, 4.0),
        ]

        self.assertTupleAlmostEqual(
            vector_centroid(extract_measurement_vectors(steps)),
            (2.5, 2.5, 0.0),
        )

        fit = fit_principal_axis(steps)

        self.assertEqual(fit.count, 4)
        self.assertTupleAlmostEqual(fit.centroid, (2.5, 2.5, 0.0))
        self.assertAlmostEqual(fit.declination_deg, 45.0)
        self.assertAlmostEqual(fit.inclination_deg, 0.0)
        self.assertAlmostEqual(fit.variance_fraction, 1.0)
        self.assertAlmostEqual(fit.rms_perpendicular, 0.0)
        self.assertFalse(fit.is_degenerate)

    def test_vector_orientation_handles_vertical_and_zero_vectors(self) -> None:
        self.assertTupleAlmostEqual(vector_orientation((0.0, 0.0, 2.0)), (0.0, 90.0))
        self.assertTupleAlmostEqual(vector_orientation((0.0, 0.0, 0.0)), (0.0, 0.0))

    def test_moment_statistics_include_decay_evidence(self) -> None:
        steps = [
            _step("NRM", 0.0, 0.0, 0.0, 10.0),
            _step("AF10", 0.0, 0.0, 0.0, 5.0),
            _step("AF20", 0.0, 0.0, 0.0, 2.5),
        ]

        stats = moment_statistics(steps)

        self.assertEqual(stats.count, 3)
        self.assertEqual(stats.minimum, 2.5)
        self.assertEqual(stats.maximum, 10.0)
        self.assertAlmostEqual(stats.mean, 5.833333333)
        self.assertEqual(stats.median, 5.0)
        self.assertAlmostEqual(stats.population_stdev, 3.118047822)
        self.assertAlmostEqual(stats.decay.final_ratio, 0.25)
        self.assertAlmostEqual(stats.decay.percent_loss, 75.0)
        self.assertTupleAlmostEqual(stats.decay.per_step_ratios, (0.5, 0.5))
        self.assertTrue(stats.decay.monotonic_nonincreasing)
        self.assertAlmostEqual(stats.decay.log_decay_slope, math.log(0.5))

    def test_moment_value_statistics_matches_step_statistics(self) -> None:
        steps = [
            _step("NRM", 1.0, 0.0, 0.0, 10.0),
            _step("AF10", 0.5, 0.0, 0.0, 5.0),
        ]

        self.assertEqual(
            moment_value_statistics([10.0, 5.0]),
            moment_statistics(steps),
        )

    def test_moment_decay_handles_zero_initial_and_non_monotonic_sequences(self) -> None:
        decay = moment_decay_summary([0.0, 2.0, 1.0, 3.0])

        self.assertEqual(decay.initial, 0.0)
        self.assertEqual(decay.final, 3.0)
        self.assertIsNone(decay.final_ratio)
        self.assertIsNone(decay.percent_loss)
        self.assertEqual(decay.per_step_ratios, (None, 0.5, 3.0))
        self.assertFalse(decay.monotonic_nonincreasing)
        self.assertIsNone(decay.log_decay_slope)

    def test_reading_cycle_statistics_preserve_axis_ranges_and_quality(self) -> None:
        cycle = reading_cycle_statistics(
            [
                (1.0, 2.0, 3.0),
                (3.0, 2.0, 1.0),
            ]
        )

        self.assertEqual(cycle.count, 2)
        self.assertTupleAlmostEqual(cycle.mean_vector, (2.0, 2.0, 2.0))
        self.assertTupleAlmostEqual(cycle.axis_ranges, (2.0, 0.0, 2.0))
        self.assertAlmostEqual(cycle.mean_magnitude, math.sqrt(12.0))
        self.assertAlmostEqual(cycle.rms_spread, math.sqrt(2.0))
        self.assertAlmostEqual(cycle.signal_to_drift or 0.0, math.sqrt(6.0))
        self.assertIsNotNone(cycle.directional_spread_deg)
        self.assertGreater(cycle.directional_spread_deg or 0.0, 0.0)

    def test_reading_cycle_statistics_reject_empty_and_nonfinite_samples(self) -> None:
        with self.assertRaisesRegex(ValueError, "at least one"):
            reading_cycle_statistics([])
        with self.assertRaisesRegex(ValueError, "non-finite"):
            reading_cycle_statistics([(1.0, math.nan, 3.0)])


if __name__ == "__main__":
    unittest.main()
