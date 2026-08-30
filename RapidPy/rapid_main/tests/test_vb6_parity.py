"""Deterministic VB6 parity fixtures for the bracketed measurement path.

Each fixture is computed by hand from the legacy source so a regression in the
RapidPy reduction shows up as a numeric difference, not just a shape change:

* ``MeasurementBlock.BaselineAdjustedSample`` -> ``(1 - i/5)`` / ``i/5`` weights
  and holder subtraction;
* ``MeasurementBlock.CorrectedSample`` -> the four rotations, with the
  ``mvarDirection`` sign for sample-up and sample-down;
* ``MeasurementBlock.Average`` / ``Kappa`` / ``FischerSD`` / ``induced`` /
  ``SigDrift`` / ``SigHolder`` / ``SigInduced``;
* ``RangeFact`` moment scaling and the ``Measure_Unfold`` specimen direction;
* the categorical flux-count rejection on X, Y, and Z.
"""
from __future__ import annotations

import math
from pathlib import Path
import tempfile
import unittest

from rapid_main.geometry import Cartesian3D, cartesian3d_to_angular3d
from rapid_main.holder_measurement import HolderMeasurementService
from rapid_main.holder_state import HolderStateStore
from rapid_main.io.measurement_bundle import MeasurementBundleWriter
from rapid_main.magnetometer import (
    DEFAULT_MAX_CONTINUOUS_ZERO_DRIFT,
    BlockAudit,
    BracketedMeasurementBlock,
    FluxCountDiscontinuityError,
    reduce_bracketed_measurement,
)

from tests.test_output_integrity import _meta, _step

# Archived one-flux-count increments in calibrated 2G raw units.
ARCHIVED_STEPS = {"X": 0.09000563, "Y": 0.106, "Z": 0.066}

ZERO_BEFORE = (0.100, 0.200, 0.300)
ZERO_AFTER = (0.110, 0.210, 0.310)
POSITIONS = (
    (1.000, 2.000, 3.000),
    (1.100, 2.100, 3.100),
    (1.200, 2.200, 3.200),
    (1.300, 2.300, 3.300),
)
HOLDER = (
    (0.010, 0.020, 0.030),
    (0.011, 0.021, 0.031),
    (0.012, 0.022, 0.032),
    (0.013, 0.023, 0.033),
)


def _expected_baseline_adjusted() -> list[tuple[float, float, float]]:
    """VB6 ``BaselineAdjustedSample``: sample - interpolated zero - holder."""

    adjusted = []
    for index in range(1, 5):
        end_weight = index / 5.0
        start_weight = 1.0 - end_weight
        vector = []
        for axis in range(3):
            baseline = start_weight * ZERO_BEFORE[axis] + end_weight * ZERO_AFTER[axis]
            vector.append(POSITIONS[index - 1][axis] - baseline - HOLDER[index - 1][axis])
        adjusted.append((vector[0], vector[1], vector[2]))
    return adjusted


def _expected_corrected(adjusted, *, is_up: bool):
    """VB6 ``CorrectedSample`` with ``mvarDirection`` = +1 up, -1 down."""

    direction = 1.0 if is_up else -1.0
    (x1, y1, z1), (x2, y2, z2), (x3, y3, z3), (x4, y4, z4) = adjusted
    return [
        (x1, -y1 * direction, z1 * direction),
        (y2, x2 * direction, z2 * direction),
        (-x3, y3 * direction, z3 * direction),
        (-y4, -x4 * direction, z4 * direction),
    ]


def _block(**overrides) -> BracketedMeasurementBlock:
    payload = dict(
        zero_before=ZERO_BEFORE,
        positions=POSITIONS,
        zero_after=ZERO_AFTER,
        holder_positions=HOLDER,
        is_up=True,
        range_factor=1.0e-5,
    )
    payload.update(overrides)
    return BracketedMeasurementBlock(**payload)


class BaselineAndHolderParityTests(unittest.TestCase):
    def test_baseline_interpolation_uses_the_i_over_five_weights(self) -> None:
        result = reduce_bracketed_measurement(_block())

        expected = _expected_baseline_adjusted()
        for measured, want in zip(result.baseline_adjusted_raw, expected):
            for axis in range(3):
                self.assertAlmostEqual(measured[axis], want[axis], places=12)

    def test_weights_match_the_legacy_constants_exactly(self) -> None:
        # BaselineFactor(1) = 1 - i/5 ; BaselineFactor(2) = i/5
        block = _block(
            zero_before=(0.0, 0.0, 0.0),
            zero_after=(5.0, 0.0, 0.0),
            positions=((0.0, 0.0, 0.0),) * 4,
            holder_positions=((0.0, 0.0, 0.0),) * 4,
        )
        # A 5.0 span exceeds the drift limit, so validate the weights through a
        # block that passes and scale analytically instead.
        with self.assertRaises(FluxCountDiscontinuityError):
            reduce_bracketed_measurement(block)

        small = _block(
            zero_before=(0.0, 0.0, 0.0),
            zero_after=(0.010, 0.0, 0.0),
            positions=((0.0, 0.0, 0.0),) * 4,
            holder_positions=((0.0, 0.0, 0.0),) * 4,
        )
        adjusted = reduce_bracketed_measurement(small).baseline_adjusted_raw
        for index in range(1, 5):
            self.assertAlmostEqual(adjusted[index - 1][0], -(index / 5.0) * 0.010, places=12)

    def test_holder_subtraction_is_per_position(self) -> None:
        without_holder = reduce_bracketed_measurement(
            _block(holder_positions=((0.0, 0.0, 0.0),) * 4)
        )
        with_holder = reduce_bracketed_measurement(_block())

        for index in range(4):
            for axis in range(3):
                self.assertAlmostEqual(
                    without_holder.baseline_adjusted_raw[index][axis]
                    - with_holder.baseline_adjusted_raw[index][axis],
                    HOLDER[index][axis],
                    places=12,
                )


class RotationParityTests(unittest.TestCase):
    def test_sample_up_rotations_match_corrected_sample(self) -> None:
        result = reduce_bracketed_measurement(_block(is_up=True))

        expected = _expected_corrected(_expected_baseline_adjusted(), is_up=True)
        for measured, want in zip(result.holder_frame_raw, expected):
            for axis in range(3):
                self.assertAlmostEqual(measured[axis], want[axis], places=12)

    def test_sample_down_rotations_flip_the_direction_terms(self) -> None:
        result = reduce_bracketed_measurement(_block(is_up=False))

        expected = _expected_corrected(_expected_baseline_adjusted(), is_up=False)
        for measured, want in zip(result.holder_frame_raw, expected):
            for axis in range(3):
                self.assertAlmostEqual(measured[axis], want[axis], places=12)

    def test_down_direction_only_changes_the_signed_terms(self) -> None:
        up = reduce_bracketed_measurement(_block(is_up=True)).holder_frame_raw
        down = reduce_bracketed_measurement(_block(is_up=False)).holder_frame_raw

        # Position 1: X is direction-independent, Y and Z carry the sign.
        self.assertAlmostEqual(up[0][0], down[0][0], places=12)
        self.assertAlmostEqual(up[0][1], -down[0][1], places=12)
        self.assertAlmostEqual(up[0][2], -down[0][2], places=12)
        # Position 2: X is direction-independent.
        self.assertAlmostEqual(up[1][0], down[1][0], places=12)
        self.assertAlmostEqual(up[1][1], -down[1][1], places=12)
        # Position 4: only X of the rotated frame is direction-independent.
        self.assertAlmostEqual(up[3][0], down[3][0], places=12)
        self.assertAlmostEqual(up[3][1], -down[3][1], places=12)


class StatisticsParityTests(unittest.TestCase):
    def setUp(self) -> None:
        self.result = reduce_bracketed_measurement(_block())
        self.adjusted = _expected_baseline_adjusted()
        self.corrected = _expected_corrected(self.adjusted, is_up=True)

    def test_average_vector_matches_measurement_block_average(self) -> None:
        for axis in range(3):
            expected = sum(vector[axis] for vector in self.corrected) / 4.0
            self.assertAlmostEqual(self.result.mean_raw[axis], expected, places=12)

    def test_moment_scaling_applies_range_fact(self) -> None:
        for axis in range(3):
            self.assertAlmostEqual(
                self.result.moment_emu[axis], self.result.mean_raw[axis] * 1.0e-5, places=18
            )
        self.assertAlmostEqual(
            self.result.moment_magnitude_emu,
            math.sqrt(sum(value * value for value in self.result.moment_emu)),
            places=18,
        )

    def test_induced_vector_negates_z_for_positions_three_and_four(self) -> None:
        expected_x = sum(vector[0] for vector in self.adjusted) / 4.0
        expected_y = sum(vector[1] for vector in self.adjusted) / 4.0
        expected_z = (
            self.adjusted[0][2] + self.adjusted[1][2] - self.adjusted[2][2] - self.adjusted[3][2]
        ) / 4.0

        self.assertAlmostEqual(self.result.induced_raw[0], expected_x, places=12)
        self.assertAlmostEqual(self.result.induced_raw[1], expected_y, places=12)
        self.assertAlmostEqual(self.result.induced_raw[2], expected_z, places=12)

    def test_fischer_csd_matches_eighty_one_over_root_kappa(self) -> None:
        unit_sum = [0.0, 0.0, 0.0]
        for vector in self.corrected:
            magnitude = math.sqrt(sum(value * value for value in vector))
            for axis in range(3):
                unit_sum[axis] += vector[axis] / magnitude
        resultant = math.sqrt(sum(value * value for value in unit_sum))
        kappa = (4 - 1) / (4 - resultant)

        self.assertAlmostEqual(self.result.fischer_sd_deg, 81.0 / math.sqrt(kappa), places=9)

    def test_signal_ratios_match_the_legacy_definitions(self) -> None:
        average_magnitude = math.sqrt(sum(value * value for value in self.result.mean_raw))
        drift = tuple(after - before for before, after in zip(ZERO_BEFORE, ZERO_AFTER))
        drift_magnitude = math.sqrt(sum(value * value for value in drift))
        induced_magnitude = math.sqrt(sum(value * value for value in self.result.induced_raw))

        self.assertAlmostEqual(self.result.average_magnitude_raw, average_magnitude, places=12)
        self.assertAlmostEqual(self.result.sig_drift, average_magnitude / drift_magnitude, places=9)
        self.assertAlmostEqual(
            self.result.sig_induced, average_magnitude / induced_magnitude, places=9
        )
        holder_magnitude = math.sqrt(
            sum(value * value for value in self.result.holder_average_raw)
        )
        self.assertAlmostEqual(
            self.result.sig_holder, average_magnitude / holder_magnitude, places=9
        )

    def test_specimen_direction_matches_measure_unfold_core_coordinates(self) -> None:
        vector = Cartesian3D(*self.result.moment_emu)
        direction = cartesian3d_to_angular3d(vector)

        expected_dec = math.degrees(math.atan2(vector.y, vector.x)) % 360.0
        horizontal = math.hypot(vector.x, vector.y)
        expected_inc = math.degrees(math.atan2(vector.z, horizontal))

        self.assertAlmostEqual(direction.dec, expected_dec, places=9)
        self.assertAlmostEqual(direction.inc, expected_inc, places=9)


class FluxCountRejectionParityTests(unittest.TestCase):
    def test_each_axis_one_count_step_is_rejected(self) -> None:
        for index, (axis, step) in enumerate(ARCHIVED_STEPS.items()):
            with self.subTest(axis=axis):
                zero_after = list(ZERO_BEFORE)
                zero_after[index] += step
                block = _block(zero_after=tuple(zero_after))

                with self.assertRaises(FluxCountDiscontinuityError) as ctx:
                    reduce_bracketed_measurement(block)

                validation = ctx.exception.validation
                self.assertEqual(validation.discontinuous_axes, (axis,))
                self.assertEqual(validation.flux_step_axes, (axis,))
                self.assertAlmostEqual(validation.deltas[index], step, places=9)

    def test_drift_just_below_the_limit_is_accepted(self) -> None:
        zero_after = (
            ZERO_BEFORE[0] + DEFAULT_MAX_CONTINUOUS_ZERO_DRIFT * 0.99,
            ZERO_BEFORE[1],
            ZERO_BEFORE[2],
        )

        result = reduce_bracketed_measurement(_block(zero_after=zero_after))

        self.assertTrue(result.validation.valid)

    def test_drift_just_above_the_limit_is_rejected(self) -> None:
        zero_after = (
            ZERO_BEFORE[0] + DEFAULT_MAX_CONTINUOUS_ZERO_DRIFT * 1.01,
            ZERO_BEFORE[1],
            ZERO_BEFORE[2],
        )

        with self.assertRaises(FluxCountDiscontinuityError):
            reduce_bracketed_measurement(_block(zero_after=zero_after))


class HolderReplacementParityTests(unittest.TestCase):
    def setUp(self) -> None:
        self._tmp = tempfile.TemporaryDirectory()
        self.tmp = Path(self._tmp.name)
        self.store = HolderStateStore(self.tmp / "holder.json")

    def tearDown(self) -> None:
        self._tmp.cleanup()

    def _holder_block(self, magnitude: float, *, zero_after=(0.0, 0.0, 0.0)):
        return BracketedMeasurementBlock(
            zero_before=(0.0, 0.0, 0.0),
            positions=(
                (magnitude, 0.0, 0.0),
                (0.0, magnitude, 0.0),
                (-magnitude, 0.0, 0.0),
                (0.0, -magnitude, 0.0),
            ),
            zero_after=zero_after,
            audit=BlockAudit(block_id="holder", is_holder_block=True),
        )

    def test_a_valid_replacement_holder_takes_over(self) -> None:
        HolderMeasurementService(lambda: self._holder_block(0.20), self.store).measure(
            holder_id="HOLDER-1"
        )
        HolderMeasurementService(lambda: self._holder_block(0.40), self.store).measure(
            holder_id="HOLDER-2"
        )

        self.assertEqual(self.store.current.holder_id, "HOLDER-2")
        self.assertAlmostEqual(self.store.current.positions[0][0], 0.40, places=12)

    def test_a_rejected_replacement_retains_the_previous_holder(self) -> None:
        HolderMeasurementService(lambda: self._holder_block(0.20), self.store).measure(
            holder_id="HOLDER-1"
        )

        outcome = HolderMeasurementService(
            lambda: self._holder_block(0.40, zero_after=(ARCHIVED_STEPS["X"], 0.0, 0.0)),
            self.store,
        ).measure(holder_id="HOLDER-2")

        self.assertFalse(outcome.installed)
        self.assertEqual(self.store.current.holder_id, "HOLDER-1")
        self.assertAlmostEqual(self.store.current.positions[0][0], 0.20, places=12)


class InterruptionAndResumeParityTests(unittest.TestCase):
    def setUp(self) -> None:
        self._tmp = tempfile.TemporaryDirectory()
        self.tmp = Path(self._tmp.name)

    def tearDown(self) -> None:
        self._tmp.cleanup()

    def test_interrupted_run_publishes_nothing_and_resume_completes_the_file(self) -> None:
        # Interrupted: two steps staged, then the run dies.
        interrupted = MeasurementBundleWriter(self.tmp, _meta("PAR01"))
        interrupted.append_step(_step("NRM"))
        interrupted.append_step(_step("AF10"))
        interrupted.abort()
        self.assertFalse((self.tmp / "PAR01").exists())

        # Re-run from the top: nothing was published, so nothing is skipped.
        with MeasurementBundleWriter(self.tmp, _meta("PAR01"), resume=True) as rerun:
            self.assertTrue(rerun.append_step(_step("NRM")))
            self.assertTrue(rerun.append_step(_step("AF10")))
            self.assertEqual(rerun.skipped_labels, ())

        # A later resume against the published file skips what already exists.
        with MeasurementBundleWriter(self.tmp, _meta("PAR01"), resume=True) as resumed:
            self.assertFalse(resumed.append_step(_step("NRM")))
            self.assertTrue(resumed.append_step(_step("AF20")))

        text = (self.tmp / "PAR01").read_text(encoding="latin-1")
        self.assertEqual(text.count("NRM"), 1)
        self.assertEqual(text.count("AF10"), 1)
        self.assertEqual(text.count("AF20"), 1)


if __name__ == "__main__":
    unittest.main(verbosity=2)
