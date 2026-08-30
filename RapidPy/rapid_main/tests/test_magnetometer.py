from __future__ import annotations

import math
import unittest

from rapid_main.config import CalibrationConfig
from rapid_main.magnetometer import (
    MagnetometerCalibration,
    calibrate_magnetometer_reading,
)


class MagnetometerTests(unittest.TestCase):
    def test_calibrates_raw_volts_to_moment_vector(self) -> None:
        reading = calibrate_magnetometer_reading(
            (1.0, -2.0, 0.5),
            MagnetometerCalibration(xcal=2.0, ycal=-3.0, zcal=4.0, range_factor=1e-5),
        )

        self.assertEqual(reading.raw_volts, (1.0, -2.0, 0.5))
        self.assertEqual(reading.corrected_volts, (1.0, -2.0, 0.5))
        self.assertAlmostEqual(reading.moment_emu[0], 2e-5)
        self.assertAlmostEqual(reading.moment_emu[1], 6e-5)
        self.assertAlmostEqual(reading.moment_emu[2], 2e-5)
        self.assertAlmostEqual(reading.moment_magnitude_emu, math.sqrt(44) * 1e-5)
        self.assertTrue(reading.ok)

    def test_subtracts_background_from_config_before_calibration(self) -> None:
        config = CalibrationConfig(
            cal_x=10.0,
            cal_y=20.0,
            cal_z=30.0,
            range_factor=1e-6,
            bg_x=0.1,
            bg_y=0.2,
            bg_z=0.3,
            bg_subtract=True,
        )

        reading = calibrate_magnetometer_reading((1.1, 1.2, 1.3), config)

        self.assertAlmostEqual(reading.corrected_volts[0], 1.0)
        self.assertAlmostEqual(reading.corrected_volts[1], 1.0)
        self.assertAlmostEqual(reading.corrected_volts[2], 1.0)
        self.assertAlmostEqual(reading.moment_emu[0], 10e-6)
        self.assertAlmostEqual(reading.moment_emu[1], 20e-6)
        self.assertAlmostEqual(reading.moment_emu[2], 30e-6)

    def test_quality_flags_identify_saturation_and_low_signal(self) -> None:
        reading = calibrate_magnetometer_reading(
            (9.7, 0.0, 0.0),
            MagnetometerCalibration(minimum_signal_emu=1.0, saturation_limit_v=9.5),
        )

        self.assertIn("near-saturation", reading.flags)
        self.assertIn("below-minimum-signal", reading.flags)
        self.assertFalse(reading.ok)

    def test_quality_flags_identify_zero_calibration_terms(self) -> None:
        reading = calibrate_magnetometer_reading(
            (1.0, 1.0, 1.0),
            MagnetometerCalibration(xcal=0.0, range_factor=0.0),
        )

        self.assertIn("zero-axis-calibration", reading.flags)
        self.assertIn("zero-range-factor", reading.flags)

    def test_rejects_malformed_axis_vectors(self) -> None:
        with self.assertRaisesRegex(ValueError, "exactly three"):
            calibrate_magnetometer_reading((1.0, 2.0))
        with self.assertRaisesRegex(TypeError, "calibration"):
            calibrate_magnetometer_reading((1.0, 2.0, 3.0), object())  # type: ignore[arg-type]


if __name__ == "__main__":
    unittest.main(verbosity=2)
