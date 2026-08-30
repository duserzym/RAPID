from __future__ import annotations

import unittest

from rapid_main.geometry import (
    Angular3D,
    Cartesian3D,
    InterpolationRange,
    InterpolationRanges,
    angular3d_to_cartesian3d,
    angular3d_to_viewer_cartesian3d,
    cartesian3d_diff_angle,
    cartesian3d_to_angular3d,
    cartesian3d_sum,
    cartesian3d_difference,
)


class TestGeometryConversions(unittest.TestCase):
    def test_angular_to_cartesian_matches_measurement_round_trip(self) -> None:
        source = Cartesian3D(x=1.0, y=2.0, z=3.0)
        angular = cartesian3d_to_angular3d(source)
        reconstructed = angular3d_to_cartesian3d(Angular3D(dec=angular.dec, inc=angular.inc, mag=angular.mag))
        self.assertAlmostEqual(reconstructed.x, source.x, places=12)
        self.assertAlmostEqual(reconstructed.y, source.y, places=12)
        self.assertAlmostEqual(reconstructed.z, source.z, places=12)

    def test_viewer_conversion_keeps_existing_z_convention(self) -> None:
        directional = Angular3D(dec=30.0, inc=10.0, mag=1.0)
        normal = angular3d_to_cartesian3d(directional)
        viewer = angular3d_to_viewer_cartesian3d(directional)
        self.assertAlmostEqual(viewer.x, normal.x)
        self.assertAlmostEqual(viewer.y, normal.y)
        self.assertAlmostEqual(viewer.z, -normal.z)

    def test_cartesian_vector_math(self) -> None:
        v1 = Cartesian3D(1.0, 0.0, 0.0)
        v2 = Cartesian3D(0.0, 1.0, 0.0)
        v3 = cartesian3d_sum(v1, v2)
        v4 = cartesian3d_difference(v3, v2)
        self.assertEqual(v4, v1)
        self.assertAlmostEqual(cartesian3d_diff_angle(v1, v2), 90.0)

    def test_interpolation_range_linear_value_and_bounds(self) -> None:
        segment = InterpolationRange(start=0.0, end=10.0, start_value=100.0, end_value=200.0)

        self.assertTrue(segment.contains(0.0))
        self.assertTrue(segment.contains(10.0))
        self.assertAlmostEqual(segment.value_at(2.5), 125.0)
        self.assertAlmostEqual(segment.value_at(-5.0, clamp=True), 100.0)
        self.assertAlmostEqual(segment.value_at(15.0, clamp=True), 200.0)
        with self.assertRaises(ValueError):
            segment.value_at(15.0)

    def test_interpolation_range_supports_descending_input_axis(self) -> None:
        segment = InterpolationRange(start=10.0, end=0.0, start_value=0.0, end_value=100.0)

        self.assertTrue(segment.contains(5.0))
        self.assertAlmostEqual(segment.value_at(5.0), 50.0)
        self.assertAlmostEqual(segment.value_at(12.0, clamp=True), 0.0)

    def test_interpolation_ranges_order_and_lookup(self) -> None:
        ranges = InterpolationRanges(
            [
                InterpolationRange(10.0, 20.0, 100.0, 200.0),
                InterpolationRange(0.0, 10.0, 0.0, 100.0),
            ]
        )

        self.assertEqual(ranges.low, 0.0)
        self.assertEqual(ranges.high, 20.0)
        self.assertAlmostEqual(ranges.value_at(5.0), 50.0)
        self.assertAlmostEqual(ranges.value_at(15.0), 150.0)
        self.assertAlmostEqual(ranges.value_at(-1.0, clamp=True), 0.0)
        self.assertAlmostEqual(ranges.value_at(25.0, clamp=True), 200.0)
        with self.assertRaises(ValueError):
            ranges.value_at(25.0)

    def test_interpolation_range_rejects_degenerate_segments(self) -> None:
        with self.assertRaises(ValueError):
            InterpolationRange(1.0, 1.0, 0.0, 10.0)

        with self.assertRaises(ValueError):
            InterpolationRanges([])


if __name__ == "__main__":
    unittest.main(verbosity=2)
