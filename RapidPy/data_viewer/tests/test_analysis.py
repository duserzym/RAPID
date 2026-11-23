from __future__ import annotations

import sys
import unittest
from dataclasses import replace
from datetime import datetime
import math
from pathlib import Path

import numpy as np

VIEWER_ROOT = Path(__file__).resolve().parents[1]
RAPIDPY_ROOT = Path(__file__).resolve().parents[2]
RAPID_MAIN_ROOT = RAPIDPY_ROOT / "rapid_main"

sys.path.insert(0, str(VIEWER_ROOT))
sys.path.insert(0, str(RAPID_MAIN_ROOT))

from rapid_main.data_model import MeasurementStep, SpecimenMeta
from data_viewer.app import _has_paleointensity_data, _orthogonal_projection_components, _step_label_indices
from data_viewer.analysis import (
    next_step_suggestion, principal_component_fit, principal_component_segment, stereonet_csd_arcs,
    summarize_paleointensity, vector_for_step,
)
from data_viewer.data_loading import PaleointensityPoint, ViewerSpecimen


def _step(label: str, x: float, y: float, z: float) -> MeasurementStep:
    moment = (x * x + y * y + z * z) ** 0.5
    return MeasurementStep(
        demag_label=label,
        gdec=0.0,
        ginc=0.0,
        sdec=0.0,
        sinc=0.0,
        moment=moment,
        error_angle=3.0,
        crdec=0.0,
        crinc=0.0,
        sdx=x,
        sdy=y,
        sdz=z,
        operator="test",
        timestamp=datetime(2026, 5, 26, 12, 0, 0),
    )


class TestViewerAnalysis(unittest.TestCase):
    def test_principal_component_fit(self) -> None:
        steps = [_step("NRM", 1.0, 0.0, 0.0), _step("AF20", 0.8, 0.1, 0.0), _step("AF40", 0.6, 0.1, 0.0)]
        fit = principal_component_fit(steps)
        self.assertIsNotNone(fit)
        assert fit is not None
        self.assertGreater(fit.eigenvalues[0], fit.eigenvalues[1])
        self.assertLess(fit.mad, 25.0)

    def test_free_fit_reports_the_removed_remanence_direction(self) -> None:
        steps = [_step("NRM", 3, 0, 0), _step("AF20", 2, 0, 0), _step("AF40", 1, 0, 0)]
        fit = principal_component_fit(steps)
        assert fit is not None
        np.testing.assert_allclose(fit.direction, [1, 0, 0])
        self.assertAlmostEqual(fit.dang, 0.0)

    def test_anchored_fit_points_from_the_origin_toward_the_readings(self) -> None:
        steps = [_step("NRM", 1, 0, 0), _step("AF20", 2, 0, 0), _step("AF40", 3, 0, 0)]
        for order in (steps, steps[::-1]):
            fit = principal_component_fit(order, include_origin=True)
            assert fit is not None
            np.testing.assert_allclose(fit.direction, [1, 0, 0])

    def test_fit_angles_are_independent_of_the_moment_scale(self) -> None:
        steps = [
            replace(_step(label, moment, 0, 0), crdec=dec, crinc=inc)
            for label, moment, dec, inc in (
                ("NRM", 1.0, 22.0, 43.0),
                ("AF20", 0.8, 26.0, 39.0),
                ("AF40", 0.5, 31.0, 36.0),
                ("AF60", 0.3, 28.0, 32.0),
            )
        ]
        reference = principal_component_fit(steps)
        assert reference is not None
        self.assertGreater(reference.mad, 0.1)
        for factor in (1e-9, 1e-18):
            with self.subTest(factor=factor):
                scaled = principal_component_fit([replace(step, moment=step.moment * factor) for step in steps])
                assert scaled is not None
                self.assertAlmostEqual(scaled.mad, reference.mad, places=10)
                self.assertAlmostEqual(scaled.dang, reference.dang, places=10)
                np.testing.assert_allclose(scaled.center / factor, reference.center, rtol=1e-12)
                np.testing.assert_allclose(scaled.eigenvalues / factor**2, reference.eigenvalues, rtol=1e-10)

    def test_identical_steps_do_not_define_a_free_fit(self) -> None:
        step = _step("NRM", 1e-9, 0.0, 0.0)
        self.assertIsNone(principal_component_fit([step] * 3))
        self.assertIsNotNone(principal_component_fit([step] * 3, include_origin=True))

    def test_coordinate_system_mapping(self) -> None:
        step = MeasurementStep(
            demag_label="NRM",
            gdec=10.0,
            ginc=20.0,
            sdec=30.0,
            sinc=40.0,
            moment=1.0,
            error_angle=3.0,
            crdec=50.0,
            crinc=60.0,
            sdx=0.0,
            sdy=0.0,
            sdz=0.0,
            operator="test",
            timestamp=datetime(2026, 5, 26, 12, 0, 0),
        )

        specimen_vec = vector_for_step(step, "specimen")
        geographic_vec = vector_for_step(step, "geographic")
        tilt_vec = vector_for_step(step, "tilt-corrected")

        self.assertNotEqual(tuple(specimen_vec), tuple(geographic_vec))
        self.assertNotEqual(tuple(specimen_vec), tuple(tilt_vec))
        self.assertNotEqual(tuple(geographic_vec), tuple(tilt_vec))

    def test_af_suggestion(self) -> None:
        specimen = ViewerSpecimen(
            name="AF01",
            meta=SpecimenMeta(name="AF01"),
            steps=[_step("NRM", 1.0, 0.0, 0.0), _step("AF10", 0.9, 0.02, 0.0), _step("AF20", 0.78, 0.03, 0.0)],
            experiment_type="AF",
            source_kind="cit_sam",
        )
        suggestion = next_step_suggestion(specimen)
        self.assertIn("AF", suggestion.title)
        self.assertIn("mT", suggestion.suggested_step)

    def test_izzi_summary(self) -> None:
        points = [
            PaleointensityPoint(300.0, "ZF", 0.82, 0.00, "ZF300"),
            PaleointensityPoint(300.0, "IF", 0.70, 0.30, "IF300"),
            PaleointensityPoint(350.0, "ZF", 0.60, 0.30, "ZF350"),
            PaleointensityPoint(350.0, "IF", 0.48, 0.52, "IF350"),
        ]
        summary = summarize_paleointensity(points)
        self.assertEqual(summary.point_count, 4)
        self.assertIsNotNone(summary.slope)

    def test_orthogonal_projections_match_vb6_on_upward_positive_screen_axes(self) -> None:
        vector = np.array([1.5, 0.25, -0.75], dtype=float)
        self.assertEqual(_orthogonal_projection_components(vector), (0.25, 1.5, -0.75))
        self.assertEqual(_orthogonal_projection_components(vector, "ew"), (1.5, -0.25, -0.75))

    def test_step_label_modes_include_single_endpoint_once(self) -> None:
        self.assertEqual(_step_label_indices(0, "endpoints"), ())
        self.assertEqual(_step_label_indices(1, "endpoints"), (0,))
        self.assertEqual(_step_label_indices(4, "endpoints"), (0, 3))
        self.assertEqual(tuple(_step_label_indices(4, "all")), (0, 1, 2, 3))
        self.assertEqual(_step_label_indices(4, "hidden"), ())

    def test_fit_segment_spans_selected_readings_at_their_own_magnitude(self) -> None:
        steps = [_step("NRM", 8e-10, 0, 0), _step("AF20", 3e-10, 0, 0)]
        free = principal_component_segment(steps)
        anchored = principal_component_segment(steps, include_origin=True)
        self.assertIsNotNone(free)
        self.assertIsNotNone(anchored)
        np.testing.assert_allclose(free, [[8e-10, 0, 0], [3e-10, 0, 0]], rtol=1e-12, atol=1e-23)
        np.testing.assert_allclose(anchored, [[8e-10, 0, 0], [0, 0, 0]], rtol=1e-12, atol=1e-23)

    def test_csd_contour_at_each_pole_has_the_expected_radius(self) -> None:
        for inclination in (-90, 90):
            lower, points = stereonet_csd_arcs(31, inclination, 25)[0]
            self.assertEqual(lower, inclination > 0)
            radius = math.sqrt(1 - math.cos(math.radians(25)))
            np.testing.assert_allclose(np.linalg.norm(points, axis=1), radius, atol=1e-12)
            np.testing.assert_allclose(points[0], points[-1], atol=1e-12)
        for invalid in (0, float("nan"), 180, 999):
            self.assertEqual(stereonet_csd_arcs(30, 20, invalid), [])

    def test_csd_contour_crosses_equator_at_the_recorded_angular_distance(self) -> None:
        dec, inc, csd = map(math.radians, (32, 12, 25))
        axis = np.array([math.cos(inc) * math.cos(dec), math.cos(inc) * math.sin(dec), math.sin(inc)])
        arcs = stereonet_csd_arcs(32, 12, 25)
        self.assertEqual({lower for lower, _ in arcs}, {False, True})
        for (first_lower, first), (next_lower, following) in zip(arcs, arcs[1:]):
            self.assertNotEqual(first_lower, next_lower)
            self.assertAlmostEqual(math.hypot(*first[-1]), 1.0)
            np.testing.assert_allclose(first[-1], following[0])
        for lower, points in arcs:
            for x, y in points:
                depth = (1 - x*x - y*y) * (1 if lower else -1)
                scale = math.sqrt(1 + abs(depth))
                self.assertAlmostEqual(np.dot(axis, [y*scale, x*scale, depth]), math.cos(csd), places=12)

    def test_paleointensity_helper_detects_points(self) -> None:
        specimen = ViewerSpecimen(
            name="IZZI01",
            meta=SpecimenMeta(name="IZZI01"),
            steps=[_step("NRM", 1.0, 0.0, 0.0)],
            experiment_type="IZZI Thellier",
            source_kind="magic_dir",
            paleointensity_points=[PaleointensityPoint(300.0, "ZF", 0.82, 0.00, "ZF300")],
        )

        self.assertTrue(_has_paleointensity_data(specimen))
        self.assertFalse(_has_paleointensity_data(None))


if __name__ == "__main__":
    unittest.main()
