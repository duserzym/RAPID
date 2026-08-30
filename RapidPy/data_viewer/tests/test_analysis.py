from __future__ import annotations

import sys
import unittest
from datetime import datetime
from pathlib import Path

import numpy as np

VIEWER_ROOT = Path(__file__).resolve().parents[1]
RAPIDPY_ROOT = Path(__file__).resolve().parents[2]
RAPID_MAIN_ROOT = RAPIDPY_ROOT / "rapid_main"

sys.path.insert(0, str(VIEWER_ROOT))
sys.path.insert(0, str(RAPID_MAIN_ROOT))

from rapid_main.data_model import MeasurementStep, SpecimenMeta
from data_viewer.app import _has_paleointensity_data, _orthogonal_projection_components
from data_viewer.analysis import next_step_suggestion, principal_component_fit, summarize_paleointensity, vector_for_step
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

    def test_orthogonal_projection_matches_pmagpy_default_orientation(self) -> None:
        east, north, down = _orthogonal_projection_components(np.array([1.5, -0.25, -0.75], dtype=float))

        self.assertEqual(east, -0.25)
        self.assertEqual(north, 1.5)
        self.assertEqual(down, 0.75)

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