from __future__ import annotations

import json
import sys
import tempfile
import unittest
from datetime import datetime
from pathlib import Path

VIEWER_ROOT = Path(__file__).resolve().parents[1]
RAPIDPY_ROOT = Path(__file__).resolve().parents[2]
RAPID_MAIN_ROOT = RAPIDPY_ROOT / "rapid_main"

sys.path.insert(0, str(VIEWER_ROOT))
sys.path.insert(0, str(RAPID_MAIN_ROOT))

from rapid_main.data_model import MeasurementStep, SpecimenMeta
from rapid_main.io.magic_writer import write_measurements
from rapid_main.io.measurement_bundle import MeasurementBundleWriter
from rapid_main.io.rmg_writer import RmgTreatment
from rapid_main.io.specimen_writer import append_step, write_header
from data_viewer.data_loading import (
    build_paleointensity_points,
    detect_experiment_type,
    load_cit_sam,
    load_magic_directory,
    watch_paths_for_dataset,
)


def _step(label: str, x: float = 1.0, y: float = 0.2, z: float = 0.1, error_angle: float = 2.0) -> MeasurementStep:
    return MeasurementStep(
        demag_label=label,
        gdec=10.0,
        ginc=30.0,
        sdec=10.0,
        sinc=30.0,
        moment=(x * x + y * y + z * z) ** 0.5,
        error_angle=error_angle,
        crdec=10.0,
        crinc=30.0,
        sdx=x,
        sdy=y,
        sdz=z,
        operator="tester",
        timestamp=datetime(2026, 5, 26, 12, 0, 0),
    )


class TestViewerDataLoading(unittest.TestCase):
    def test_detects_cit_af_dataset(self) -> None:
        with tempfile.TemporaryDirectory() as td:
            root = Path(td) / "ExampleSet"
            root.mkdir()
            sam_path = root / "ExampleSet.sam"
            sam_path.write_text("SPEC01\n", encoding="latin-1")
            specimen_path = root / "SPEC01"
            meta = SpecimenMeta(name="SPEC01", sample="S1")
            write_header(specimen_path, meta)
            append_step(specimen_path, _step("NRM"))
            append_step(specimen_path, _step("AF20", x=0.8, y=0.15, z=0.06))
            append_step(specimen_path, _step("AF40", x=0.5, y=0.08, z=0.03))

            dataset = load_cit_sam(sam_path)

        self.assertEqual(dataset.source_kind, "cit_sam")
        self.assertEqual(dataset.specimen_names, ["SPEC01"])
        self.assertEqual(dataset.specimens[0].experiment_type, "AF")
        self.assertEqual(len(dataset.specimens[0].steps), 3)
        watched = watch_paths_for_dataset(dataset)
        self.assertIn(sam_path, watched)
        self.assertIn(specimen_path, watched)

    def test_detects_magic_thermal_dataset(self) -> None:
        with tempfile.TemporaryDirectory() as td:
            root = Path(td)
            meta = SpecimenMeta(name="THERM01", sample="SAM1", site="SITE1", location="LOC1")
            steps = [_step("NRM"), _step("TT200", x=0.8, y=0.15, z=0.02), _step("TT300", x=0.5, y=0.07, z=0.01)]
            write_measurements(root / "measurements.txt", meta, steps)

            dataset = load_magic_directory(root)

        self.assertEqual(dataset.source_kind, "magic_dir")
        self.assertEqual(dataset.specimens[0].name, "THERM01")
        self.assertEqual(dataset.specimens[0].experiment_type, "Thermal")
        self.assertEqual(dataset.specimens[0].meta.site, "SITE1")
        watched = watch_paths_for_dataset(dataset)
        self.assertEqual(watched, [root, root / "measurements.txt"])

    def test_cit_reads_measured_susceptibility_and_watches_its_sidecar(self) -> None:
        with tempfile.TemporaryDirectory() as td:
            root = Path(td)
            sam = root / "set.sam"
            sam.write_text("SPEC.01\n", encoding="latin-1")
            specimen_path = root / "SPEC.01"
            write_header(specimen_path, SpecimenMeta(name="SPEC.01"))
            append_step(specimen_path, _step("NRM"))
            rmg = root / "SPEC.01.rmg"
            rmg.write_text(
                ",Level,Bias Field,Spin Speed,Hold Time,Mz,Std. Dev.,Mz/Vol,Moment Susceptibility\n"
                "NRM,0,,,,,,,0\n"
                "AF,250,,,,,,,0.002\n"
                "AF,250,,,,,,,-0.001\n"
                "TT,450,,,,,,,0.003\n"
                "SUSC,0,,,,,,,\n"
                "SUSC,0,,,,,,,0\n"
                "TT,500,,,,,,,nan\n",
                encoding="latin-1",
            )

            dataset = load_cit_sam(sam)
            specimen, = dataset.specimens
            self.assertIn(rmg, watch_paths_for_dataset(dataset))
            self.assertEqual([point.rmg_step for point in specimen.susceptibility_points], [2, 3, 4, 6])
            self.assertEqual(
                [(point.treatment, point.level, point.level_kind, point.susceptibility)
                 for point in specimen.susceptibility_points],
                [("AF", 250, "af", .002), ("AF", 250, "af", -.001),
                 ("TT", 450, "thermal", .003), ("SUSC", 0, "", 0)],
            )

    def test_loads_distinct_susceptibility_readings_from_current_and_archived_bundles(self) -> None:
        with tempfile.TemporaryDirectory() as td:
            root = Path(td)
            meta = SpecimenMeta(name="A")
            with MeasurementBundleWriter(root, meta) as writer:
                writer.append_step(_step("AF25"), susceptibility=.001,
                                   rmg_treatment=RmgTreatment("AF25", "AF", 250.0))
            with MeasurementBundleWriter(root, meta) as writer:
                writer.append_step(_step("TT450"), susceptibility=-.002,
                                   rmg_treatment=RmgTreatment("TT450", "TT", 450.0))
            primary = load_magic_directory(root)
            current, = primary.specimens
            reading, = current.susceptibility_points
            self.assertEqual((reading.level_kind, reading.level, reading.susceptibility), ("thermal", 450, -.002))
            self.assertIn(root / "A.rmg", watch_paths_for_dataset(primary))

            previous = json.loads((root / "provenance.json").read_bytes())["previous_publication"]
            archived, = load_magic_directory((root / previous["relative_path"]).parent).specimens
            reading, = archived.susceptibility_points
            self.assertEqual((reading.level_kind, reading.level, reading.susceptibility), ("af", 250, .001))
            self.assertIsNone(archived.rmg_path)

    def test_detects_izzi_and_builds_points(self) -> None:
        steps = [
            _step("NRM", x=1.0, y=0.0, z=0.0),
            _step("ZF300", x=0.85, y=0.02, z=0.0),
            _step("IF300", x=0.72, y=0.03, z=0.0),
            _step("PTRM300", x=0.75, y=0.03, z=0.0),
            _step("ZF350", x=0.61, y=0.03, z=0.0),
            _step("IF350", x=0.49, y=0.04, z=0.0),
        ]

        exp_type = detect_experiment_type(steps)
        points = build_paleointensity_points(steps)

        self.assertEqual(exp_type, "IZZI Thellier")
        self.assertGreaterEqual(len(points), 4)
        self.assertEqual(points[0].step_kind, "ZF")
        self.assertGreater(points[1].ptrm_gained, 0.0)


if __name__ == "__main__":
    unittest.main()
