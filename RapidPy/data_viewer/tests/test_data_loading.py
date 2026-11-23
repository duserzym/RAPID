from __future__ import annotations

import json
import sys
import tempfile
import unittest
from datetime import datetime, timedelta
import hashlib
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
    load_cit_specimen,
    load_input,
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


def _write_review_index(root: Path, specimen: str) -> None:
    files = ("measurements.txt", f"{specimen}.rmg")
    (root / "artifact_index.json").write_text(json.dumps({
        "schema": "rapidpy.measurement.artifact_index.v1",
        "sample": specimen,
        "aborted": False,
        "artifacts": [
            {"relative_path": name, "exists": True, "sha256": hashlib.sha256((root / name).read_bytes()).hexdigest()}
            for name in files
        ],
    }), encoding="utf-8")


class TestViewerDataLoading(unittest.TestCase):
    def test_recovers_only_missing_csd_from_the_matching_indexed_rmg(self) -> None:
        with tempfile.TemporaryDirectory() as td:
            root = Path(td)
            with MeasurementBundleWriter(root, SpecimenMeta("A")) as writer:
                for number, (label, csd) in enumerate((("NRM", 2), ("AF20", 7), ("AF20", 11))):
                    step = _step(label, error_angle=csd)
                    step.timestamp += timedelta(seconds=number)
                    writer.append_step(step, retain_repeats=True)
            magic_path = root / "measurements.txt"
            header, columns, *rows = magic_path.read_text(encoding="utf-8").splitlines()
            rows = [row.rsplit("\t", 1)[0] + "\t" + csd for row, csd in zip(rows, ("", "0", "4"))]
            magic_path.write_text("\n".join([header, columns, *rows]) + "\n", encoding="utf-8")
            rmg_path = root / "A.rmg"
            original_rmg = rmg_path.read_bytes()

            def angles() -> list[float]:
                return [step.error_angle for step in load_magic_directory(root).specimens[0].steps]

            self.assertEqual(angles(), [0, 0, 4])
            _write_review_index(root, "A")
            old_magic = magic_path.read_bytes()
            self.assertEqual(angles(), [2, 0, 4])
            self.assertEqual(magic_path.read_bytes(), old_magic)
            self.assertEqual(rmg_path.read_bytes(), original_rmg)

            magic_path.write_bytes(old_magic + b"\n")
            self.assertEqual(angles(), [0, 0, 4])
            magic_path.write_bytes(old_magic)
            rmg_path.write_bytes(original_rmg + b"\r\n")
            self.assertEqual(angles(), [0, 0, 4])
            rmg_rows = original_rmg.splitlines()
            rmg_rows[2], rmg_rows[3] = rmg_rows[3], rmg_rows[2]
            rmg_path.write_bytes(b"\r\n".join(rmg_rows) + b"\r\n")
            _write_review_index(root, "A")
            self.assertEqual(angles(), [0, 0, 4])

    def test_recovers_legacy_csd_from_verified_archive_members(self) -> None:
        with tempfile.TemporaryDirectory() as td:
            root = Path(td)
            meta = SpecimenMeta("A")
            with MeasurementBundleWriter(root, meta) as writer:
                writer.append_step(_step("NRM", error_angle=9.5))
                writer.append_step(_step("AF20", error_angle=7.25))
                writer.append_step(_step("AF20", error_angle=5.5), retain_repeats=True)
            table, *rows = (root / "measurements.txt").read_text(encoding="utf-8").splitlines()
            legacy = "\n".join([table, *(row.rsplit("\t", 1)[0] for row in rows)]) + "\n"
            (root / "measurements.txt").write_text(legacy, encoding="utf-8")
            with MeasurementBundleWriter(root, meta) as writer:
                writer.append_step(_step("AF20", error_angle=13))
            provenance = json.loads((root / "provenance.json").read_bytes())
            archive = (root / provenance["previous_publication"]["relative_path"]).parent
            old_steps = load_magic_directory(archive).specimens[0].steps
            self.assertEqual([(step.demag_label, step.error_angle) for step in old_steps],
                             [("NRM", 9.5), ("AF20", 7.25), ("AF20", 5.5)])
            self.assertEqual(load_magic_directory(root).specimens[0].steps[0].error_angle, 13)

    def test_opens_individual_cit_files_and_preserves_the_full_specimen_name(self) -> None:
        with tempfile.TemporaryDirectory() as td:
            root = Path(td)
            for name, reader in (("SPEC01", load_input), ("SPEC.02", load_input), ("SPEC.csv", load_cit_specimen)):
                with self.subTest(name=name):
                    source = root / name
                    write_header(source, SpecimenMeta(name=name, volume=3.2, comment="Original specimen"))
                    append_step(source, _step("AF20"))
                    rmg_path = root / f"{name}.rmg"
                    rmg_path.write_text("AF,200,,,,,,,0.002\n", encoding="latin-1")
                    dataset = reader(source)
                    specimen, = dataset.specimens
                    self.assertEqual(dataset.source_kind, "cit_specimen")
                    self.assertEqual(specimen.magnitude_label, "Recorded magnitude")
                    self.assertEqual((specimen.name, specimen.meta.comment, specimen.meta.volume),
                                     (name, "Original specimen", 3.2))
                    measurement, = specimen.steps
                    self.assertEqual(measurement.demag_label, "AF20")
                    self.assertAlmostEqual(measurement.crinc, 30.0)
                    reading, = specimen.susceptibility_points
                    self.assertEqual((reading.level, reading.susceptibility), (200, .002))
                    self.assertEqual(watch_paths_for_dataset(dataset), [source, root, rmg_path])

    def test_opens_an_unmeasured_cit_specimen_but_rejects_unrecognized_content(self) -> None:
        with tempfile.TemporaryDirectory() as td:
            source = Path(td) / "SPEC01"
            write_header(source, SpecimenMeta(name=source.name, volume=3.4, core_orientation_available=False))
            specimen, = load_input(source).specimens
            self.assertEqual(specimen.steps, [])
            self.assertIn(source.parent, watch_paths_for_dataset(load_input(source)))

            original_header = source.read_text(encoding="latin-1")
            for content in ("unrelated file\n", original_header + "unreadable record\n"):
                with self.subTest(content=content):
                    source.write_text(content, encoding="latin-1")
                    with self.assertRaisesRegex(ValueError, "no readable CIT"):
                        load_input(source)

            write_header(source, SpecimenMeta(name=source.name))
            append_step(source, _step("NRM", x=float("nan")))
            with self.assertRaisesRegex(ValueError, "nonfinite CIT measurement"):
                load_input(source)

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
        self.assertEqual(dataset.specimens[0].steps[0].error_angle, 2.0)
        self.assertEqual(dataset.specimens[0].steps[0].crinc, 30.0)
        self.assertLess(dataset.specimens[0].steps[0].sdz, 0)
        self.assertEqual(dataset.specimens[0].magnitude_label, "Moment (A·m²)")
        self.assertAlmostEqual(dataset.specimens[0].steps[0].moment, steps[0].magn_moment_Am2(), places=8)
        watched = watch_paths_for_dataset(dataset)
        self.assertEqual(watched, [root, root / "measurements.txt"])

    def test_magic_components_use_down_positive_z_when_direction_is_missing(self) -> None:
        with tempfile.TemporaryDirectory() as td:
            root = Path(td)
            (root / "measurements.txt").write_text(
                "tab delimited\tmeasurements\n"
                "measurement\tspecimen\tmagn_x\tmagn_y\tmagn_z\tdir_csd\n"
                "A-NRM\tA\t0\t0\t2e-9\t10\n",
                encoding="utf-8",
            )
            specimen, = load_magic_directory(root).specimens
            step, = specimen.steps
            self.assertEqual((step.crinc, step.sinc, step.ginc, step.sdz), (90.0, 90.0, 90.0, -2e-9))
            self.assertEqual(step.error_angle, 10)

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
            self.assertEqual(archived.magnitude_label, current.magnitude_label)
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
