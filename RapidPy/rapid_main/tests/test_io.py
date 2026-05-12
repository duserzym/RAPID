"""Unit tests for rapid_main I/O modules and data model."""
from __future__ import annotations

import sys
import tempfile
import unittest
from datetime import datetime
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[2]))

from rapid_main.data_model import MeasurementStep, RmgRecord, SpecimenMeta
from rapid_main.io.magic_method_codes import label_to_method_codes
from rapid_main.io.magic_writer import append_measurement, write_measurements
from rapid_main.io.rmg_reader import read_rmg
from rapid_main.io.rmg_writer import append_rmg_record, write_rmg_records
from rapid_main.io.sam_reader import find_sam_files, read_sam, specimen_path
from rapid_main.io.sequence_io import (
    load_sequence_json,
    load_sequence_txt,
    save_sequence_json,
    save_sequence_txt,
)
from rapid_main.io.specimen_reader import read_specimen
from rapid_main.io.specimen_writer import append_step, write_header
from rapid_main.runtime_estimator import RuntimeEstimator


# ─────────────────────────────────────────────────────────────────────────────
# Fixtures
# ─────────────────────────────────────────────────────────────────────────────

def _meta(name: str = "TEST01") -> SpecimenMeta:
    return SpecimenMeta(
        name=name,
        comment="Unit test specimen",
        core_plate_strike=10.0,
        core_plate_dip=5.0,
        bedding_strike=315.0,
        bedding_dip=12.0,
        volume=11.0,
        sample="TEST",
        site="SITE01",
        location="TestLoc",
    )


def _step(label: str = "AF20") -> MeasurementStep:
    return MeasurementStep(
        demag_label=label,
        gdec=15.3,
        ginc=42.1,
        sdec=10.0,
        sinc=38.0,
        moment=1.23e-6,
        error_angle=3.5,
        crdec=12.0,
        crinc=40.0,
        sdx=9e-7,
        sdy=6e-7,
        sdz=4e-7,
        operator="JSmith",
        timestamp=datetime(2026, 5, 12, 10, 30, 0),
    )


# ─────────────────────────────────────────────────────────────────────────────
# SAM reader
# ─────────────────────────────────────────────────────────────────────────────

class TestSamReader(unittest.TestCase):
    def test_read_sam(self) -> None:
        with tempfile.TemporaryDirectory() as td:
            sam = Path(td) / "MySet.sam"
            sam.write_text("# comment\nSPEC01\nSPEC02\n\nSPEC03\n")
            names = read_sam(sam)
            self.assertEqual(names, ["SPEC01", "SPEC02", "SPEC03"])

    def test_read_sam_empty(self) -> None:
        with tempfile.TemporaryDirectory() as td:
            sam = Path(td) / "empty.sam"
            sam.write_text("# all comments\n# another\n\n")
            self.assertEqual(read_sam(sam), [])

    def test_specimen_path(self) -> None:
        with tempfile.TemporaryDirectory() as td:
            sam = Path(td) / "MySam" / "MySam.sam"
            sam.parent.mkdir()
            sam.write_text("")
            p = specimen_path(sam, "SPEC01")
            self.assertEqual(p, sam.parent / "SPEC01")

    def test_find_sam_files(self) -> None:
        with tempfile.TemporaryDirectory() as td:
            root = Path(td)
            (root / "A.sam").write_text("")
            (root / "sub").mkdir()
            (root / "sub" / "B.sam").write_text("")
            found = find_sam_files(root)
            names = sorted(f.name for f in found)
            self.assertEqual(names, ["A.sam", "B.sam"])


# ─────────────────────────────────────────────────────────────────────────────
# Specimen writer / reader round-trip
# ─────────────────────────────────────────────────────────────────────────────

class TestSpecimenIO(unittest.TestCase):
    def _roundtrip(self, fold_axis=None, fold_plunge=None) -> None:
        meta = _meta()
        meta.fold_axis = fold_axis
        meta.fold_plunge = fold_plunge
        step1 = _step("NRM")
        step2 = _step("AF20")
        with tempfile.TemporaryDirectory() as td:
            sp = Path(td) / "TEST01"
            write_header(sp, meta)
            append_step(sp, step1)
            append_step(sp, step2)
            meta2, steps2 = read_specimen(sp, "TEST01")
        self.assertAlmostEqual(meta2.volume, 11.0, places=1)
        self.assertAlmostEqual(meta2.core_plate_strike, 10.0, places=0)
        self.assertAlmostEqual(meta2.bedding_dip, 12.0, places=0)
        self.assertEqual(len(steps2), 2)
        self.assertEqual(steps2[0].demag_label, "NRM")
        self.assertEqual(steps2[1].demag_label, "AF20")
        self.assertAlmostEqual(steps2[0].gdec, 15.3, places=1)
        self.assertAlmostEqual(steps2[0].moment, 1.23e-6, places=10)

    def test_roundtrip_basic(self) -> None:
        self._roundtrip()

    def test_roundtrip_with_fold(self) -> None:
        self._roundtrip(fold_axis=45.0, fold_plunge=10.0)

    def test_single_step(self) -> None:
        with tempfile.TemporaryDirectory() as td:
            sp = Path(td) / "SINGLE"
            write_header(sp, _meta("SINGLE"))
            append_step(sp, _step("TT400"))
            _, steps = read_specimen(sp)
            self.assertEqual(len(steps), 1)
            self.assertEqual(steps[0].demag_label, "TT400")


# ─────────────────────────────────────────────────────────────────────────────
# RMG writer / reader round-trip
# ─────────────────────────────────────────────────────────────────────────────

class TestRmgIO(unittest.TestCase):
    def test_roundtrip_af(self) -> None:
        with tempfile.TemporaryDirectory() as td:
            p = Path(td) / "TEST01.rmg"
            append_rmg_record(p, _step("AF20"), susceptibility=0.012)
            append_rmg_record(p, _step("AF50"), susceptibility=0.008)
            records = read_rmg(p)
        self.assertEqual(len(records), 2)
        self.assertEqual(records[0].step_type, "AF")
        self.assertAlmostEqual(records[0].step_value, 20.0, places=0)
        self.assertAlmostEqual(records[0].susceptibility, 0.012, places=5)
        self.assertEqual(records[1].step_type, "AF")
        self.assertAlmostEqual(records[1].step_value, 50.0, places=0)

    def test_write_records(self) -> None:
        records = [
            RmgRecord("IRM", 1000.0, 0.005, []),
            RmgRecord("AF", 200.0, 0.003, []),
        ]
        with tempfile.TemporaryDirectory() as td:
            p = Path(td) / "out.rmg"
            write_rmg_records(p, records)
            read_back = read_rmg(p)
        self.assertEqual(len(read_back), 2)
        self.assertEqual(read_back[0].step_type, "IRM")
        self.assertAlmostEqual(read_back[0].step_value, 1000.0, places=0)


# ─────────────────────────────────────────────────────────────────────────────
# MagIC method codes
# ─────────────────────────────────────────────────────────────────────────────

class TestMagicMethodCodes(unittest.TestCase):
    _CASES = [
        ("NRM",    "LP-NRM"),
        ("AF20",   "LP-DIR-AF"),
        ("AF200",  "LP-DIR-AF"),
        ("TT400",  "LP-DIR-T"),
        ("TH350",  "LP-DIR-T"),
        ("IRM500", "LP-IRM"),
        ("IRM1000","LP-IRM"),
        ("ARM100", "LP-ARM"),
    ]

    def test_label_mapping(self) -> None:
        for label, expected in self._CASES:
            with self.subTest(label=label):
                self.assertEqual(label_to_method_codes(label), expected)


# ─────────────────────────────────────────────────────────────────────────────
# MagIC writer
# ─────────────────────────────────────────────────────────────────────────────

class TestMagicWriter(unittest.TestCase):
    def test_write_measurements(self) -> None:
        meta = _meta()
        steps = [_step("NRM"), _step("AF20"), _step("AF50")]
        with tempfile.TemporaryDirectory() as td:
            p = Path(td) / "measurements.txt"
            write_measurements(p, meta, steps)
            lines = p.read_text(encoding="utf-8").splitlines()
        self.assertEqual(lines[0], "tab delimited\tmeasurements")
        self.assertIn("measurement", lines[1])
        self.assertEqual(len(lines), 5)  # header + col row + 3 data rows

    def test_append_measurement(self) -> None:
        meta = _meta()
        with tempfile.TemporaryDirectory() as td:
            p = Path(td) / "measurements.txt"
            append_measurement(p, meta, _step("NRM"))
            append_measurement(p, meta, _step("AF20"))
            lines = p.read_text(encoding="utf-8").splitlines()
        # 2 header lines + 2 data rows
        self.assertEqual(len(lines), 4)

    def test_unit_conversions(self) -> None:
        """moment in A·m² (×1e-3), AF field in T (×1e-4)."""
        meta = _meta()
        step = _step("AF20")
        # AF20 label → treat_ac_field_T() = 20 * 1e-4 = 0.002 T
        self.assertAlmostEqual(step.treat_ac_field_T(), 0.002, places=6)
        # moment 1.23e-6 emu * 1e-3 = 1.23e-9 A·m²
        self.assertAlmostEqual(step.magn_moment_Am2(), 1.23e-9, places=15)


# ─────────────────────────────────────────────────────────────────────────────
# Sequence I/O
# ─────────────────────────────────────────────────────────────────────────────

class TestSequenceIO(unittest.TestCase):
    _LABELS = ["NRM", "AF20", "AF50", "AF100", "IRM500", "ARM100"]

    def test_json_roundtrip(self) -> None:
        with tempfile.TemporaryDirectory() as td:
            p = Path(td) / "seq.json"
            save_sequence_json(p, self._LABELS)
            loaded = load_sequence_json(p)
        self.assertEqual(loaded, self._LABELS)

    def test_txt_roundtrip(self) -> None:
        with tempfile.TemporaryDirectory() as td:
            p = Path(td) / "seq.txt"
            save_sequence_txt(p, self._LABELS)
            loaded = load_sequence_txt(p)
        self.assertEqual(loaded, self._LABELS)

    def test_missing_file(self) -> None:
        self.assertEqual(load_sequence_json(Path("/nonexistent/path.json")), [])


# ─────────────────────────────────────────────────────────────────────────────
# RuntimeEstimator
# ─────────────────────────────────────────────────────────────────────────────

class TestRuntimeEstimator(unittest.TestCase):
    def setUp(self) -> None:
        self.est = RuntimeEstimator()

    def test_step_seconds(self) -> None:
        self.assertEqual(self.est.step_seconds("NRM"), 120)
        self.assertEqual(self.est.step_seconds("AF20"), 180)
        self.assertEqual(self.est.step_seconds("TT400"), 600)
        self.assertEqual(self.est.step_seconds("IRM1000"), 300)
        self.assertEqual(self.est.step_seconds("ARM100"), 240)

    def test_estimate_sequence(self) -> None:
        labels = ["NRM", "AF20", "AF50"]
        td = self.est.estimate_sequence(labels)
        self.assertEqual(int(td.total_seconds()), 120 + 180 + 180)

    def test_estimate_remaining(self) -> None:
        labels = ["NRM", "AF20", "AF50", "IRM500"]
        rem = self.est.estimate_remaining(labels, 2)
        # AF50 + IRM500 = 180 + 300
        self.assertEqual(int(rem.total_seconds()), 480)

    def test_format_duration(self) -> None:
        from datetime import timedelta
        self.assertEqual(self.est.format_duration(timedelta(seconds=45)), "45s")
        self.assertEqual(self.est.format_duration(timedelta(seconds=90)), "1m 30s")
        self.assertEqual(self.est.format_duration(timedelta(seconds=3700)), "1h 01m")

    def test_total_bar_text(self) -> None:
        labels = ["NRM", "AF20"]
        text = self.est.total_bar_text(labels)
        self.assertIn("2 steps", text)
        self.assertIn("~", text)

    def test_status_bar_text(self) -> None:
        labels = ["NRM", "AF20", "IRM500"]
        text = self.est.status_bar_text(labels, 1)
        self.assertIn("Step 2/3", text)
        self.assertIn("Est. remaining", text)

    def test_custom_step_times(self) -> None:
        est = RuntimeEstimator({"NRM": 999, "AF": 999, "_default": 999})
        self.assertEqual(est.step_seconds("NRM"), 999)

    def test_step_times_update(self) -> None:
        """step_times attribute can be updated at runtime (settings change)."""
        self.est.step_times = {"NRM": 200, "AF": 300, "_default": 100}
        self.assertEqual(self.est.step_seconds("NRM"), 200)


# ─────────────────────────────────────────────────────────────────────────────
# Config
# ─────────────────────────────────────────────────────────────────────────────

class TestAppConfig(unittest.TestCase):
    def test_default_roundtrip(self) -> None:
        from rapid_main.config import AppConfig
        with tempfile.TemporaryDirectory() as td:
            p = Path(td) / "config.json"
            cfg = AppConfig()
            cfg.general.operator = "TestOp"
            cfg.sequence.NRM = 999
            cfg.save(p)
            cfg2 = AppConfig.load(p)
        self.assertEqual(cfg2.general.operator, "TestOp")
        self.assertEqual(cfg2.sequence.NRM, 999)

    def test_missing_config_returns_defaults(self) -> None:
        from rapid_main.config import AppConfig
        cfg = AppConfig.load(Path("/nonexistent/path/config.json"))
        self.assertEqual(cfg.general.operator, "")
        self.assertEqual(cfg.sequence.NRM, 120)

    def test_as_estimator_dict(self) -> None:
        from rapid_main.config import SequenceTimesConfig
        sc = SequenceTimesConfig(NRM=111, AF=222)
        d = sc.as_estimator_dict()
        self.assertEqual(d["NRM"], 111)
        self.assertEqual(d["AF"], 222)
        self.assertIn("_default", d)


if __name__ == "__main__":
    unittest.main(verbosity=2)
