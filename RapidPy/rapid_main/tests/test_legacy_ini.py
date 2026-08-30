"""Tests for VB6 legacy INI import migration."""
from __future__ import annotations

import sys
import tempfile
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[2]))

from rapid_main.config import AppConfig
from rapid_main.legacy_ini import import_vb6_ini


class TestLegacyIniImport(unittest.TestCase):
    def test_import_from_paleomag_v3_ini(self) -> None:
        ini_path = Path(__file__).resolve().parents[3] / "VB6" / "settings" / "Paleomag_v3.INI"
        cfg = AppConfig()
        report = import_vb6_ini(cfg, ini_path)

        self.assertEqual(cfg.general.operator, "kateakin")
        self.assertEqual(cfg.general.data_dir, "C:\\Dropbox\\Hargraves_Data\\")
        self.assertEqual(cfg.general.backup_dir, "c")
        self.assertEqual(cfg.general.sample_dir, "C:\\Users\\RockMagnetometer\\Desktop\\PaleoMag 2013\\Settings")

        self.assertEqual(cfg.squid.port, "COM1")
        self.assertEqual(cfg.vacuum.port, "COM1")
        self.assertEqual(cfg.changer.port, "COM3")

        self.assertAlmostEqual(cfg.calibration.cal_x, 2.2792, places=5)
        self.assertAlmostEqual(cfg.calibration.cal_y, -2.2940, places=5)
        self.assertAlmostEqual(cfg.calibration.cal_z, 1.6717, places=5)
        self.assertAlmostEqual(cfg.calibration.range_factor, 1e-5, places=12)
        self.assertAlmostEqual(cfg.squid.settle_time, 1.0, places=3)

        self.assertAlmostEqual(cfg.af_demag.settle, 90.0, places=3)
        self.assertEqual(cfg.af_demag.ramp_speed, "Slow (3 Hz)")

        self.assertAlmostEqual(cfg.irm_arm.arm_peak_af, 10.1, places=3)
        self.assertAlmostEqual(cfg.irm_arm.arm_bias, 0.208, places=3)
        self.assertAlmostEqual(cfg.irm_arm.irm_max_field, 420.0, places=3)
        self.assertEqual(cfg.irm_arm.irm_axis, "Y axis")

        self.assertTrue(cfg.vacuum.auto_pump)
        self.assertEqual(cfg.changer.speed_z, 46.0)

        self.assertIn("Program.LastLogin -> general.operator", report.mapped_fields)
        self.assertTrue(any("AFRampRate" in item for item in report.mapped_fields))

    def test_invalid_numbers_are_defaulted_with_warnings(self) -> None:
        ini_text = """[Program]
LastLogin=operator1

[MagnetometerCalibration]
XCal=bad

[AF]
AFWait=bad
AFRampRate=bad

[COMPorts]
COMPortSquids=0

[IRMAxial]
IRMAxialVoltMax=bad

[IRMTrans]
IRMTransVoltMax=not_a_number

[SampleChanger]
HoleSlotNum=bad
"""
        with tempfile.TemporaryDirectory() as td:
            ini_path = Path(td) / "legacy.ini"
            ini_path.write_text(ini_text, encoding="utf-8")

            cfg = AppConfig()
            report = import_vb6_ini(cfg, ini_path)

        self.assertEqual(cfg.general.operator, "operator1")
        self.assertEqual(cfg.af_demag.settle, 1.5)
        self.assertEqual(cfg.af_demag.ramp_speed, "Medium (8 Hz)")
        self.assertEqual(cfg.squid.port, "COM1")
        self.assertEqual(cfg.irm_arm.irm_axis, "Z (up-axis)")
        self.assertIn("MagnetometerCalibration.XCal", "\n".join(report.warnings))
        self.assertIn("AFWait", "\n".join(report.warnings))


if __name__ == "__main__":
    unittest.main(verbosity=2)

