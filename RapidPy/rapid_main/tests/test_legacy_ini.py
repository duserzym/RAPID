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
        self.assertEqual(cfg.susceptibility.port, "COM7")
        self.assertEqual(cfg.susceptibility.baud, 1200)
        self.assertEqual(cfg.susceptibility.parity, "N")
        self.assertEqual(cfg.susceptibility.bytesize, 8)
        self.assertEqual(cfg.susceptibility.stopbits, 2.0)
        self.assertEqual(cfg.susceptibility.coil_position, -22700)
        self.assertFalse(cfg.susceptibility.enabled)
        self.assertAlmostEqual(cfg.susceptibility.scale_factor, 1.0)
        self.assertAlmostEqual(cfg.susceptibility.moment_factor_cgs, 0.0000097914)

        self.assertAlmostEqual(cfg.calibration.cal_x, 2.2792, places=5)
        self.assertAlmostEqual(cfg.calibration.cal_y, -2.2940, places=5)
        self.assertAlmostEqual(cfg.calibration.cal_z, 1.6717, places=5)
        self.assertAlmostEqual(cfg.calibration.range_factor, 1e-5, places=12)
        self.assertAlmostEqual(cfg.squid.settle_time, 1.0, places=3)

        self.assertAlmostEqual(cfg.af_demag.settle, 90.0, places=3)
        self.assertEqual(cfg.af_demag.ramp_speed, "Slow (3 Hz)")

        self.assertAlmostEqual(cfg.irm_arm.arm_bias_max_mT, 1.01, places=3)
        self.assertAlmostEqual(cfg.irm_arm.arm_voltage_per_mT, 2.08, places=3)
        self.assertAlmostEqual(cfg.irm_arm.arm_voltage_max, 2.1, places=3)
        self.assertEqual(cfg.irm_arm.arm_peak_af, 100)
        self.assertEqual(cfg.irm_arm.arm_bias, .05)
        self.assertEqual((cfg.irm_arm.arm_board, cfg.irm_arm.arm_dac_channel, cfg.irm_arm.arm_gate_bit), (0,0,0))
        self.assertAlmostEqual(cfg.irm_arm.irm_max_field, 1200.0, places=3)
        self.assertEqual(cfg.irm_arm.irm_axis, "Y")

        self.assertTrue(cfg.vacuum.auto_pump)
        self.assertEqual(cfg.changer.speed_z, 15.0)
        self.assertEqual(cfg.motor_station.hole_slot, 46)
        self.assertEqual(cfg.motor_station.ports, {"changer_x": "COM3", "changer_y": "COM4", "updown": "COM5", "turning": "COM6"})
        self.assertEqual(set(cfg.motor_station.addresses.values()), {16})
        self.assertTrue(cfg.motor_station.use_xy_table)
        self.assertEqual(cfg.motor_station.xy_home, [-3, -2])
        self.assertEqual(cfg.motor_station.xy_positions['1'], [9590, -11916])
        self.assertEqual(cfg.motor_station.xy_positions['46'], [0, 39])
        self.assertEqual(len(cfg.motor_station.xy_positions), 100)

        self.assertIn("Program.LastLogin -> general.operator", report.mapped_fields)
        self.assertTrue(any("AFRampRate" in item for item in report.mapped_fields))
        self.assertTrue(any("COMPortSusceptibility" in item for item in report.mapped_fields))

    def test_xy_import_never_rounds_or_defaults_an_incomplete_coordinate_pair(self):
        text = '''[XYTable]
UseXYTableAPS=invalid
XYHomeX=-3
XY1X=1.2
XY1Y=4
XY2X=2147483648
XY2Y=4
XY3X=0
XY4X=nan
XY4Y=4
XY5X=0
XY5Y=39
'''
        config = AppConfig()
        config.motor_station.xy_positions = {'1': [500, 600]}
        config.motor_station.xy_home = [0, 0]
        config.motor_station.use_xy_table = True
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / 'station.ini'
            path.write_text(text)
            report = import_vb6_ini(config, path)
        self.assertIsNone(config.motor_station.use_xy_table)
        self.assertEqual(config.motor_station.xy_home, [])
        self.assertEqual(config.motor_station.xy_positions, {'5': [0, 39]})
        self.assertTrue(any('coordinate pair required' in warning for warning in report.warnings))

    def test_xy_calibration_survives_config_roundtrip(self):
        config = AppConfig()
        path = Path(__file__).resolve().parents[3] / 'VB6' / 'settings' / 'Paleomag_v3.INI'
        import_vb6_ini(config, path)
        with tempfile.TemporaryDirectory() as directory:
            saved = Path(directory) / 'config.json'
            config.save(saved)
            restored = AppConfig.load(saved)
        self.assertEqual(restored.motor_station.xy_positions, config.motor_station.xy_positions)
        self.assertEqual(restored.motor_station.xy_home, [-3, -2])
        self.assertTrue(restored.motor_station.use_xy_table)

    def test_new_partial_station_import_cannot_reuse_previous_geometry_or_wiring(self):
        config = AppConfig()
        original = Path(__file__).resolve().parents[3] / 'VB6' / 'settings' / 'Paleomag_v3.INI'
        import_vb6_ini(config, original)
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / 'different.ini'
            path.write_text('[COMPorts]\nCOMPortChanger=9\n[SampleChanger]\nSlotMin=1\n')
            import_vb6_ini(config, path)
        self.assertEqual(config.motor_station.ports, {'changer_x': 'COM9'})
        self.assertEqual(config.motor_station.addresses, {})
        self.assertEqual(set(config.motor_station.controller), {'slot_min', 'sample_height'})
        self.assertEqual(config.motor_station.hole_slot, 0)
        self.assertIsNone(config.motor_station.use_xy_table)
        self.assertEqual(config.motor_station.xy_positions, {})
        self.assertEqual(config.motor_station.xy_home, [])

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

    def test_dropoff_delay_import_is_finite_nonnegative_or_unaccepted(self):
        for raw, expected in (('1.2', 1.2), ('0', 0), ('-1', -1), ('nan', -1), ('bad', -1)):
            with self.subTest(raw=raw), tempfile.TemporaryDirectory() as directory:
                path = Path(directory) / 'station.ini'
                path.write_text('[Vacuum]\nDropoffVacuumDelay=' + raw)
                config = AppConfig()
                report = import_vb6_ini(config, path)
                self.assertEqual(config.vacuum.dropoff_delay_s, expected)
                if expected == -1:
                    self.assertTrue(any('DropoffVacuumDelay' in warning for warning in report.warnings))


if __name__ == "__main__":
    unittest.main(verbosity=2)

