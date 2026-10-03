from dataclasses import asdict
from pathlib import Path
import unittest

from rapid_main.config import AppConfig, PulseIrmConfig
from rapid_main.legacy_ini import import_vb6_ini
from rapid_main.pulse_irm import plan_pulse_irm


def pulse_config(**overrides):
    values = dict(system="Old",calibration_source="test-fixture",axial_enabled=True,transverse_enabled=True,
                  axial_calibrated=True,transverse_calibrated=True,
                  axial_calibration=[[10,10],[100,100]],transverse_calibration=[[20,10],[200,100]],
                  axial_min_mT=1,axial_max_mT=100,transverse_min_mT=1,transverse_max_mT=100,
                  axial_capacitor_max_v=100,transverse_capacitor_max_v=200,
                  control_v_per_capacitor_v=.02,feedback_v_per_capacitor_v=.01,control_max_v=10)
    values.update(overrides)
    return PulseIrmConfig(**values)


class PulseIrmPlanTests(unittest.TestCase):
    def test_capacitor_and_control_voltage_are_independent_from_af_calibration(self):
        axial = plan_pulse_irm(50,"axial",pulse_config())
        transverse = plan_pulse_irm(50,"transverse",pulse_config())
        self.assertEqual((axial.capacitor_v,axial.control_v,axial.fire_hold_s),(50,1,1))
        self.assertEqual((transverse.capacitor_v,transverse.control_v,transverse.fire_hold_s),(100,2,3))
        self.assertEqual(axial.capacitor_tolerance_v,.5)

    def test_asc_boost_is_planned_without_silent_clipping(self):
        plan = plan_pulse_irm(50,"axial",pulse_config(system="ASC"))
        self.assertEqual(plan.charge_control_v,1.31)
        with self.assertRaisesRegex(ValueError,"DAC limit"):
            plan_pulse_irm(50,"axial",pulse_config(system="ASC",control_max_v=1))

    def test_backfield_requires_explicit_polarity_permission(self):
        with self.assertRaisesRegex(ValueError,"polarity"):
            plan_pulse_irm(-50,"axial",pulse_config())
        plan = plan_pulse_irm(-50,"axial",pulse_config(backfield_enabled=True))
        self.assertTrue(plan.backfield)
        self.assertEqual(plan.capacitor_v,50)

    def test_invalid_fields_and_calibration_do_not_mutate_configuration(self):
        for cfg,field in ((pulse_config(),float("nan")),(pulse_config(),101),
                          (pulse_config(axial_calibrated=False),50),(pulse_config(axial_enabled=False),50),
                          (pulse_config(axial_calibration=[[10,10],[9,100]]),50),
                          (pulse_config(feedback_v_per_capacitor_v=0),50),
                          (pulse_config(axial_capacitor_max_v=30),50)):
            previous = asdict(cfg)
            with self.subTest(field=field),self.assertRaises(ValueError):
                plan_pulse_irm(field,"axial",cfg)
            self.assertEqual(asdict(cfg),previous)

    def test_zero_field_requires_wiring_and_feedback_but_no_magnetic_field_table(self):
        cfg = pulse_config(axial_calibrated=False, axial_calibration=[], axial_min_mT=0, axial_max_mT=0)
        plan = plan_pulse_irm(0, "axial", cfg)
        self.assertEqual(plan.operation, "zero_field")
        self.assertEqual((plan.capacitor_v, plan.control_v, plan.charge_control_v), (0, 0, 0))
        with self.assertRaises(ValueError): plan_pulse_irm(0, "axial", pulse_config(axial_enabled=False))
        with self.assertRaises(ValueError): plan_pulse_irm(0, "axial", pulse_config(feedback_v_per_capacitor_v=0))

    def test_station_profile_import_preserves_disabled_modules_and_correct_units(self):
        cfg = AppConfig()
        source = Path(__file__).resolve().parents[3]/"VB6"/"settings"/"Paleomag_v3.INI"
        import_vb6_ini(cfg,source)
        pulse = cfg.pulse_irm
        self.assertEqual(pulse.system,"ASC")
        self.assertFalse(pulse.axial_enabled)
        self.assertFalse(pulse.transverse_enabled)
        self.assertFalse(pulse.backfield_enabled)
        self.assertEqual(pulse.coil_position,-36000)
        self.assertEqual(pulse.axial_max_mT,1200)
        self.assertEqual(pulse.axial_min_mT,6.5)
        self.assertEqual(len(pulse.axial_calibration),25)
        self.assertAlmostEqual(pulse.axial_calibration[0][1],5.46)
        with self.assertRaisesRegex(ValueError,"Enabled"):
            plan_pulse_irm(100,"axial",pulse)
        # Only the fixture is enabled here. This never configures a device.
        pulse.axial_enabled = True
        plan = plan_pulse_irm(100,"axial",pulse)
        self.assertAlmostEqual(plan.capacitor_v,1000/30.125)
        self.assertAlmostEqual(plan.control_v,plan.capacitor_v*.01998)

    def test_pulse_config_survives_app_config_serialization(self):
        cfg = AppConfig(pulse_irm=pulse_config())
        restored = AppConfig._from_dict(asdict(cfg))
        self.assertEqual(asdict(restored.pulse_irm),asdict(cfg.pulse_irm))

    def test_matsusada_charge_scaling_matches_each_source_interval(self):
        for target,factor in ((10,1.5),(30,1.33),(70,1.2),(150,1.1),(250,1.02),(400,1.01)):
            with self.subTest(target=target):
                cfg = pulse_config(system="Matsusada",axial_calibration=[[1,1],[500,500]],
                                   axial_max_mT=500,axial_capacitor_max_v=500)
                plan = plan_pulse_irm(target,"axial",cfg)
                self.assertAlmostEqual(plan.charge_control_v,plan.control_v*factor)
