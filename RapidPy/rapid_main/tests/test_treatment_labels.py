import unittest
from dataclasses import asdict

from PySide6 import QtWidgets
from rapid_main.treatment_labels import parse_field_treatment
from rapid_main.data_model import RockmagStep
from rapid_main.panels.sequence import SequencePanel
from rapid_main.rockmag import rockmag_the_works, compile_rockmag_routine
from tests import test_live_treatment_validation
from tests.af_fakes import configured_af


class TreatmentLabelTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls): cls.app = QtWidgets.QApplication.instance() or QtWidgets.QApplication([])

    def test_explicit_units_and_axis_convert_at_actuator_boundary(self):
        for label, family, field, bias, axis in (
            ("IRM100G", "IRM", 10, None, None),
            ("IRM100mT", "IRM", 100, None, None),
            ("IRMY-100G", "IRM", -10, None, "Y"),
            ("IRMZ0G", "IRM", 0, None, "Z"),
            ("ARM100mT_0.5G", "ARM", 100, .05, None),
            ("ARM100G_0.05mT", "ARM", 10, .05, None),
            ("AF100G", "AF", 10, None, None),
            ("IRM30", "IRM", 30, None, None),
        ):
            with self.subTest(label=label):
                request = parse_field_treatment(label)
                self.assertEqual((request.family, request.field_mT, request.bias_mT, request.axis), (family, field, bias, axis))

    def test_invalid_units_and_bias_on_wrong_family_fail(self):
        for label in ("IRM100T", "IRM100G_5G", "ARMmT", "AF100G_1G", "IRM-BF", "ARM100mT_-1G"):
            with self.subTest(label=label), self.assertRaises(ValueError): parse_field_treatment(label)

    def test_gauss_irm_executes_correct_mT_and_explicit_axis_without_config_mutation(self):
        for label, coil, orientation, requested in (("IRMZ100G", "axial", 0, 10), ("IRMY100G", "transverse", 90, 10), ("IRMX100G", "transverse", 0, 10)):
            with self.subTest(label=label):
                backend = test_live_treatment_validation.LiveTreatmentValidationTests().backend()
                previous = asdict(backend._config)
                self.assertTrue(backend.validate_treatment_plan((label,)).ok)
                backend.set_demag_step(label)
                record = backend.pulse_treatment_records[-1]
                self.assertEqual((record.circuit.plan.field_mT, record.circuit.plan.coil, record.orientation_deg), (requested, coil, orientation))
                self.assertEqual(asdict(backend._config), previous)

    def test_negative_gauss_uses_backfield_gate_and_retains_signed_mT(self):
        backend = test_live_treatment_validation.LiveTreatmentValidationTests().backend()
        self.assertFalse(backend.validate_treatment_plan(("IRMZ-100G",)).ok)
        backend._config.pulse_irm.backfield_enabled = True
        backend.set_demag_step("IRMZ-100G")
        plan = backend.pulse_treatment_records[-1].circuit.plan
        self.assertEqual(plan.field_mT, -10)
        self.assertTrue(plan.backfield)

    def test_arm_peak_and_bias_have_independent_units_and_survive_execution(self):
        backend = test_live_treatment_validation.LiveTreatmentValidationTests().backend()
        backend.set_demag_step("ARM250G_0.5G")
        record = backend.af_treatment_records[-1]
        self.assertEqual(record.plan.passes[0].ramp.field_mT, 25)
        self.assertEqual(record.bias_mT, .05)
        backend._arm_bias.set_bias_mT.assert_called_once_with(.05)

    def test_af_gauss_is_not_replaced_by_default_peak(self):
        from rapid_main.af_treatment import plan_af_treatment
        from rapid_main.diagnostic_services import plan_af_demag_command
        self.assertEqual(plan_af_treatment("AFZ100G", configured_af(), 3000).passes[0].ramp.field_mT, 10)
        self.assertEqual(plan_af_demag_command("AFZ100G", configured_af()).field_mT, 10)

    def test_routine_values_preserve_declared_units_in_model_and_executor(self):
        plan = compile_rockmag_routine(rockmag_the_works(af_fields_mT=(), irm_fields_g=(100,), arm_fields_g=(.5,)))
        self.assertEqual(plan.labels, ["NRM", "IRM100G", "ARM100MT_0.5G", "IRM-100G", "SUSC"])
        self.assertEqual((RockmagStep.from_label("IRM10MT").value, RockmagStep.from_label("IRM10MT").unit), (10, "mT"))
        self.assertEqual(RockmagStep.from_label("IRM-100G").family, "BACKFIELD")

    def test_sequence_rrm_keeps_af_speed_axis_bias_and_signed_rotation(self):
        panel = SequencePanel()
        panel._chk_rrm.setChecked(True)
        panel._rrm_step.setValue(.5)
        panel._rrm_max.setValue(1)
        panel._rrm_af.setValue(50)
        panel._rrm_coil.setCurrentIndex(1)
        panel._chk_rrm_bias.setChecked(True)
        panel._rrm_bias.setValue(.05)
        panel._chk_rrm_neg.setChecked(True)
        self.assertEqual(panel.generate_labels(), ["RRMZ50/0.5@0.05", "RRMZ50/1@0.05", "RRMZ50/-0.5@0.05", "RRMZ50/-1@0.05"])
        panel.deleteLater()

    def test_sequence_arm_uses_displayed_peak_af_and_bias_in_gauss(self):
        panel = SequencePanel()
        panel._chk_arm.setChecked(True)
        panel._arm_step.setValue(.5)
        panel._arm_max.setValue(1)
        panel._arm_af.setValue(25)
        self.assertEqual(panel.generate_labels(), ["ARM25mT_0.5G", "ARM25mT_1G"])
        panel.deleteLater()

    def test_sequence_logarithmic_gauss_limits_are_not_applied_as_mT(self):
        panel = SequencePanel()
        panel._chk_irm.setChecked(True)
        panel._irm_min.setValue(5)
        panel._irm_af_max.setValue(1)
        panel._irm_irm_max.setValue(10)
        panel._irm_log.setValue(1)
        self.assertEqual(panel.generate_labels(), ["AF0.5", "AF1", "IRM5G", "IRM10G"])
        panel._chk_backfield.setChecked(True)
        self.assertEqual(panel.generate_labels()[-2:], ["IRM-5G", "IRM-10G"])
        panel.deleteLater()

    def test_sequence_controls_remain_reachable_in_short_window(self):
        from PySide6 import QtCore
        panel = SequencePanel()
        panel.resize(1100, 700)
        panel.show()
        self.app.processEvents()
        self.assertEqual(panel.height(), 700)
        self.assertGreater(panel._steps_scroll.verticalScrollBar().maximum(), 0)
        panel._steps_scroll.ensureWidgetVisible(panel._chk_susc)
        panel._steps_scroll.horizontalScrollBar().setValue(0)
        self.app.processEvents()
        point = panel._chk_susc.mapTo(panel._steps_scroll.viewport(), QtCore.QPoint(5, panel._chk_susc.height() // 2))
        self.assertTrue(panel._steps_scroll.viewport().rect().contains(point))
        panel.hide()
        panel.deleteLater()

    def test_hawaiian_gauss_preset_executes_2_point_5_to_80_mT_not_25_to_800(self):
        panel = SequencePanel()
        panel._preset_hawaiian()
        fields = [parse_field_treatment(label).field_mT for label in panel.generate_labels()[1:]]
        self.assertEqual(fields, [2.5, 5, 10, 20, 40, 80])
        panel.deleteLater()

    def test_field_formatter_never_rounds_small_request_into_zero(self):
        from rapid_main.treatment_labels import format_field_value
        for value in (.000001, .1234567, 100, 1e-12):
            label = "IRM" + format_field_value(value) + "MT"
            self.assertEqual(parse_field_treatment(label).field_mT, value)
