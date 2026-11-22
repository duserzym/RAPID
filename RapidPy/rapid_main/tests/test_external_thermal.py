"""Operator handoff for externally treated live specimen measurements."""
from datetime import datetime, timezone
from tempfile import TemporaryDirectory
import unittest
from unittest.mock import Mock, patch

from PySide6 import QtWidgets

from rapid_main.app import MainWindow
from rapid_main.dialogs.external_thermal import ExternalThermalDialog
from rapid_main.thermal import ExternalThermalTreatment, compile_thermal_routine
from test_measurement_worker import ReturnAwareBackend


class ExternalThermalDialogTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.qt = QtWidgets.QApplication.instance() or QtWidgets.QApplication([])

    def test_dialog_requires_attestation_and_measured_cooldown(self):
        plan = compile_thermal_routine([100], name="Furnace study")
        dialog = ExternalThermalDialog(None, plan, "SPEC1", "operator")
        self.addCleanup(dialog.deleteLater)
        dialog.step.setCurrentIndex(0)
        dialog.log.setText("Oven 2 / log 384")
        dialog.peak.setValue(99.)
        dialog.hold.setValue(601)
        dialog.handling.setValue(51.)
        self.assertFalse(dialog.measure.isEnabled())
        dialog.confirmed.setChecked(True)
        dialog.accept()
        self.assertIsNone(dialog.treatment)
        self.assertIn("cooldown", dialog.error.text())
        dialog.handling.setValue(30.)
        dialog.accept()
        self.assertEqual(dialog.result(), QtWidgets.QDialog.DialogCode.Accepted)
        self.assertEqual(dialog.treatment.label, "TT100")
        self.assertEqual(dialog.treatment.furnace_log, "Oven 2 / log 384")

    def test_live_measure_start_passes_only_attested_step_and_keeps_plan(self):
        with TemporaryDirectory() as data_dir:
            window = MainWindow()
            self.addCleanup(window.deleteLater)
            window.config.general.data_dir = data_dir
            window.config.general.operator = "operator"
            window.config.general.nocomm = False
            window.set_current_sample("SPEC1")
            plan = compile_thermal_routine([100, 200], name="Furnace study")
            window._install_thermal_plan(plan)
            backend = ReturnAwareBackend()
            receipt = ExternalThermalTreatment("SPEC1", "TT100", "operator", "Oven 2 / log 384",
                99., 601, 30., datetime.now(timezone.utc).isoformat())
            lease = Mock()
            with (patch.object(window, "measurement_backend", return_value=backend),
                  patch.object(window, "acquire_devices", return_value=lease) as acquire,
                  patch("rapid_main.panels.measurement.ExternalThermalDialog") as prompt,
                  patch("rapid_main.panels.measurement.MeasurementWorker") as worker,
                  patch.object(QtWidgets.QMessageBox, "warning", side_effect=AssertionError("Unexpected setup warning"))):
                prompt.return_value.exec.return_value = QtWidgets.QDialog.DialogCode.Accepted
                prompt.return_value.treatment = receipt
                try:
                    self.assertTrue(window._measurement.start_measurement_for_sample("SPEC1"))
                    arguments = worker.call_args.kwargs
                    self.assertEqual(arguments["labels"], ["TT100"])
                    self.assertEqual(arguments["thermal_context"]["queue_labels"], ["TT100", "TT200"])
                    self.assertIs(arguments["external_thermal"], receipt)
                    self.assertEqual(acquire.call_args.args[0], ("measurement", "changer", "af_demag", "squid", "vacuum"))
                    self.assertEqual(window._sequence_labels, plan.labels)
                    worker.return_value.start.assert_called_once()
                finally:
                    window._measurement._worker = None
                    window._measurement._release_measurement_ownership()
                    window.stop_run()
