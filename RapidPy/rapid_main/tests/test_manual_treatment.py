import hashlib
import json
from pathlib import Path
import tempfile
import threading
import time
import unittest
from unittest.mock import Mock, patch

from PySide6 import QtWidgets
from rapid_main.manual_treatment import ManualArmTreatment
from rapid_main.dialogs.irm_arm import IrmArmDialog
from rapid_main.diagnostic_services import IrmArmNoCommBackend
from tests import test_live_treatment_validation


class ManualTreatmentTests(unittest.TestCase):
    def backend(self):
        return test_live_treatment_validation.LiveTreatmentValidationTests().backend()

    def test_manual_arm_publishes_current_attempt_and_restores_queue_context(self):
        backend = self.backend()
        previous = (backend._sample_name,backend._run_id,backend._treatment_label,backend._halt_check)
        with tempfile.TemporaryDirectory() as directory:
            service = ManualArmTreatment(backend,directory)
            message = service.apply(sample_id="S2",peak_af_mT=25,bias_mT=.5,should_cancel=lambda:False)
            self.assertIn("completed",message)
            folders = list(Path(directory).iterdir())
            self.assertEqual(len(folders),1)
            index = json.loads((folders[0]/"artifact_index.json").read_text())
            self.assertEqual(index["state"],"completed")
            self.assertEqual(index["sample_id"],"S2")
            artifact = folders[0]/index["artifacts"][0]["relative_path"]
            self.assertEqual(index["artifacts"][0]["sha256"],hashlib.sha256(artifact.read_bytes()).hexdigest())
            record = json.loads(artifact.read_text())
            self.assertEqual(record["plan"]["passes"][0]["ramp"]["field_mT"],25)
            self.assertEqual(record["bias_mT"],.5)
        self.assertEqual((backend._sample_name,backend._run_id,backend._treatment_label,backend._halt_check),previous)

    def test_failed_treatment_publishes_failure_and_bias_cleanup_evidence(self):
        backend = self.backend()
        backend._af_demag.apply_calibrated_af.side_effect = RuntimeError("ramp lost")
        with tempfile.TemporaryDirectory() as directory:
            with self.assertRaisesRegex(RuntimeError,"ramp lost"):
                ManualArmTreatment(backend,directory).apply(sample_id="S2",peak_af_mT=25,bias_mT=.5,should_cancel=lambda:False)
            folder = next(Path(directory).iterdir())
            index = json.loads((folder/"artifact_index.json").read_text())
            self.assertEqual(index["state"],"failed")
            record = json.loads((folder/index["artifacts"][0]["relative_path"]).read_text())
            self.assertTrue(record["safe_state_confirmed"])
            self.assertIn("ramp lost",record["error"])

    def test_publication_failure_cannot_claim_success(self):
        backend = self.backend()
        with tempfile.TemporaryDirectory() as directory, patch("rapid_main.manual_treatment.write_af_treatment_record",side_effect=OSError("disk full")):
            with self.assertRaisesRegex(RuntimeError,"publication failed.*disk full"):
                ManualArmTreatment(backend,directory).apply(sample_id="S2",peak_af_mT=25,bias_mT=.5,should_cancel=lambda:False)
            index = json.loads((next(Path(directory).iterdir())/"artifact_index.json").read_text())
            self.assertEqual(index["state"],"failed")

    def test_invalid_input_or_unwritable_destination_precedes_motion(self):
        backend = self.backend()
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)/"file"
            root.write_text("occupied")
            service = ManualArmTreatment(backend,root)
            for sample,peak in (("",25),("S1",float("nan"))):
                with self.assertRaises(ValueError):
                    service.apply(sample_id=sample,peak_af_mT=peak,bias_mT=.5,should_cancel=lambda:False)
            with self.assertRaises(OSError):
                service.apply(sample_id="S1",peak_af_mT=25,bias_mT=.5,should_cancel=lambda:False)
        backend._af_demag.apply_calibrated_af.assert_not_called()
        backend._arm_bias.set_bias_mT.assert_not_called()

    def test_manual_pulse_uses_actual_circuit_and_restores_axis(self):
        backend = self.backend()
        old_axis = backend._config.irm_arm.irm_axis
        with tempfile.TemporaryDirectory() as directory:
            message = ManualArmTreatment(backend,directory).apply_irm(sample_id="S1",max_field_mT=30,
                                axis="Z (up-axis)",should_cancel=lambda:False)
            self.assertIn("IRM completed",message)
            folder = next(Path(directory).iterdir())
            index = json.loads((folder/"artifact_index.json").read_text())
            record = json.loads((folder/index["artifacts"][0]["relative_path"]).read_text())
            self.assertEqual(record["schema"],"rapidpy.irm.treatment.v1")
            self.assertTrue(record["safe_state_confirmed"])
            self.assertTrue(record["circuit"]["fired"])
            self.assertEqual(record["circuit"]["plan"]["capacitor_v"],30)
        self.assertEqual(backend._config.irm_arm.irm_axis,old_axis)

    def test_unsafe_pulse_blocks_new_treatments_and_motor_commands(self):
        backend = self.backend()
        backend._pulse_treatment_records.append(Mock(safe_state_confirmed=False))
        self.assertFalse(backend.validate_treatment_plan(("AF20",)).ok)
        with self.assertRaisesRegex(RuntimeError,"motion is inhibited"):
            backend._ensure_connected()

    def test_manual_reset_publishes_discharge_recovery_evidence(self):
        from rapidpy_common.hardware import MotorAxisConfig
        backend = self.backend()
        backend._pulse_treatment_records.append(Mock(safe_state_confirmed=False))
        backend._sample_loaded = False
        backend._axes.update(changer_x=MotorAxisConfig("ChangerX",1,1),changer_y=MotorAxisConfig("ChangerY",4,4))
        backend._motor_communication_logger = Mock()
        with tempfile.TemporaryDirectory() as directory:
            message = ManualArmTreatment(backend,directory).reset()
            self.assertIn("safe-state recovery completed",message)
            folder = next(Path(directory).iterdir())
            index = json.loads((folder/"artifact_index.json").read_text())
            record = json.loads((folder/index["artifacts"][0]["relative_path"]).read_text())
            self.assertTrue(record["safe_state_confirmed"])
            self.assertFalse(record["fired"])
            self.assertEqual(record["schema"],"rapidpy.irm.pulse.v1")
            self.assertFalse(backend.has_unresolved_pulse_fault)


class ManualDialogTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.app = QtWidgets.QApplication.instance() or QtWidgets.QApplication([])

    def wait_done(self, dialog):
        end = time.monotonic()+3
        while dialog._task is not None and time.monotonic()<end:
            self.app.processEvents()
            time.sleep(.005)
        self.assertIsNone(dialog._task)

    def test_manual_arm_runs_off_gui_thread_and_close_waits_for_cleanup(self):
        main_thread = threading.get_ident()
        observed = []
        cleanup = threading.Event()
        def operation(**kwargs):
            observed.append(threading.get_ident())
            end = time.monotonic()+2
            while not kwargs["should_cancel"]() and time.monotonic()<end:
                time.sleep(.005)
            cleanup.set()
            return "cancelled and cleaned"
        service = Mock()
        service.apply.side_effect = operation
        dialog = IrmArmDialog(backend=IrmArmNoCommBackend(),manual_arm=service)
        dialog._mode.setCurrentIndex(1)
        dialog._sample_id.setText("S1")
        dialog._apply()
        self.assertIsNotNone(dialog._task)
        self.assertFalse(dialog._apply_btn.isEnabled())
        dialog.reject()
        self.assertTrue(dialog._close_pending)
        self.wait_done(dialog)
        self.assertTrue(cleanup.is_set())
        self.assertNotEqual(observed[0],main_thread)
        service.apply.assert_called_once()
        dialog.deleteLater()
