"""Transverse scan sequencing, failure recovery, and operator controls."""
from dataclasses import replace
import json
from pathlib import Path
from tempfile import TemporaryDirectory
import unittest
from unittest.mock import Mock, patch

from PySide6 import QtCore, QtTest, QtWidgets

from rapid_main.config import AppConfig
from rapid_main.diagnostic_services import AfDemagBackendAdapter, DCMotorBackendAdapter
from rapid_main.panels.transverse_scan import TransverseScanDialog
from rapid_main.transverse import plan_transverse_scan
from rapid_main.transverse_scan import NativeTransverseScan, SimulatedTransverseScan
from rapidpy_common.adwin_af import AdwinBoardConfig, AdwinCoilLimits, AdwinRampResult
from rapidpy_common.gaussmeter import GaussmeterReading
from rapidpy_common.hardware import HardwareError, MotorControllerConfig, MoveResult, convert_angle_to_pos
from rapidpy_common.hardware_safety import HardwareSafetyError, HardwareSafetyStore

from test_af_treatment import configured_af


class NativeTransverseScanTests(unittest.TestCase):
    def setUp(self):
        directory = TemporaryDirectory()
        self.addCleanup(directory.cleanup)
        self.store = HardwareSafetyStore(Path(directory.name) / "safety.json")
        self.config = AppConfig()
        self.config.af_demag = configured_af()
        self.config.motor_station.lift_counts_per_cm = 1000
        self.config.general.operator = "Test Operator"
        self.motors = DCMotorBackendAdapter(port="COM3", baud=9600)
        self.motors._station_baud = 9600
        self.motors._connected = True
        self.motors._safety_store = self.store
        client = self.motors._client = Mock(config=MotorControllerConfig(), is_connected=True)
        self.actual = 0.
        self.height = self.config.af_demag.coil_position
        self.events = []
        self.cancel = False
        self.stop_failure = False
        self.af_failure = False

        def read(axis, registers):
            if axis.name == "UpDown":
                return self.height, 0
            return convert_angle_to_pos(self.actual, client.config.turning_motor_full_rotation), int(self.stop_failure)

        def rotate(axis, angle):
            self.assertEqual(axis.name, "Turning")
            self.assertIsNotNone(self.store.pending())
            self.actual = angle
            self.events.append(("move", angle))
            position, _ = read(axis, None)
            return MoveResult(position, position, True)

        client.read_registers.side_effect = read
        client.turning_motor_rotate.side_effect = rotate
        controller = Mock(board=AdwinBoardConfig(), limits=AdwinCoilLimits())
        controller.last_safe_field_recovery = {
            "outputs_zero_confirmed": True, "safe_state_confirmed": True, "simulated": False,
        }
        with patch("rapid_main.diagnostic_services._verify_adwin_controller", return_value=1):
            self.af = AfDemagBackendAdapter(self.config.af_demag, controller=controller)

        def safe_field():
            self.events.append(("field_off", self.actual))
            if self.af_failure:
                raise RuntimeError("relay readback failure")

        def ramp(request, *, should_cancel):
            self.assertIsNotNone(self.store.pending())
            self.assertEqual(request.active_coil, "transverse")
            self.assertFalse(should_cancel())
            self.events.append(("ramp", self.actual))
            return AdwinRampResult(100, 100, 20, 35, 1., 1.25, .1, .001, 10.)

        controller.recover_safe_field.side_effect = safe_field
        controller.run_ramp.side_effect = ramp
        self.meter = Mock()
        self.meter.reset_peak.side_effect = lambda: self.events.append(("reset", self.actual))
        readings = iter([0., -12., 0., 4., 0., 5.])
        self.meter.read.side_effect = lambda **_: GaussmeterReading(next(readings), 0., 1, "G", "G", 1, "DC Peak", 3, None)
        self.plan = plan_transverse_scan(0., 90., 45., 50., 14)
        self.station = NativeTransverseScan(self.config, self.motors, self.af,
                                            meter_factory=lambda **_: self.meter, wait=lambda *_: None)
        self.options = dict(cancelled=lambda: self.cancel, progress=lambda _: None)

    def test_signed_peak_and_separate_bounded_apply(self):
        result = self.station.scan(self.plan, **self.options)
        self.assertEqual(self.actual, 90.)
        self.assertAlmostEqual(result.best.peak_mT, -1.2)
        record = json.loads(Path(result.evidence_path).read_text())
        self.assertTrue(record["safe_state_confirmed"])
        self.assertEqual(record["operation"]["ramp"]["coil"], "transverse")
        ramp_events = [i for i, event in enumerate(self.events) if event[0] == "ramp"]
        self.assertEqual(len(ramp_events), 3)
        for index in ramp_events:
            self.assertEqual(self.events[index - 1][0], "reset")
            self.assertEqual(self.events[index + 1][0], "field_off")
        with self.assertRaisesRegex(HardwareSafetyError, "matches"):
            self.station.apply(replace(result, best=result.samples[1]), **self.options)
        self.assertEqual(self.station.apply(result, **self.options), 0.)
        self.assertEqual([event for event in self.events if event[0] == "ramp"],
                         [("ramp", 0.), ("ramp", 45.), ("ramp", 90.)])

    def test_rejects_wrong_lift_height_and_unbounded_start_before_motion(self):
        self.height -= 1001
        with self.assertRaisesRegex(HardwareError, "AF height"):
            self.station.scan(self.plan, **self.options)
        self.height += 1001
        self.actual = 360.
        with self.assertRaisesRegex(HardwareError, "encoder interval"):
            self.station.scan(self.plan, **self.options)
        self.motors._client.turning_motor_rotate.assert_not_called()
        self.meter.connect.assert_not_called()
        self.assertIsNone(self.store.pending())

    def test_cancelled_acquisition_checks_safe_state_without_next_ramp(self):
        self.station.wait = lambda seconds, _: setattr(self, "cancel", seconds == .5)
        with self.assertRaisesRegex(HardwareError, "stopped"):
            self.station.scan(self.plan, **self.options)
        self.af._controller.run_ramp.assert_called_once()
        self.assertTrue(self.store.read()["record"]["safe_state_confirmed"])
        self.meter.disconnect.assert_called_once()

    def test_recovery_requires_both_channels_and_never_replays_motion(self):
        self.af_failure = self.stop_failure = True
        with self.assertRaisesRegex(HardwareSafetyError, "recovery required"):
            self.station.scan(self.plan, **self.options)
        self.assertIsNotNone(self.store.pending())
        self.af_failure = False
        with self.assertRaisesRegex(HardwareSafetyError, "Motor stop"):
            self.station.recover(**self.options)
        self.assertIsNotNone(self.store.pending())
        self.stop_failure = False
        self.station.recover(**self.options)
        self.assertIsNone(self.store.pending())
        self.motors._client.turning_motor_rotate.assert_not_called()
        self.af._controller.run_ramp.assert_not_called()

    def test_recovery_reopens_disconnected_motor_ports_for_stop_only(self):
        self.motors._axes["Turning"].port = "COM5"
        self.motors._axes["Up/Down"].port = "COM6"
        self.motors._station_baud = 57600
        client = self.motors._client
        self.motors._connected = False
        client._connections = {}
        client.connect_for_stop.return_value = ({"COM5", "COM6"}, [])
        self.store.begin("station_diagnostic", {"action": "transverse_probe_scan"}, self.station._profile())
        self.assertEqual(self.store.pending()["profile"]["resources"]["motors"]["baud"], 57600)
        self.motors._station_baud = 115200
        with self.assertRaisesRegex(HardwareSafetyError, "profile"):
            self.station.recover(**self.options)
        client.connect_for_stop.assert_not_called()
        self.motors._station_baud = 57600
        self.station.recover(**self.options)
        client.connect.assert_not_called()
        client.connect_for_stop.assert_called_once()
        client.turning_motor_rotate.assert_not_called()
        client.disconnect.assert_called_once()
        self.assertIsNone(self.store.pending())

    def test_range_change_stops_the_scan_and_retains_signed_reading(self):
        self.meter.read.side_effect = [
            GaussmeterReading(0., 0., 1, "G", "G", 1, "DC Peak", 3, None),
            GaussmeterReading(-2., -2., 1, "G", "G", 1, "DC Peak", 2, None),
        ]
        with self.assertRaisesRegex(HardwareError, "range or units changed"):
            self.station.scan(self.plan, **self.options)
        observations = self.store.read()["record"]["observations"]
        self.assertEqual(next(item["dc_peak"]["peak_mT"] for item in observations if "dc_peak" in item), -.2)
        self.assertIsNone(self.store.pending())


class TransverseDialogTests(unittest.TestCase):
    def test_simulated_scan_reports_signed_peak_and_applies_separately(self):
        app = QtWidgets.QApplication.instance() or QtWidgets.QApplication([])
        dialog = TransverseScanDialog(None, SimulatedTransverseScan(), lambda **_: None, simulated=True)
        self.addCleanup(dialog.close)
        self.assertFalse(dialog.apply_button.isEnabled())
        dialog.start_button.click()
        for _ in range(100):
            if not dialog.active:
                break
            QtTest.QTest.qWait(20)
        self.assertTrue(dialog.apply_button.isEnabled(), dialog.message.text())
        self.assertGreater(dialog.table.rowCount(), 2)
        self.assertIn("Simulated peak", dialog.message.text())
        dialog.apply_button.click()
        for _ in range(30):
            if not dialog.active:
                break
            QtCore.QCoreApplication.processEvents()
            QtTest.QTest.qWait(10)
        self.assertIn("Simulated angle", dialog.message.text())
        self.assertFalse(dialog.apply_button.isEnabled())
        app.processEvents()
