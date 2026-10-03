from pathlib import Path
import unittest
from unittest.mock import patch

from rapid_main.config import AppConfig
from rapid_main.legacy_ini import import_vb6_ini
from rapid_main.hardware_contracts import _build_motor_controller_config
from rapidpy_common.hardware import HardwareError, MotorAxisConfig, MotorControllerConfig
from rapidpy_common.motor_routing import RoutedMotorSerialClient


class FakeSerial:
    def __init__(self, **kwargs):
        self.port = kwargs["port"]
        self.kwargs = kwargs
        self.is_open = True
        self.writes = []

    def write(self, value):
        self.writes.append(value)

    def flush(self): pass
    def reset_input_buffer(self): pass
    def reset_output_buffer(self): pass
    def close(self): self.is_open = False
    def read_until(self, _terminator): return b"@16 ACK 0000\r"


class MotorRoutingTests(unittest.TestCase):
    def test_incomplete_native_settings_block_preflight_without_constructing_motor_transport(self):
        from rapid_main.hardware_contracts import QueueHardwareBackend
        from unittest.mock import Mock
        cfg = AppConfig()
        cfg.motor_station.calibration_source = "incomplete-import"
        cfg.motor_station.controller = {"turning_motor_full_rotation": -8000}
        with patch.object(QueueHardwareBackend, "_build_component", return_value=Mock()), patch("rapid_main.hardware_contracts.MotorSerialClient") as native:
            backend = QueueHardwareBackend(cfg)
            result = backend.preflight()
        self.assertFalse(result.ok)
        self.assertTrue(any("Incomplete native motor calibration" in reason for reason in result.blockers))
        self.assertIsNone(backend._client)
        native.assert_not_called()

    def test_imported_settings_show_all_ports_and_disable_percentage_speed_controls(self):
        from PySide6 import QtWidgets
        from rapid_main.panels.settings_panel import SettingsPanel
        app = QtWidgets.QApplication.instance() or QtWidgets.QApplication([])
        cfg = AppConfig()
        import_vb6_ini(cfg, Path(__file__).resolve().parents[3] / "VB6/settings/Paleomag_v3.INI")
        panel = SettingsPanel()
        panel.load_from_config(cfg)
        self.assertIn("turning: COM6 @16", panel._motor_station_summary.text())
        self.assertEqual(panel._ch_baud.currentText(), "57600")
        self.assertFalse(panel._ch_speed_z.isEnabled())
        panel.deleteLater()
        app.processEvents()

    def test_imported_calibration_preserves_station_units_and_roundtrips(self):
        cfg = AppConfig()
        import_vb6_ini(cfg, Path(__file__).resolve().parents[3] / "VB6/settings/Paleomag_v3.INI")
        from dataclasses import asdict
        cfg = AppConfig._from_dict(asdict(cfg))
        motor = _build_motor_controller_config(cfg)
        self.assertEqual(motor.turning_motor_full_rotation, -8000)
        self.assertEqual(motor.turning_motor_1rps, 16_000_000)
        self.assertEqual(motor.lift_speed_slow, 25_000_000)
        self.assertEqual(motor.lift_acceleration, 50_000)
        self.assertEqual(motor.meas_pos, -97500)
        self.assertEqual(motor.sample_height, 3000)
        self.assertAlmostEqual(motor.one_step, -1010.1010101)
        self.assertIsInstance(motor.lift_speed_slow, int)
        self.assertEqual(cfg.motor_station.baud, 57600)

    def test_same_address_routes_by_axis_and_evidence_names_actual_port(self):
        axes = [MotorAxisConfig("X", 1, 16, "COM3"), MotorAxisConfig("Turning", 2, 16, "COM6")]
        events = []
        client = RoutedMotorSerialClient(MotorControllerConfig(), axes, trace=lambda *event: events.append(event))
        with patch("rapidpy_common.hardware.serial.Serial", side_effect=FakeSerial), patch("rapidpy_common.hardware.time.sleep"):
            client.connect(baudrate=57600)
            x = client._connections["COM3"]._serial
            turn = client._connections["COM6"]._serial
            client.stop(axes[0])
            client.stop(axis=axes[1])
            self.assertEqual(x.writes[-1], turn.writes[-1])
            self.assertEqual(len(x.writes), 2)
            self.assertEqual(len(turn.writes), 2)
            client.set_torques(axes[0], 28, 28, 28, 28)
            self.assertEqual(x.writes[-1], b"@16 149 8960 8960 8960 8960\r\n")
            self.assertTrue(any("port=COM6" in detail for direction, payload, detail in events if direction == "TX"))
            with self.assertRaises(HardwareError): client.query_ascii("@16 0")
            # A failed port does not prevent stopping a separate healthy axis.
            turn.close()
            self.assertFalse(client.is_connected)
            client.stop(axes[0])
            with self.assertRaises(HardwareError): client.stop(axes[1])
            client.disconnect()
            self.assertFalse(x.is_open)

    def test_shared_port_opens_once_and_duplicate_binding_is_rejected(self):
        axes = [MotorAxisConfig("X", 1, 16, "COM3"), MotorAxisConfig("Turning", 2, 17, "com3")]
        client = RoutedMotorSerialClient(MotorControllerConfig(), axes)
        with patch("rapidpy_common.hardware.serial.Serial", side_effect=FakeSerial) as factory, patch("rapidpy_common.hardware.time.sleep"):
            client.connect()
            self.assertEqual(factory.call_count, 1)
            client.disconnect()
        axes[1].address = 16
        with self.assertRaisesRegex(HardwareError, "Duplicate"): RoutedMotorSerialClient(MotorControllerConfig(), axes)

    def test_partial_connect_failure_closes_other_ports(self):
        axes = [MotorAxisConfig("X", 1, 16, "COM3"), MotorAxisConfig("Turning", 2, 16, "COM6")]
        opened = FakeSerial(port="COM3")
        client = RoutedMotorSerialClient(MotorControllerConfig(), axes)
        with patch("rapidpy_common.hardware.serial.Serial", side_effect=[opened, OSError("offline")]), patch("rapidpy_common.hardware.time.sleep"):
            with self.assertRaises(OSError): client.connect()
        self.assertFalse(opened.is_open)
        self.assertFalse(client.is_connected)

    def test_fractional_negative_measurement_position_uses_vb6_floor(self):
        cfg = AppConfig()
        cfg.motion.zero_pos = -50000
        cfg.motion.meas_pos = -97500
        cfg.motion.sample_top = -2000
        cfg.motion.sample_bottom = -5001
        self.assertEqual(cfg.motion.zero_position(), -48500)
        self.assertEqual(cfg.motion.measurement_position(), -96000)

    def test_stop_spin_requires_stationary_verification_and_propagates_timeout(self):
        from rapidpy_common.hardware import MotorSerialClient
        from unittest.mock import Mock
        client = MotorSerialClient()
        axis = MotorAxisConfig("Turning", 2, 16)
        client.stop = Mock()
        client.wait_for_motor_stop = Mock(side_effect=HardwareError("still moving"))
        client.read_position = Mock(return_value=10)
        with self.assertRaisesRegex(HardwareError, "still moving"): client.turning_motor_spin(axis, 0)
        client.stop.assert_called_once_with(axis)
        client.read_position.assert_not_called()
        client.wait_for_motor_stop.side_effect = None
        self.assertTrue(client.turning_motor_spin(axis, 0).success)

    def test_invalid_native_motion_units_fail_before_any_serial_command(self):
        from rapidpy_common.hardware import MotorSerialClient
        from unittest.mock import Mock
        client = MotorSerialClient()
        client.poll_motor = Mock()
        for target, velocity in ((2**31, 1000), (10, 0), (10, 2**31), (1.5, 1000), (float("nan"), 1000)):
            with self.subTest(target=target, velocity=velocity), self.assertRaises(HardwareError):
                client.move_motor(MotorAxisConfig("Turning", 2, 16), target, velocity)
        client.poll_motor.assert_not_called()

    def test_mutated_axis_binding_is_rejected_after_registration(self):
        axis = MotorAxisConfig("Turning", 2, 16, "COM6")
        client = RoutedMotorSerialClient(MotorControllerConfig(), [axis])
        with patch("rapidpy_common.hardware.serial.Serial", side_effect=FakeSerial), patch("rapidpy_common.hardware.time.sleep"):
            client.connect()
            transport = client._connections["COM6"]._serial
            axis.address = 17
            with self.assertRaisesRegex(HardwareError, "Unregistered"): client.stop(axis)
            self.assertEqual(len(transport.writes), 1)
            client.disconnect()

    def test_spin_reference_normalizes_thousands_of_turns_without_limited_wrapping(self):
        from rapid_main.squid_transport import MotorTurningController
        from rapidpy_common.hardware import MoveResult
        from unittest.mock import Mock
        for rotation, raw, expected in ((-8000, 12_000_111, 111), (8000, -12_000_111, -111)):
            with self.subTest(rotation=rotation):
                client = Mock(config=MotorControllerConfig(turning_motor_full_rotation=rotation))
                axis = MotorAxisConfig("Turning", 2, 16)
                client.read_position.return_value = raw
                client.turning_motor_rotate.return_value = MoveResult(0, 0, True)
                result = MotorTurningController(client, axis).restore_spin_reference()
                self.assertTrue(result.ok)
                self.assertEqual(client.relabel_pos.call_args_list[0].args, (axis, expected))
                self.assertEqual(client.relabel_pos.call_args_list[-1].args, (axis, 0))
                client.turning_motor_rotate.assert_called_once_with(axis, 0., wait_for_stop=True)
