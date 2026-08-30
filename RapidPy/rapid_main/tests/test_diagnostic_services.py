from __future__ import annotations

import unittest
from unittest import mock

from rapidpy_common.hardware import MotorAxisConfig, MotorTelemetry

from rapid_main.config import AfDemagConfig, IrmArmConfig, SquidConfig, VacuumConfig
from rapid_main import diagnostic_services
from rapid_main.diagnostic_services import (
    DCMotorNoCommBackend,
    AfDemagNoCommBackend,
    build_dcmotor_backend,
    collect_diagnostic_status,
    VacuumNoCommBackend,
    IrmArmNoCommBackend,
    SquidNoCommBackend,
    read_vacuum_snapshot,
    read_squid_snapshot,
    plan_af_demag_command,
    require_squid_ready,
    require_vacuum_ready,
    build_af_demag_backend,
    build_vacuum_backend,
    build_irm_arm_backend,
    build_squid_backend,
)


class TestDiagnosticServices(unittest.TestCase):
    def test_af_demag_command_planning_and_no_comm_recording(self) -> None:
        cfg = AfDemagConfig(peak=200.0, ramp_speed="Slow (3 Hz)", settle=2.5, tumble=True)
        command = plan_af_demag_command("AF50", cfg)

        self.assertEqual(command.field_mT, 50.0)
        self.assertEqual(command.sine_freq_hz, 3.0)
        self.assertEqual(command.settle_s, 2.5)
        self.assertTrue(command.tumble)
        self.assertAlmostEqual(command.ramp_peak_voltage, 2.5)

        backend = AfDemagNoCommBackend(cfg)
        status = backend.apply_af(command)

        self.assertIn("AF requested", status)
        self.assertEqual(backend.commands, [command])
        self.assertIn("reset", backend.reset_field().lower())

    def test_vacuum_no_comm_pump_and_pressure_update(self) -> None:
        cfg = VacuumConfig(target_pressure=4.0, warn_threshold=5.0, poll_interval=0.1)
        backend = VacuumNoCommBackend(cfg)

        self.assertTrue(backend.is_connected())
        self.assertEqual(backend.status(), "No-comm simulation")
        self.assertFalse(backend.is_pump_on())

        before = backend.read_pressure()
        backend.set_pump(True)
        mid = backend.read_pressure()
        backend.set_pump(False)
        after = backend.read_pressure()

        self.assertGreaterEqual(before, 0.0)
        self.assertGreaterEqual(mid, 0.0)
        self.assertGreaterEqual(after, 0.0)

    def test_vacuum_snapshot_flags_pressure_faults_for_queue_guard(self) -> None:
        class _HighPressureVacuum:
            def is_connected(self) -> bool:
                return True

            def read_pressure(self) -> float:
                return 250.0

            def set_pump(self, on: bool) -> None:
                del on

            def is_pump_on(self) -> bool:
                return True

            def status(self) -> str:
                return "pressure high"

        snapshot = read_vacuum_snapshot(_HighPressureVacuum(), warn_threshold=20.0)

        self.assertTrue(snapshot.fault)
        self.assertEqual(snapshot.pressure_mtorr, 250.0)
        self.assertTrue(snapshot.pump_on)
        self.assertIn("exceeds", snapshot.fault_reason)
        with self.assertRaisesRegex(diagnostic_services.DiagnosticContractError, "exceeds"):
            require_vacuum_ready(_HighPressureVacuum(), warn_threshold=20.0)

    def test_irm_arm_no_comm_records_state_and_status(self) -> None:
        cfg = IrmArmConfig(irm_max_field=10.0, irm_axis="X", irm_ramp="Slow (60 s)")
        backend = IrmArmNoCommBackend(cfg)

        self.assertEqual(backend.status(), "No-comm IRM/ARM simulator")
        irm_status = backend.apply_irm(max_field_mT=150.0, axis="Y", ramp_label="Fast (10 s)", steps=4)
        self.assertIn("150.00 mT", irm_status)
        self.assertIn("Y", irm_status)
        self.assertEqual(cfg.irm_max_field, 150.0)
        self.assertEqual(cfg.irm_axis, "Y")
        self.assertIn("Fast (10 s)", irm_status)

        arm_status = backend.apply_arm(peak_af_mT=80.0, bias_mT=0.2, steps=3)
        self.assertIn("ARM requested", arm_status)
        self.assertEqual(cfg.arm_peak_af, 80.0)
        self.assertEqual(cfg.arm_bias, 0.2)
        self.assertEqual(cfg.irm_steps, 3)

        reset = backend.reset_field()
        self.assertIn("field reset", reset)

    def test_irm_dac_planning_blocks_over_limit_without_clamping(self) -> None:
        command = diagnostic_services._plan_irm_dac_command(1500.0, "Z (up-axis)")

        self.assertFalse(command.allowed)
        self.assertEqual(command.channel_number, 0)
        self.assertAlmostEqual(command.voltage_v, 15.0)
        self.assertIn("outside 0..10 V", command.reason)
        with self.assertRaisesRegex(diagnostic_services.HardwareError, "BLOCKED"):
            diagnostic_services._field_mT_to_volts(1500.0, "Z (up-axis)")

    def test_irm_dac_planning_maps_transverse_axis_to_distinct_channel(self) -> None:
        command = diagnostic_services._plan_irm_dac_command(250.0, "X")

        self.assertTrue(command.allowed)
        self.assertEqual(command.channel_number, 1)
        self.assertAlmostEqual(command.voltage_v, 2.5)

    def test_irm_dac_planning_uses_configured_voltage_calibration(self) -> None:
        cfg = IrmArmConfig(
            irm_voltage_slope=0.002,
            irm_voltage_intercept=0.25,
            irm_max_voltage=2.0,
        )

        command = diagnostic_services._plan_irm_dac_command(500.0, "Z (up-axis)", cfg)

        self.assertTrue(command.allowed)
        self.assertAlmostEqual(command.voltage_v, 1.25)
        blocked = diagnostic_services._plan_irm_dac_command(1000.0, "Z (up-axis)", cfg)
        self.assertFalse(blocked.allowed)
        self.assertAlmostEqual(blocked.voltage_v, 2.25)

    def test_squid_no_comm_test_connection_sets_connected(self) -> None:
        backend = SquidNoCommBackend(SquidConfig(port="COM7", baud=19200))
        self.assertFalse(backend.is_connected())
        self.assertTrue(backend.test_connection())
        self.assertTrue(backend.is_connected())
        self.assertIn("COM7", backend.status())

    def test_squid_snapshot_flags_disconnected_transport_for_queue_guard(self) -> None:
        class _DisconnectedSquid:
            def is_connected(self) -> bool:
                return False

            def test_connection(self) -> bool:
                return False

            def status(self) -> str:
                return "SQUID transport not connected"

            def read_squid(self) -> tuple[float, float, float]:
                return (0.0, 0.0, 0.0)

            def read_susceptibility(self) -> float:
                return 0.0

        snapshot = read_squid_snapshot(_DisconnectedSquid())

        self.assertTrue(snapshot.fault)
        self.assertFalse(snapshot.connected)
        self.assertIn("not connected", snapshot.fault_reason.lower())
        with self.assertRaisesRegex(diagnostic_services.DiagnosticContractError, "not connected"):
            require_squid_ready(_DisconnectedSquid())

    def test_diagnostic_factory_returns_no_comm_in_nocomm_mode(self) -> None:
        self.assertIsInstance(build_vacuum_backend(VacuumConfig(), nocomm=True), VacuumNoCommBackend)
        self.assertIsInstance(build_af_demag_backend(AfDemagConfig(), nocomm=True), AfDemagNoCommBackend)
        self.assertIsInstance(build_irm_arm_backend(IrmArmConfig(), nocomm=True), IrmArmNoCommBackend)
        self.assertIsInstance(build_squid_backend(SquidConfig(), nocomm=True), SquidNoCommBackend)

    def test_collect_diagnostic_status_summarizes_backend_readiness(self) -> None:
        class _Backend:
            simulated = True

            def is_connected(self) -> bool:
                return False

            def status(self) -> str:
                return "No-comm test backend"

        lines = collect_diagnostic_status({"Test": _Backend()})

        self.assertEqual(len(lines), 1)
        self.assertEqual(lines[0].name, "Test")
        self.assertFalse(lines[0].connected)
        self.assertTrue(lines[0].simulated)
        self.assertEqual(lines[0].level, "WARNING")
        self.assertIn("No-comm test backend", lines[0].format_for_console())

    def test_collect_diagnostic_status_surfaces_backend_read_failures(self) -> None:
        class _BrokenBackend:
            simulated = False

            def is_connected(self) -> bool:
                raise RuntimeError("port unavailable")

            def status(self) -> str:
                raise RuntimeError("status unavailable")

        line = collect_diagnostic_status({"Broken": _BrokenBackend()})[0]

        self.assertTrue(line.fault)
        self.assertEqual(line.level, "ERROR")
        self.assertIn("port unavailable", line.fault_reason)

    def test_build_irm_arm_factory_prefers_adapter_when_available(self) -> None:
        class _DummyAdapter:
            def __init__(self, cfg: IrmArmConfig | None = None) -> None:
                del cfg

            def is_connected(self) -> bool:
                return True

            def status(self) -> str:
                return "dummy adapter"

            def apply_irm(self, *, max_field_mT: float, axis: str, ramp_label: str, steps: int) -> str:
                return f"IRM dummy {max_field_mT} {axis} {ramp_label} {steps}"

            def apply_arm(self, *, peak_af_mT: float, bias_mT: float, steps: int | None = None) -> str:
                return f"ARM dummy {peak_af_mT} {bias_mT}"

            def reset_field(self) -> str:
                return "dummy reset"

        with mock.patch("rapid_main.diagnostic_services.IrmArmBackendAdapter", _DummyAdapter):
            backend = build_irm_arm_backend(IrmArmConfig(), nocomm=False)

        self.assertEqual(backend.status(), "dummy adapter")

    def test_build_irm_arm_factory_falls_back_to_no_comm_on_adapter_failure(self) -> None:
        class _BrokenAdapter:
            def __init__(self, cfg: IrmArmConfig) -> None:
                del cfg
                raise RuntimeError("adapter not available")

        with mock.patch("rapid_main.diagnostic_services.IrmArmBackendAdapter", _BrokenAdapter):
            backend = build_irm_arm_backend(IrmArmConfig(), nocomm=False)

        self.assertIsInstance(backend, IrmArmNoCommBackend)

    def test_dc_motor_no_comm_simulator_reports_motion_and_state(self) -> None:
        backend = DCMotorNoCommBackend(port="COM1", baud=9600)
        self.assertFalse(backend.is_connected())
        backend.connect("COM1", 9600)
        self.assertTrue(backend.is_connected())
        self.assertIn("Connected (sim)", backend.status())

        moved = backend.move_motor("Changer (X)", target=1234, speed=800, wait_for_stop=True)
        self.assertEqual(moved[0], 1234)
        self.assertTrue(moved[2])

        spinning = backend.spin_turning(speed_rps=1.5, duration_s=10.0)
        self.assertTrue(spinning[2])

        hole = backend.goto_hole(hole=42.0)
        self.assertEqual(hole, 42.0)

        backend.home_to_top()
        self.assertIn("sim", backend.status().lower())

        x, y = backend.home_xy_to_center()
        self.assertEqual(x, 0)
        self.assertEqual(y, 0)

        pickup = backend.sample_pickup()
        dropoff = backend.sample_dropoff()
        self.assertTrue(pickup[2])
        self.assertTrue(dropoff[2])

    def test_dc_motor_factory_returns_no_comm_for_nocomm_and_fallback(self) -> None:
        self.assertIsInstance(
            build_dcmotor_backend(port="COM1", baud=9600, nocomm=True),
            DCMotorNoCommBackend,
        )

        # With a non-connectable serial port, fallback should remain deterministic and safe.
        with mock.patch("rapid_main.diagnostic_services.DCMotorBackendAdapter", side_effect=RuntimeError("boom")):
            backend = build_dcmotor_backend(port="COM_DOES_NOT_EXIST", baud=9600, nocomm=False)

        self.assertIsInstance(backend, DCMotorNoCommBackend)
        self.assertFalse(backend.is_connected())

    def test_dc_motor_no_comm_reads_telemetry(self) -> None:
        backend = DCMotorNoCommBackend(port="COM1", baud=9600)
        backend.connect("COM1", 9600)
        backend.move_motor("Turning", target=15000, speed=1200, wait_for_stop=True)
        sample = backend.read_telemetry("Turning")

        self.assertIsInstance(sample, MotorTelemetry)
        self.assertEqual(sample.axis_name, "Turning")
        self.assertGreaterEqual(sample.timestamp, 0.0)

    def test_dc_motor_adapter_telemetry_contract(self) -> None:
        class _FakeMotorSerialClient:
            def __init__(self) -> None:
                self.connected = False

            def connect(self, port: str, baudrate: int = 9600) -> None:
                self.connected = True
                self.port = port
                self.baudrate = baudrate

            def disconnect(self) -> None:
                self.connected = False

            @property
            def is_connected(self) -> bool:
                return self.connected

            def move_motor(self, *args, **kwargs) -> object:
                return mock.Mock(target=42, final_position=43, success=True)

            def turning_motor_spin(self, *args, **kwargs) -> object:
                return mock.Mock(target=42, final_position=40, success=True)

            def changer_motor_to_hole(self, *args, **kwargs) -> object:
                return mock.Mock(final_position=444, target=444, success=True)

            def read_position(self, *args, **kwargs) -> int:
                return 0

            def home_to_top(self, *args, **kwargs) -> object:
                return mock.Mock(target=0, final_position=0, success=True)

            def home_xy_to_center(self, *args, **kwargs) -> tuple[object, object]:
                return (mock.Mock(final_position=1), mock.Mock(final_position=2))

            def move_xy_to_corner(self, *args, **kwargs) -> tuple[object, object]:
                return (mock.Mock(final_position=3), mock.Mock(final_position=4))

            def sample_pickup(self, *args, **kwargs) -> object:
                return mock.Mock(target=1, final_position=1, success=True)

            def sample_dropoff(self, *args, **kwargs) -> object:
                return mock.Mock(target=2, final_position=2, success=True)

            def read_telemetry(self, axis: MotorAxisConfig) -> MotorTelemetry:
                if axis.name != "Turning":
                    raise ValueError("unexpected axis")
                return MotorTelemetry(
                    timestamp=123.45,
                    axis_name=axis.name,
                    target_position=11_000,
                    actual_position=11_001,
                    position_error=1,
                    velocity_1=12,
                    velocity_2=3,
                    actual_torque=250,
                )

        with mock.patch("rapid_main.diagnostic_services.MotorSerialClient", _FakeMotorSerialClient):
            backend = build_dcmotor_backend(port="COM1", baud=115200, nocomm=False)
            backend.connect("COM1", 115200)
            sample = backend.read_telemetry("Turning")

        self.assertIsInstance(sample, MotorTelemetry)
        self.assertEqual(sample.axis_name, "Turning")
        self.assertEqual(sample.actual_torque, 250)
        self.assertEqual(sample.velocity_1, 12)
