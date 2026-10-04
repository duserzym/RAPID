from __future__ import annotations

import unittest
from types import SimpleNamespace
from unittest import mock

from rapidpy_common.hardware import HardwareError, MotorAxisConfig, MotorTelemetry

from rapid_main.config import (
    AfDemagConfig,
    IrmArmConfig,
    SquidConfig,
    SusceptibilityConfig,
    VacuumConfig,
)
from rapid_main import diagnostic_services
from tests.af_fakes import configured_af
from rapid_main.communication_log import CommunicationDirection
from rapid_main.diagnostic_services import (
    DCMotorNoCommBackend,
    AfDemagBackendAdapter,
    AfDemagNoCommBackend,
    build_dcmotor_backend,
    collect_diagnostic_status,
    VacuumNoCommBackend,
    VacuumBackendAdapter,
    IrmArmNoCommBackend,
    IrmArmBackendAdapter,
    SquidNoCommBackend,
    read_vacuum_snapshot,
    read_squid_snapshot,
    plan_af_demag_command,
    require_squid_ready,
    require_vacuum_ready,
    build_af_demag_backend,
    build_backend_or_unavailable,
    build_vacuum_backend,
    build_irm_arm_backend,
    build_squid_backend,
    HardwareUnavailableError,
    UnavailableBackend,
)


class TestDiagnosticServices(unittest.TestCase):
    @staticmethod
    def _adwin_result() -> SimpleNamespace:
        return SimpleNamespace(
            out_count=120,
            in_count=118,
            up_count=40,
            down_start=80,
            monitor_peak_v=2.5,
            ramp_peak_v=5.0,
            down_slope_vps=1.25,
            timestep_s=0.001,
            points_per_period=100.0,
        )

    def test_af_live_adapter_records_request_result_and_confirmed_reset(self) -> None:
        result = self._adwin_result()

        class _Controller:
            def __init__(self) -> None:
                self.requests = []

            def test_version(self) -> int:
                return 91

            def run_ramp(self, request):
                self.requests.append(request)
                return result

            def set_af_relays(self, active_coil: str, one_chan_on: bool = True) -> int:
                self.relay_call = (active_coil, one_chan_on)
                return 0

        controller = _Controller()
        cfg = configured_af(board=2)
        backend = AfDemagBackendAdapter(cfg, controller=controller)
        command = plan_af_demag_command("AF50", cfg)

        status = backend.apply_af(command)
        reset_status = backend.reset_field()

        self.assertIn("AF completed", status)
        self.assertEqual(reset_status, "AF field reset confirmed.")
        self.assertEqual(controller.relay_call, ("off", True))
        events = backend.communication_events()
        self.assertEqual(
            [event.direction for event in events],
            [
                CommunicationDirection.TX,
                CommunicationDirection.RX,
                CommunicationDirection.TX,
                CommunicationDirection.RX,
            ],
        )
        self.assertEqual(events[0].channel, "ADWIN_AF")
        self.assertEqual(events[0].port, "board:2")
        self.assertIn('"action":"run_ramp"', events[0].payload)
        self.assertIn('"active_coil":"axial"', events[0].payload)
        self.assertIn('"action":"run_ramp_result"', events[1].payload)
        self.assertIn('"monitor_peak_v":2.5', events[1].payload)
        self.assertIn('"action":"set_af_relays"', events[2].payload)
        self.assertIn('"relay_word":0', events[3].payload)

    def test_af_live_adapter_logs_error_and_does_not_claim_completion(self) -> None:
        class _Controller:
            def test_version(self) -> int:
                return 91

            def run_ramp(self, request):
                del request
                raise RuntimeError("ADwin process timeout")

        cfg = configured_af()
        backend = AfDemagBackendAdapter(cfg, controller=_Controller())
        command = plan_af_demag_command("AF25", cfg)

        with self.assertRaisesRegex(RuntimeError, "process timeout"):
            backend.apply_af(command)

        events = backend.communication_events()
        self.assertEqual(
            [event.direction for event in events],
            [CommunicationDirection.TX, CommunicationDirection.ERROR],
        )
        self.assertIn("AF25 failed", events[-1].detail)
        self.assertNotIn("completed", backend.status().lower())

    def test_af_live_adapter_rejects_incomplete_result_evidence(self) -> None:
        class _Controller:
            def test_version(self) -> int:
                return 91

            def run_ramp(self, request):
                del request
                return SimpleNamespace(out_count=1)

        cfg = configured_af()
        backend = AfDemagBackendAdapter(cfg, controller=_Controller())
        command = plan_af_demag_command("AF25", cfg)

        with self.assertRaises(AttributeError):
            backend.apply_af(command)

        events = backend.communication_events()
        self.assertEqual(
            [event.direction for event in events],
            [CommunicationDirection.TX, CommunicationDirection.ERROR],
        )
        self.assertNotIn("completed", backend.status().lower())

    def test_irm_live_adapter_cannot_substitute_an_af_waveform_for_a_capacitor_pulse(self) -> None:
        result = self._adwin_result()

        class _Controller:
            def __init__(self) -> None:
                self.requests = []

            def test_version(self) -> int:
                return 91

            def run_ramp(self, request):
                self.requests.append(request)
                return result

        controller = _Controller()
        backend = IrmArmBackendAdapter(IrmArmConfig(), controller=controller)

        with self.assertRaisesRegex(RuntimeError,"capacitor charge/readback/fire/discharge"):
            backend.apply_irm(max_field_mT=30.0,axis="Z (up-axis)",ramp_label="Fast (10 s)",steps=3)
        self.assertEqual(controller.requests, [])
        self.assertEqual(backend.communication_events(), ())

    def test_adwin_readiness_probe_failure_blocks_live_adapters(self) -> None:
        class _Unreachable:
            def test_version(self) -> int:
                return 0

        with self.assertRaisesRegex(HardwareUnavailableError, "unreachable"):
            AfDemagBackendAdapter(AfDemagConfig(), controller=_Unreachable())
        with self.assertRaisesRegex(HardwareUnavailableError, "unreachable"):
            IrmArmBackendAdapter(IrmArmConfig(), controller=_Unreachable())

    def test_disconnected_vacuum_snapshot_fails_without_reading_fake_pressure(self) -> None:
        class _Disconnected:
            simulated = False

            def is_connected(self) -> bool:
                return False

            def is_pump_on(self) -> bool:
                return False

            def status(self) -> str:
                return "serial unavailable"

            def read_pressure(self) -> float:
                raise AssertionError("disconnected backend must not be sampled")

        snapshot = read_vacuum_snapshot(_Disconnected(), warn_threshold=20.0)

        self.assertTrue(snapshot.fault)
        self.assertIsNone(snapshot.pressure_mtorr)
        self.assertIn("not connected", snapshot.fault_reason.lower())

    def test_command_only_vacuum_requires_explicit_threshold_opt_out(self) -> None:
        class _CommandOnly:
            simulated = False
            pressure_telemetry_available = False

            def is_connected(self) -> bool:
                return True

            def is_pump_on(self) -> bool:
                return True

            def status(self) -> str:
                return "legacy controller connected"

            def read_pressure(self) -> float:
                raise AssertionError("command-only controller has no pressure value")

        blocked = read_vacuum_snapshot(_CommandOnly(), warn_threshold=20.0)
        accepted = read_vacuum_snapshot(_CommandOnly(), warn_threshold=0.0)

        self.assertTrue(blocked.fault)
        self.assertIn("no pressure telemetry", blocked.fault_reason)
        self.assertFalse(accepted.fault)
        self.assertIsNone(accepted.pressure_mtorr)
        self.assertIn("command-only mode explicitly enabled", accepted.status)

    def test_live_vacuum_adapter_is_acknowledged_and_never_models_pressure(self) -> None:
        class _Controller:
            def __init__(self, *, trace=None) -> None:
                self._trace = trace
                self.is_connected = False
                self.is_enabled = False
                self.acknowledgements = []

            def connect(self, port: str, baudrate: int) -> None:
                self._port, self._baud = port, baudrate
                self.is_connected = bool(port and baudrate)

            def disconnect(self):
                self.is_connected = False

            def set_enabled(self, enabled: bool) -> None:
                commands = ("E", "10MFF", "O", "10VFF") if enabled else (
                    "C", "10V00", "D", "10M00"
                )
                for command in commands:
                    self._trace("TX", command, "vacuum command")
                self._trace("RX", "ACK", "vacuum response")
                self.acknowledgements.append({'commands': list(commands), 'reply': 'ACK'})
                self.is_enabled = bool(enabled)

        with mock.patch.object(diagnostic_services, "VacuumController", _Controller):
            backend = VacuumBackendAdapter(VacuumConfig(port="COM8", baud=9600))
            backend.connect()

        backend.set_pump(True)
        events = backend.communication_events()

        self.assertTrue(backend.is_connected())
        self.assertTrue(backend.is_pump_on())
        self.assertEqual(
            [(event.direction.value, event.payload) for event in events],
            [
                ("TX", "E"),
                ("TX", "10MFF"),
                ("TX", "O"),
                ("TX", "10VFF"),
                ("RX", "ACK"),
            ],
        )
        with self.assertRaisesRegex(HardwareUnavailableError, "no pressure telemetry"):
            backend.read_pressure()
        backend.set_pump(False)
        backend.disconnect()

    def test_live_vacuum_adapter_connection_failure_does_not_become_simulation(self) -> None:
        class _BrokenController:
            def __init__(self, *, trace=None) -> None:
                del trace

            def connect(self, port: str, baudrate: int) -> None:
                raise OSError(f"{port}:{baudrate} refused")

        with mock.patch.object(diagnostic_services, "VacuumController", _BrokenController):
            with self.assertRaisesRegex(HardwareUnavailableError, "COM9:9600 refused"):
                backend = build_vacuum_backend(
                    VacuumConfig(port="COM9", baud=9600),
                    nocomm=False,
                )
                backend.connect()

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
        backend.disconnect()
        self.assertFalse(backend.is_connected())

    def test_live_squid_disconnect_closes_and_forgets_reader(self) -> None:
        class _Reader:
            def __init__(self) -> None:
                self.disconnected = False

            def disconnect(self) -> None:
                self.disconnected = True

        reader = _Reader()
        backend = diagnostic_services.SquidBackendAdapter()
        backend._reader = reader
        backend._baseline_raw = (1.0, 2.0, 3.0)

        backend.disconnect()

        self.assertTrue(reader.disconnected)
        self.assertIsNone(backend._reader)
        self.assertIsNone(backend._baseline_raw)
        self.assertFalse(backend.is_connected())

    def test_live_squid_adapter_does_not_substitute_moment_for_susceptibility(self) -> None:
        backend = diagnostic_services.SquidBackendAdapter.__new__(
            diagnostic_services.SquidBackendAdapter
        )

        with self.assertRaisesRegex(
            diagnostic_services.DiagnosticContractError,
            "magnetic moment, not susceptibility",
        ):
            backend.read_susceptibility()

    def test_susceptibility_adapter_routes_zero_measure_and_events(self) -> None:
        class _Client:
            def __init__(self) -> None:
                self.is_connected = False
                self.calls: list[str] = []

            def connect(self) -> None:
                self.calls.append("connect")
                self.is_connected = True

            def zero(self) -> str:
                self.calls.append("zero")
                return "OK\r"

            def measure(self) -> float:
                self.calls.append("measure")
                return 0.0123

            def close(self) -> None:
                self.calls.append("close")
                self.is_connected = False

            def communication_events(self):
                return ()

        client = _Client()
        cfg = SusceptibilityConfig(enabled=True, port="COM7")
        backend = diagnostic_services.SusceptibilityBackendAdapter(cfg, client=client)

        self.assertFalse(backend.is_connected())
        self.assertEqual(client.calls, [])
        self.assertTrue(backend.test_connection())
        self.assertEqual(backend.zero(), "OK\r")
        self.assertAlmostEqual(backend.measure(), 0.0123)
        self.assertEqual(backend.communication_events(), ())
        backend.disconnect()
        self.assertEqual(client.calls, ["connect", "zero", "measure", "close"])
        self.assertFalse(backend.is_connected())

    def test_susceptibility_adapter_refuses_commands_until_explicit_connect(self) -> None:
        class _Client:
            is_connected = False

            def zero(self):
                raise AssertionError("zero must not be called while disconnected")

            def measure(self):
                raise AssertionError("measure must not be called while disconnected")

            def communication_events(self):
                return ()

        backend = diagnostic_services.SusceptibilityBackendAdapter(
            SusceptibilityConfig(enabled=True, port="COM7"), client=_Client()
        )

        with self.assertRaisesRegex(
            diagnostic_services.DiagnosticContractError, "disconnected"
        ):
            backend.zero()

    def test_susceptibility_factory_fails_closed_when_disabled(self) -> None:
        cfg = SusceptibilityConfig(enabled=False, port="COM7")

        with self.assertRaisesRegex(
            diagnostic_services.HardwareUnavailableError,
            "disabled",
        ):
            diagnostic_services.build_susceptibility_backend(cfg, nocomm=False)

        simulated = diagnostic_services.build_susceptibility_backend(cfg, nocomm=True)
        self.assertTrue(simulated.simulated)
        self.assertEqual(simulated.zero(), "OK")
        self.assertGreater(simulated.measure(), 0.0)

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

    def test_build_irm_arm_factory_fails_closed_in_hardware_mode(self) -> None:
        class _BrokenAdapter:
            def __init__(self, cfg: IrmArmConfig) -> None:
                del cfg
                raise RuntimeError("adapter not available")

        with mock.patch("rapid_main.diagnostic_services.IrmArmBackendAdapter", _BrokenAdapter):
            with self.assertRaises(HardwareUnavailableError) as ctx:
                build_irm_arm_backend(IrmArmConfig(), nocomm=False)
            self.assertIn("adapter not available", str(ctx.exception))

            # The simulator is still reachable, but only as an explicit opt-in.
            fallback = build_irm_arm_backend(
                IrmArmConfig(), nocomm=False, allow_simulation_fallback=True
            )
        self.assertIsInstance(fallback, IrmArmNoCommBackend)
        self.assertTrue(fallback.simulated)

    def test_missing_package_and_port_failures_block_hardware_mode(self) -> None:
        class _MissingPackage:
            def __init__(self, cfg) -> None:
                del cfg
                raise ImportError("No module named updown_control")

        class _PortUnavailable:
            def __init__(self, cfg) -> None:
                del cfg
                raise OSError("could not open port COM9")

        with mock.patch("rapid_main.diagnostic_services.SquidBackendAdapter", _MissingPackage):
            with self.assertRaisesRegex(HardwareUnavailableError, "updown_control"):
                build_squid_backend(SquidConfig(), nocomm=False)

        with mock.patch("rapid_main.diagnostic_services.VacuumBackendAdapter", _PortUnavailable):
            with self.assertRaisesRegex(HardwareUnavailableError, "COM9"):
                build_vacuum_backend(VacuumConfig(), nocomm=False)

    def test_unavailable_backend_is_not_a_simulator(self) -> None:
        class _Broken:
            def __init__(self, cfg) -> None:
                del cfg
                raise RuntimeError("driver missing")

        with mock.patch("rapid_main.diagnostic_services.AfDemagBackendAdapter", _Broken):
            backend = build_backend_or_unavailable(
                "AF demagnetizer", build_af_demag_backend, AfDemagConfig(), nocomm=False
            )

        self.assertIsInstance(backend, UnavailableBackend)
        self.assertFalse(backend.simulated)
        self.assertFalse(backend.is_connected())
        self.assertIn("driver missing", backend.status())
        with self.assertRaises(HardwareUnavailableError):
            backend.apply_af(object())

    def test_mixed_live_and_simulated_configuration_is_reported(self) -> None:
        class _WorkingAdapter:
            simulated = False

            def __init__(self, cfg) -> None:
                del cfg

            def is_connected(self) -> bool:
                return True

            def status(self) -> str:
                return "live"

        with mock.patch("rapid_main.diagnostic_services.VacuumBackendAdapter", _WorkingAdapter):
            live = build_vacuum_backend(VacuumConfig(), nocomm=False)
        simulated = build_squid_backend(SquidConfig(), nocomm=True)

        lines = collect_diagnostic_status({"Vacuum": live, "SQUID": simulated})
        by_name = {line.name: line for line in lines}

        self.assertFalse(by_name["Vacuum"].simulated)
        self.assertTrue(by_name["SQUID"].simulated)

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

    def test_dc_motor_factory_returns_no_comm_for_nocomm_and_fails_closed(self) -> None:
        self.assertIsInstance(
            build_dcmotor_backend(port="COM1", baud=9600, nocomm=True),
            DCMotorNoCommBackend,
        )

        # A constructor failure in hardware mode must not become a simulator.
        with mock.patch("rapid_main.diagnostic_services.DCMotorBackendAdapter", side_effect=RuntimeError("boom")):
            with self.assertRaisesRegex(HardwareUnavailableError, "boom"):
                build_dcmotor_backend(port="COM_DOES_NOT_EXIST", baud=9600, nocomm=False)

            backend = build_dcmotor_backend(
                port="COM_DOES_NOT_EXIST", baud=9600, nocomm=False, allow_simulation_fallback=True
            )

        self.assertIsInstance(backend, DCMotorNoCommBackend)
        self.assertTrue(backend.simulated)
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
            def __init__(self, *, trace=None) -> None:
                self.connected = False
                self.trace = trace

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

    def test_dc_motor_adapter_rejects_unsuccessful_motion_result(self) -> None:
        class _FailedMotorSerialClient:
            def __init__(self, *, trace=None) -> None:
                self.connected = False
                self.trace = trace

            def connect(self, port: str, baudrate: int = 9600) -> None:
                del port, baudrate
                self.connected = True

            @property
            def is_connected(self) -> bool:
                return self.connected

            def move_motor(self, *args, **kwargs) -> object:
                del args, kwargs
                return mock.Mock(target=100, final_position=75, success=False)

            def stop(self, axis) -> None:
                del axis

            def halt(self, axis) -> None:
                del axis

            def read_registers(self, axis, registers):
                del axis, registers
                return (75, 0)

        with mock.patch(
            "rapid_main.diagnostic_services.MotorSerialClient",
            _FailedMotorSerialClient,
        ):
            backend = build_dcmotor_backend(port="COM3", baud=9600, nocomm=False)
            backend.connect("COM3", 9600)
            with self.assertRaisesRegex(HardwareError, "target=100 final=75"):
                backend.move_motor("Changer (X)", target=100, speed=1200)

        events = backend.communication_events()
        self.assertEqual(events[-1].direction, CommunicationDirection.ERROR)
        self.assertIn("move Changer (X) failed", events[-1].detail)
