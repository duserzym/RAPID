"""Queue-level behavior of the fail-closed hardware backend.

These tests exercise ``QueueHardwareBackend`` as a whole: construction
failures, holder measurement as a real command, holder-gated sample reads, and
the recovery-hook contract the worker depends on.
"""
from __future__ import annotations

from datetime import datetime, timedelta, timezone
from pathlib import Path
import tempfile
import unittest
from unittest import mock

from rapid_main.config import AppConfig
from rapid_main.communication_log import (
    CommunicationDirection,
    CommunicationEvent,
    CommunicationLogger,
)
from rapid_main.diagnostic_services import HardwareUnavailableError
from rapid_main.hardware_contracts import (
    QueueAutomationError,
    QueueHardwareBackend,
    config_fingerprint,
)
from rapid_main.holder_state import HolderStateError, HolderStateStore
from rapid_main.magnetometer import reduce_bracketed_measurement

from tests.acquisition_fakes import (
    FakeClock,
    FakeMotorSerialClient,
    FakeRawSquidClient,
    counts_for,
    counts_with_step,
)

CALIBRATION = (0.09000563, 0.106, 0.066)
ZERO_POS = -25886
MEAS_POS = -30607
NOW = datetime(2026, 8, 29, 12, 0, 0, tzinfo=timezone.utc)


def _observations(*, zero_after_step_axis: int | None = None):
    zero_after = (
        counts_with_step((0.0, 0.0, 0.0), CALIBRATION, axis=zero_after_step_axis)
        if zero_after_step_axis is not None
        else counts_for((0.0, 0.0, 0.0), CALIBRATION)
    )
    return [
        counts_for((0.0, 0.0, 0.0), CALIBRATION),
        counts_for((0.20, 0.05, 0.30), CALIBRATION),
        counts_for((0.06, 0.19, 0.31), CALIBRATION),
        counts_for((-0.21, 0.04, 0.29), CALIBRATION),
        counts_for((-0.05, -0.20, 0.30), CALIBRATION),
        zero_after,
    ]


class _StubSquidAdapter:
    """SQUID backend that exposes the atomic 2G client, like the live adapter."""

    simulated = False

    def __init__(self, raw_client) -> None:
        self.raw_client = raw_client
        self.susceptibility_calls = 0

    def is_connected(self) -> bool:
        return True

    def test_connection(self) -> bool:
        return True

    def status(self) -> str:
        return "stub SQUID"

    def read_susceptibility(self) -> float:
        self.susceptibility_calls += 1
        return 0.001


class _PlainSquidAdapter:
    """SQUID backend with no raw client, e.g. a diagnostic-only transport."""

    simulated = False

    def __init__(self) -> None:
        self.reads = 0

    def is_connected(self) -> bool:
        return True

    def test_connection(self) -> bool:
        return True

    def read_squid(self) -> tuple[float, float, float]:
        self.reads += 1
        return (1.0, 2.0, 3.0)

    def read_susceptibility(self) -> float:
        return 0.0


class _CommunicationSource:
    def __init__(self, event: CommunicationEvent, *, simulated: bool = False) -> None:
        self._event = event
        self.simulated = simulated

    def communication_events(self) -> tuple[CommunicationEvent, ...]:
        return (self._event,)


def _config(tmp: Path, *, configured_motion: bool = True) -> AppConfig:
    cfg = AppConfig()
    cfg.general.nocomm = False
    cfg.general.data_dir = str(tmp)
    cfg.general.operator = "opr"
    cfg.changer.port = "COM3"
    cfg.squid.samples_per_pos = 1
    cfg.calibration.cal_x, cfg.calibration.cal_y, cfg.calibration.cal_z = CALIBRATION
    if configured_motion:
        cfg.motion.zero_pos = ZERO_POS
        cfg.motion.meas_pos = MEAS_POS
    return cfg


class QueueBackendFailClosedTests(unittest.TestCase):
    def setUp(self) -> None:
        self._tmp = tempfile.TemporaryDirectory()
        self.tmp = Path(self._tmp.name)

    def tearDown(self) -> None:
        self._tmp.cleanup()

    def test_communication_events_merge_live_treatment_sources_in_time_order(self) -> None:
        motor = CommunicationEvent(
            timestamp=NOW - timedelta(seconds=1),
            channel="DC_MOTOR",
            direction=CommunicationDirection.TX,
            payload="motor-request",
        )
        early = CommunicationEvent(
            timestamp=NOW,
            channel="ADWIN_AF",
            direction=CommunicationDirection.TX,
            payload="af-request",
        )
        middle = CommunicationEvent(
            timestamp=NOW + timedelta(seconds=1),
            channel="ADWIN_IRM_ARM",
            direction=CommunicationDirection.RX,
            payload="irm-result",
        )
        simulated = CommunicationEvent(
            timestamp=NOW + timedelta(seconds=2),
            channel="SIMULATED",
            direction=CommunicationDirection.INFO,
            payload="must-not-publish",
        )
        backend = object.__new__(QueueHardwareBackend)
        backend._motor_communication_logger = CommunicationLogger("DC_MOTOR")
        backend._motor_communication_logger.transcript.append(motor)
        backend._bracketed = None
        backend._af_demag = _CommunicationSource(early)
        backend._irm_arm = _CommunicationSource(middle)

        self.assertEqual(backend.communication_events(), (motor, early, middle))

        backend._af_demag = _CommunicationSource(simulated, simulated=True)
        self.assertEqual(backend.communication_events(), (motor, middle))

    def test_component_construction_failure_becomes_a_preflight_blocker(self) -> None:
        cfg = _config(self.tmp)
        with mock.patch("rapid_main.hardware_contracts.MotorSerialClient", FakeMotorSerialClient):
            with mock.patch(
                "rapid_main.diagnostic_services.build_squid_backend",
                side_effect=HardwareUnavailableError("SQUID backend is unavailable: no serial port"),
            ):
                backend = QueueHardwareBackend(cfg)

        result = backend.preflight()

        self.assertFalse(result.ok)
        self.assertTrue(any("no serial port" in blocker for blocker in result.blockers), result.blockers)
        self.assertFalse(backend.is_available())

    def test_missing_measurement_backend_never_returns_a_reading(self) -> None:
        cfg = _config(self.tmp)
        with mock.patch("rapid_main.hardware_contracts.MotorSerialClient", FakeMotorSerialClient):
            with mock.patch(
                "rapid_main.diagnostic_services.build_squid_backend",
                side_effect=HardwareUnavailableError("driver missing"),
            ):
                backend = QueueHardwareBackend(cfg)

        with self.assertRaisesRegex(Exception, "driver missing"):
            backend.read_squid()

    def test_unconfigured_lift_positions_block_preflight(self) -> None:
        cfg = _config(self.tmp, configured_motion=False)
        backend = self._backend(cfg)

        result = backend.preflight()

        self.assertFalse(result.ok)
        self.assertTrue(
            any("lift positions are not configured" in blocker for blocker in result.blockers),
            result.blockers,
        )

    def test_failed_identity_check_blocks_preflight(self) -> None:
        class _IdentityFails:
            simulated = False

            def is_connected(self) -> bool:
                return False

            def test_connection(self) -> bool:
                raise RuntimeError("SQUID connected but identity query returned garbage")

            def status(self) -> str:
                return "identity check failed"

        cfg = _config(self.tmp)
        backend = self._backend(cfg, adapter=_IdentityFails())

        result = backend.preflight()

        self.assertFalse(result.ok)
        self.assertTrue(
            any("identity query returned garbage" in blocker for blocker in result.blockers),
            result.blockers,
        )
        self.assertFalse(backend.is_available())

    def test_recovery_hook_is_absent_without_a_bracketed_backend(self) -> None:
        cfg = _config(self.tmp)
        backend = self._backend(cfg, adapter=_PlainSquidAdapter())
        backend.preflight()

        self.assertIsNone(getattr(backend, "recover_flux_count_discontinuity", None))
        with self.assertRaises(AttributeError):
            backend.recover_flux_count_discontinuity  # noqa: B018 - attribute probe

    def test_recovery_hook_appears_once_bracketed_acquisition_is_live(self) -> None:
        cfg = _config(self.tmp)
        backend = self._backend(cfg)
        backend.preflight()

        self.assertTrue(callable(getattr(backend, "recover_flux_count_discontinuity", None)))
        self.assertGreaterEqual(int(backend.flux_discontinuity_retries), 0)

    # -- helpers -----------------------------------------------------------

    def _backend(
        self,
        cfg: AppConfig,
        *,
        adapter=None,
        observations=None,
        holder_store: HolderStateStore | None = None,
    ) -> QueueHardwareBackend:
        client = FakeRawSquidClient(observations or _observations())
        with mock.patch("rapid_main.hardware_contracts.MotorSerialClient", FakeMotorSerialClient):
            with mock.patch(
                "rapid_main.diagnostic_services.build_squid_backend",
                return_value=adapter if adapter is not None else _StubSquidAdapter(client),
            ):
                with mock.patch("rapid_main.diagnostic_services.build_irm_arm_backend", return_value=object()):
                    with mock.patch(
                        "rapid_main.diagnostic_services.build_af_demag_backend", return_value=object()
                    ):
                        backend = QueueHardwareBackend(
                            cfg, holder_store=holder_store, clock=FakeClock()
                        )
        backend._client.is_connected = True
        backend._connected = True
        return backend


class QueueBackendHolderCommandTests(unittest.TestCase):
    def setUp(self) -> None:
        self._tmp = tempfile.TemporaryDirectory()
        self.tmp = Path(self._tmp.name)
        self.holder_path = self.tmp / "holder_correction.json"
        self.store = HolderStateStore(self.holder_path, clock=lambda: NOW)

    def tearDown(self) -> None:
        self._tmp.cleanup()

    def _backend(self, *, observations=None) -> QueueHardwareBackend:
        cfg = _config(self.tmp)
        client = FakeRawSquidClient(observations or _observations())
        with mock.patch("rapid_main.hardware_contracts.MotorSerialClient", FakeMotorSerialClient):
            with mock.patch(
                "rapid_main.diagnostic_services.build_squid_backend",
                return_value=_StubSquidAdapter(client),
            ):
                with mock.patch("rapid_main.diagnostic_services.build_irm_arm_backend", return_value=object()):
                    with mock.patch(
                        "rapid_main.diagnostic_services.build_af_demag_backend", return_value=object()
                    ):
                        backend = QueueHardwareBackend(
                            cfg, holder_store=self.store, clock=FakeClock()
                        )
        backend._client.is_connected = True
        backend._connected = True
        backend.preflight()
        return backend

    def test_holder_command_measures_and_installs_a_correction(self) -> None:
        backend = self._backend()

        backend.holder(0)

        self.assertTrue(self.holder_path.exists())
        correction = self.store.current
        self.assertIsNotNone(correction)
        self.assertEqual(correction.holder_id, "holder")
        self.assertEqual(correction.operator, "opr")
        self.assertEqual(correction.config_hash, config_fingerprint(backend._config))
        self.assertGreater(correction.metrics.magnitude_raw, 0.0)
        self.assertTrue(backend.last_holder_outcome.installed)
        status = backend.holder_status()
        self.assertTrue(status.present)
        self.assertTrue(status.valid)

    def test_rejected_holder_block_aborts_the_queue_and_keeps_state(self) -> None:
        backend = self._backend()
        backend.holder(0)
        first = self.store.current
        before = self.holder_path.read_text(encoding="utf-8")

        # Replace the transport with one that steps a flux count on Y.
        backend._bracketed = None
        backend._measurement = _StubSquidAdapter(
            FakeRawSquidClient(_observations(zero_after_step_axis=1) * 8)
        )
        backend._ensure_bracketed()

        with self.assertRaises(QueueAutomationError) as ctx:
            backend.holder(0)

        self.assertIn("rejected", str(ctx.exception))
        self.assertEqual(self.store.current.record_version, first.record_version)
        self.assertEqual(self.holder_path.read_text(encoding="utf-8"), before)
        self.assertFalse(backend.last_holder_outcome.installed)

    def test_sample_read_is_blocked_until_a_holder_exists(self) -> None:
        backend = self._backend()

        with self.assertRaisesRegex(HolderStateError, "No holder correction"):
            backend.read_squid()

    def test_sample_read_is_blocked_by_a_stale_holder(self) -> None:
        backend = self._backend()
        backend.holder(0)
        backend._holder_store._clock = lambda: NOW + timedelta(days=2)

        with self.assertRaisesRegex(HolderStateError, "stale"):
            backend.read_squid()

    def test_sample_block_subtracts_the_installed_holder(self) -> None:
        backend = self._backend()
        backend.holder(0)
        holder_positions = self.store.current.positions

        backend._measurement = _StubSquidAdapter(FakeRawSquidClient(_observations()))
        backend._bracketed = None
        backend._ensure_bracketed()
        backend.set_measurement_context(sample_name="RS01A", treatment_label="AF20", run_id="run-3")

        block = backend.read_squid()

        self.assertEqual(block.holder_positions, holder_positions)
        self.assertEqual(block.audit.sample_name, "RS01A")
        self.assertEqual(block.audit.treatment_label, "AF20")
        self.assertEqual(block.audit.run_id, "run-3")
        self.assertEqual(block.audit.holder_record_id, self.store.current.record_version)
        self.assertFalse(block.audit.is_holder_block)

        # An identical holder block reduces to a zero sample moment.
        result = reduce_bracketed_measurement(block)
        for axis in result.mean_raw:
            self.assertAlmostEqual(axis, 0.0, places=9)

    def test_holder_block_itself_subtracts_nothing(self) -> None:
        backend = self._backend()
        backend.holder(0)
        blocks = backend.last_holder_outcome.blocks

        self.assertEqual(len(blocks), 1)
        self.assertTrue(blocks[0].audit.is_holder_block)
        self.assertEqual(
            blocks[0].holder_positions,
            ((0.0, 0.0, 0.0), (0.0, 0.0, 0.0), (0.0, 0.0, 0.0), (0.0, 0.0, 0.0)),
        )

    def test_flip_switches_the_recorded_sample_direction(self) -> None:
        backend = self._backend()

        self.assertTrue(backend._direction_up)
        backend.flip()
        self.assertFalse(backend._direction_up)
        backend.init_up("file-1")
        self.assertTrue(backend._direction_up)

    def test_failed_flip_does_not_mutate_direction_or_flip_state(self) -> None:
        backend = self._backend()
        backend._client.turn_failure_angle = 180.0

        with self.assertRaisesRegex(QueueAutomationError, "sample flip"):
            backend.flip()

        self.assertTrue(backend._direction_up)
        self.assertFalse(backend._last_flip)
        self.assertEqual(
            backend.communication_events()[-1].direction,
            CommunicationDirection.ERROR,
        )

    def test_failed_changer_move_does_not_update_last_hole(self) -> None:
        backend = self._backend()
        backend._client.changer_failure_hole = 12.0
        previous_hole = backend._last_hole

        with self.assertRaisesRegex(QueueAutomationError, "changer move to hole 12"):
            backend.goto_hole(12)

        self.assertEqual(backend._last_hole, previous_hole)

    def test_failed_home_does_not_claim_up_direction(self) -> None:
        backend = self._backend()
        backend._direction_up = False
        backend._client.home_failure = True

        with self.assertRaisesRegex(QueueAutomationError, "home Up/Down"):
            backend.init_up("file-1")

        self.assertFalse(backend._direction_up)

    def test_safe_return_attempts_every_halt_after_other_failures(self) -> None:
        backend = self._backend()

        class _BrokenReset:
            simulated = False

            def reset_field(self) -> None:
                raise RuntimeError("relay reset unavailable")

        backend._af_demag = _BrokenReset()
        backend._sample_loaded = True
        backend._client.dropoff_failure = True
        backend._client.halt_fail_axes.add("Turning")

        with self.assertRaisesRegex(QueueAutomationError, "Safe-state return incomplete") as ctx:
            backend.return_to_safe_state()

        self.assertIn("AF reset failed", str(ctx.exception))
        self.assertIn("safe-state sample dropoff", str(ctx.exception))
        self.assertIn("halt turning failed", str(ctx.exception))
        halt_axes = [payload for name, payload in backend._client.calls if name == "halt"]
        self.assertEqual(halt_axes, ["ChangerX", "ChangerY", "Turning", "UpDown"])
        self.assertTrue(backend._sample_loaded)


if __name__ == "__main__":
    unittest.main(verbosity=2)
