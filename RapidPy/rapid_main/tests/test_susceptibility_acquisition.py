from __future__ import annotations

from datetime import datetime, timedelta, timezone
import json
from pathlib import Path
import tempfile
import unittest

from rapid_main.acquisition import MotionOutcome
from rapid_main.communication_log import CommunicationDirection, CommunicationEvent
from rapid_main.susceptibility_acquisition import (
    SUSCEPTIBILITY_ACQUISITION_SCHEMA,
    SusceptibilityAcquisitionConfig,
    SusceptibilityAcquisitionError,
    SusceptibilityAcquisitionService,
    write_susceptibility_acquisition,
)


class _Clock:
    def __init__(self) -> None:
        self.now = datetime(2026, 10, 2, 12, 0, tzinfo=timezone.utc)

    def __call__(self) -> datetime:
        value = self.now
        self.now += timedelta(milliseconds=1)
        return value


class _Bridge:
    simulated = False

    def __init__(self, value: float = 12.5, *, fail_zero: bool = False, fail_measure: bool = False) -> None:
        self.value = value
        self.fail_zero = fail_zero
        self.fail_measure = fail_measure
        self.calls: list[str] = []
        self.events: list[CommunicationEvent] = []

    def is_connected(self) -> bool:
        return True

    def zero(self) -> str:
        self.calls.append("zero")
        if self.fail_zero:
            raise TimeoutError("zero reply timeout")
        self._event(CommunicationDirection.TX, "Z\\r\\n")
        self._event(CommunicationDirection.RX, "OK\\r")
        return "OK\r"

    def measure(self) -> float:
        self.calls.append("measure")
        if self.fail_measure:
            raise TimeoutError("measure reply timeout")
        self._event(CommunicationDirection.TX, "M\\r\\n")
        self._event(CommunicationDirection.RX, f"{self.value}\\r")
        return self.value

    def communication_events(self):
        return tuple(self.events)

    def _event(self, direction, payload):
        self.events.append(
            CommunicationEvent(
                datetime(2026, 10, 2, 12, 0, tzinfo=timezone.utc),
                "SUSCEPTIBILITY",
                direction,
                payload,
                "COM7",
            )
        )


class _Motion:
    def __init__(self, *, fail_move: bool = False, fail_home_call: int = 0) -> None:
        self.current = 500
        self.calls: list[object] = []
        self.fail_move = fail_move
        self.fail_home_call = fail_home_call
        self.home_calls = 0

    def position(self) -> int:
        self.calls.append("position")
        return self.current

    def home_to_top(self) -> MotionOutcome:
        self.home_calls += 1
        self.calls.append("home")
        if self.home_calls == self.fail_home_call:
            return MotionOutcome(0, self.current, False, "home limit not confirmed")
        self.current = 0
        return MotionOutcome(0, 0, True)

    def move_to(self, position: int, *, speed_index: int = 0) -> MotionOutcome:
        self.calls.append(("move", position, speed_index))
        if self.fail_move:
            return MotionOutcome(position, 3, False, "position mismatch")
        self.current = position
        return MotionOutcome(position, position, True)


def _config(**overrides) -> SusceptibilityAcquisitionConfig:
    values = dict(coil_position=-1000, sample_height=100, moment_factor_cgs=2.0e-5)
    values.update(overrides)
    return SusceptibilityAcquisitionConfig(**values)


class SusceptibilityAcquisitionTests(unittest.TestCase):
    def _service(self, bridge=None, motion=None, **kwargs):
        return SusceptibilityAcquisitionService(
            bridge or _Bridge(),
            motion or _Motion(),
            config=kwargs.pop("config", _config()),
            clock=_Clock(),
            id_factory=lambda: "susc-test-1",
            **kwargs,
        )

    def test_sample_follows_exact_order_and_applies_holder_math(self) -> None:
        bridge = _Bridge(12.5)
        motion = _Motion()

        record = self._service(bridge, motion).acquire(
            sample_id="SAMPLE-1",
            is_holder=False,
            holder_scaled_value=2.5,
            holder_evidence_id="holder-susc-1",
        )

        self.assertEqual(bridge.calls, ["zero", "measure"])
        self.assertEqual(
            motion.calls,
            ["position", "home", ("move", -950, 0), "home"],
        )
        self.assertAlmostEqual(record.susceptibility, (12.5 - 2.5) * 2.0e-5)
        self.assertTrue(record.safe_state_confirmed)
        self.assertEqual(record.outcome, "completed")
        self.assertEqual(record.holder_evidence_id, "holder-susc-1")
        self.assertEqual(len(record.communication_events), 4)

    def test_holder_keeps_scaled_bridge_value_without_subtraction(self) -> None:
        record = self._service().acquire(sample_id="Holder", is_holder=True)

        self.assertEqual(record.bridge_scaled_value, 12.5)
        self.assertEqual(record.susceptibility, 12.5)
        self.assertIsNone(record.holder_scaled_value)

    def test_configuration_and_holder_are_validated_before_motion(self) -> None:
        motion = _Motion()
        with self.assertRaises(SusceptibilityAcquisitionError) as caught:
            self._service(motion=motion, config=_config(coil_position=0)).acquire(
                sample_id="S1", is_holder=False, holder_scaled_value=1.0, holder_evidence_id="h"
            )
        self.assertEqual(motion.calls, [])
        self.assertIn("not configured", caught.exception.record.error)

        motion = _Motion()
        with self.assertRaises(SusceptibilityAcquisitionError) as caught:
            self._service(motion=motion).acquire(sample_id="S1", is_holder=False)
        self.assertEqual(motion.calls, [])
        self.assertIn("finite holder", caught.exception.record.error)

    def test_measure_failure_returns_home_and_publishes_no_value(self) -> None:
        bridge = _Bridge(fail_measure=True)
        motion = _Motion()

        with self.assertRaises(SusceptibilityAcquisitionError) as caught:
            self._service(bridge, motion).acquire(sample_id="Holder", is_holder=True)

        record = caught.exception.record
        self.assertEqual(motion.calls[-1], "home")
        self.assertTrue(record.safe_state_confirmed)
        self.assertIsNone(record.susceptibility)
        self.assertIn("measure reply timeout", record.error)

    def test_safe_return_failure_is_distinct_and_preserves_primary_error(self) -> None:
        bridge = _Bridge(fail_zero=True)
        motion = _Motion(fail_home_call=2)

        with self.assertRaises(SusceptibilityAcquisitionError) as caught:
            self._service(bridge, motion).acquire(sample_id="Holder", is_holder=True)

        record = caught.exception.record
        self.assertIn("zero reply timeout", record.error)
        self.assertIn("safe return failed", record.error)
        self.assertIn("home limit not confirmed", record.safe_return_error)
        self.assertFalse(record.safe_state_confirmed)

    def test_motion_mismatch_never_calls_measure_and_returns_home(self) -> None:
        bridge = _Bridge()
        motion = _Motion(fail_move=True)

        with self.assertRaises(SusceptibilityAcquisitionError):
            self._service(bridge, motion).acquire(sample_id="Holder", is_holder=True)

        self.assertEqual(bridge.calls, ["zero"])
        self.assertEqual(motion.calls[-1], "home")

    def test_cancellation_is_fail_closed(self) -> None:
        checks = iter((False, True))
        motion = _Motion()
        with self.assertRaises(SusceptibilityAcquisitionError) as caught:
            self._service(motion=motion, should_cancel=lambda: next(checks)).acquire(
                sample_id="Holder", is_holder=True
            )
        self.assertIn("cancelled", caught.exception.record.error)
        self.assertEqual(motion.calls, ["position", "home", "home"])

    def test_artifact_is_atomic_immutable_and_versioned(self) -> None:
        record = self._service().acquire(sample_id="Holder", is_holder=True)
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "holder.susc.json"
            write_susceptibility_acquisition(path, record)
            payload = json.loads(path.read_text(encoding="utf-8"))
            self.assertEqual(payload["schema"], SUSCEPTIBILITY_ACQUISITION_SCHEMA)
            self.assertEqual(payload["acquisition_id"], "susc-test-1")
            self.assertEqual(list(path.parent.glob("*.tmp-*")), [])
            with self.assertRaises(FileExistsError):
                write_susceptibility_acquisition(path, record)


if __name__ == "__main__":
    unittest.main(verbosity=2)
