"""Continuous SQUID streaming: read-rate control, pacing, failure isolation."""
from __future__ import annotations

import math
import threading
import time
import unittest

from rapidpy_common.squid_stream import (
    BackgroundSquidCapture,
    PositionSample,
    SquidReadTiming,
    SquidStreamError,
    SquidStreamSampler,
    SquidTrace,
    effective_period_s,
    estimate_sample_period_s,
)


class _Reading:
    def __init__(self, counts: float, dvm: float) -> None:
        self.counts = counts
        self.dvm = dvm


class FakeStreamClient:
    """Minimal 2G client: counts and DVM follow a programmable function of time."""

    def __init__(self, *, fail_after: int | None = None, block_event: threading.Event | None = None) -> None:
        self.calls: list[tuple] = []
        self._latches = 0
        self._fail_after = fail_after
        self._block_event = block_event

    def latch(self, axis="A", *, settle_s=0.0, count_hold_s=None, data_hold_s=None):
        if self._block_event is not None:
            self._block_event.wait(5.0)
        self._latches += 1
        if self._fail_after is not None and self._latches > self._fail_after:
            raise OSError("serial link dropped")
        self.calls.append(("latch", count_hold_s, data_hold_s))
        return (f"{axis}LC", f"{axis}LD")

    def read_axis(self, axis, *, range_value=1.0):
        self.calls.append(("read_axis", axis))
        return _Reading(counts=2.0, dvm=0.25 * "XYZ".index(axis) + self._latches)

    def read_axis_data(self, axis):
        self.calls.append(("read_axis_data", axis))
        return 0.25 * "XYZ".index(axis) + self._latches


class FakeTime:
    def __init__(self) -> None:
        self.now = 100.0
        self.slept: list[float] = []

    def monotonic(self) -> float:
        self.now += 0.001
        return self.now

    def sleep(self, seconds: float) -> None:
        self.slept.append(seconds)
        self.now += seconds


class ReadTimingTests(unittest.TestCase):
    def test_defaults_match_vb6_latch_holds(self):
        timing = SquidReadTiming()
        self.assertEqual(timing.latch_count_hold_s, 0.10)
        self.assertEqual(timing.latch_data_hold_s, 0.12)
        self.assertEqual(timing.axes, ("X", "Y", "Z"))

    def test_axes_are_normalised_and_validated(self):
        self.assertEqual(SquidReadTiming(axes=("z", "x")).axes, ("X", "Z"))
        for bad in ((), ("Q",), ("X", "X")):
            with self.assertRaises(SquidStreamError):
                SquidReadTiming(axes=bad)

    def test_rejects_invalid_rates_and_counter_refresh(self):
        with self.assertRaises(SquidStreamError):
            SquidReadTiming(interval_s=-1)
        with self.assertRaises(SquidStreamError):
            SquidReadTiming(counts_every=0)
        with self.assertRaises(SquidStreamError):
            SquidReadTiming(reply_timeout_s=0.0)

    def test_link_limited_period_estimate(self):
        full = SquidReadTiming(interval_s=0.0)
        fast = SquidReadTiming(interval_s=0.0, axes=("Z",), counts_every=10, latch_count_hold_s=0.0, latch_data_hold_s=0.0)
        slow_link = estimate_sample_period_s(full, 1200)
        self.assertGreater(slow_link, 0.22 + 6 * 17 * 10 / 1200)  # holds + six query frames
        self.assertLess(estimate_sample_period_s(full, 9600), slow_link)
        self.assertLess(estimate_sample_period_s(fast, 9600), estimate_sample_period_s(full, 9600) / 5)
        # Requested rate faster than the link: the link wins.
        self.assertEqual(effective_period_s(SquidReadTiming(interval_s=0.01), 1200), estimate_sample_period_s(SquidReadTiming(interval_s=0.01), 1200))
        self.assertEqual(effective_period_s(SquidReadTiming(interval_s=5.0), 9600), 5.0)

    def test_round_trip_dict(self):
        timing = SquidReadTiming(interval_s=0.25, axes=("X", "Y"), counts_every=3)
        self.assertEqual(SquidReadTiming.from_dict(timing.to_dict()), timing)


class SamplerTests(unittest.TestCase):
    def test_paces_samples_at_requested_interval(self):
        clock = FakeTime()
        client = FakeStreamClient()
        sampler = SquidStreamSampler(client, SquidReadTiming(interval_s=0.5), monotonic=clock.monotonic, sleep=clock.sleep)
        trace = sampler.run(SquidTrace("pace", sampler.timing), max_samples=5)
        self.assertEqual(len(trace.samples), 5)
        gaps = [b.t_s - a.t_s for a, b in zip(trace.samples, trace.samples[1:])]
        for gap in gaps:
            self.assertAlmostEqual(gap, 0.5, delta=0.06)
        self.assertAlmostEqual(trace.achieved_rate_hz, 2.0, delta=0.25)
        self.assertEqual(trace.stop_reason, "sample limit")

    def test_custom_latch_holds_are_passed_per_call(self):
        client = FakeStreamClient()
        timing = SquidReadTiming(interval_s=0.0, latch_count_hold_s=0.02, latch_data_hold_s=0.03)
        SquidStreamSampler(client, timing).run(SquidTrace("holds", timing), max_samples=2)
        latches = [call for call in client.calls if call[0] == "latch"]
        self.assertEqual(latches, [("latch", 0.02, 0.03)] * 2)

    def test_counts_every_reuses_counter_and_flags_samples(self):
        client = FakeStreamClient()
        timing = SquidReadTiming(interval_s=0.0, axes=("Z",), counts_every=3)
        trace = SquidStreamSampler(client, timing).run(SquidTrace("counts", timing), max_samples=6)
        fresh = [sample.counts_fresh for sample in trace.samples]
        self.assertEqual(fresh, [True, False, False, True, False, False])
        data_only = [call for call in client.calls if call[0] == "read_axis_data"]
        self.assertEqual(len(data_only), 4)
        # VB6 getVal: -data - count * range, with the reused counter value.
        second = trace.samples[1]
        self.assertAlmostEqual(second.raw[2], -second.dvm[2] - 2.0)
        self.assertTrue(math.isnan(second.raw[0]))

    def test_transport_failure_ends_stream_without_raising(self):
        client = FakeStreamClient(fail_after=3)
        timing = SquidReadTiming(interval_s=0.0)
        trace = SquidStreamSampler(client, timing).run(SquidTrace("fail", timing), max_samples=10)
        self.assertEqual(len(trace.samples), 3)
        self.assertEqual(trace.stop_reason, "transport error")
        self.assertIn("serial link dropped", trace.errors[0])

    def test_csv_round_trip_preserves_samples_and_positions(self):
        timing = SquidReadTiming(interval_s=0.0, axes=("X", "Z"))
        trace = SquidStreamSampler(FakeStreamClient(), timing).run(SquidTrace("csv", timing, motion={"segment": "turn"}), max_samples=4)
        trace.positions.append(PositionSample(0.1, 45.0))
        restored = SquidTrace.from_csv(trace.to_csv())
        self.assertEqual(restored.label, "csv")
        self.assertEqual(restored.timing, timing)
        self.assertEqual(len(restored.samples), 4)
        self.assertEqual(restored.motion["segment"], "turn")
        self.assertAlmostEqual(restored.samples[2].raw[2], trace.samples[2].raw[2])
        self.assertTrue(math.isnan(restored.samples[0].raw[1]))
        self.assertEqual(restored.positions[0].position, 45.0)


class BackgroundCaptureTests(unittest.TestCase):
    def test_stop_joins_worker_before_returning(self):
        timing = SquidReadTiming(interval_s=0.01, latch_count_hold_s=0.0, latch_data_hold_s=0.0)
        capture = BackgroundSquidCapture(SquidStreamSampler(FakeStreamClient(), timing), "bg")
        with capture:
            time.sleep(0.08)
            capture.mark_position(10.0)
        self.assertFalse(capture.running)
        self.assertGreater(len(capture.trace.samples), 1)
        self.assertEqual(capture.trace.stop_reason, "motion complete")
        self.assertEqual(len(capture.trace.positions), 1)

    def test_unreleased_link_raises_instead_of_returning(self):
        gate = threading.Event()
        timing = SquidReadTiming(interval_s=0.0)
        capture = BackgroundSquidCapture(SquidStreamSampler(FakeStreamClient(block_event=gate), timing), "stuck", join_timeout_s=0.05)
        capture.start()
        time.sleep(0.02)
        with self.assertRaises(SquidStreamError):
            capture.stop()
        gate.set()  # let the worker finish so the test process exits cleanly
        capture._thread.join(2.0)


class RawSquidClientTimingTests(unittest.TestCase):
    def test_latch_uses_vb6_holds_by_default_and_overrides_per_call(self):
        from updown_control.app import RawSquidClient

        client = RawSquidClient()
        sent: list[str] = []
        slept: list[float] = []
        client._send = sent.append  # type: ignore[method-assign]
        import updown_control.app as module

        original_sleep = module.time.sleep
        module.time.sleep = slept.append
        try:
            client.latch("A")
            client.latch("A", count_hold_s=0.01, data_hold_s=0.02)
        finally:
            module.time.sleep = original_sleep
        self.assertEqual(sent, ["ALC", "ALD", "ALC", "ALD"])
        self.assertEqual(slept, [0.10, 0.12, 0.01, 0.02])

    def test_rejects_unbounded_holds(self):
        from updown_control.app import RawSquidClient, SquidCommunicationError

        client = RawSquidClient()
        client._send = lambda command: None  # type: ignore[method-assign]
        with self.assertRaises(SquidCommunicationError):
            client.latch("A", count_hold_s=-1.0)


if __name__ == "__main__":
    unittest.main()
