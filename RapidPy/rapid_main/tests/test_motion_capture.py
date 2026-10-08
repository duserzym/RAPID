"""Opt-in SQUID capture during descent, 90-degree turns and ascent of a block."""
from __future__ import annotations

import json
import math
import tempfile
import threading
import time
import unittest
from contextlib import contextmanager
from pathlib import Path

from rapid_main.acquisition import (
    AcquisitionConfig,
    BracketedAcquisitionService,
    MotionOutcome,
    MotionVerificationError,
)
from rapid_main.config import AppConfig
from rapid_main.magnetometer import reduce_bracketed_measurement
from rapid_main.motion_capture import (
    MotionCaptureConfig,
    MotionCaptureError,
    MotionCaptureRecorder,
    motion_capture_config_from_app_config,
    stream_timing_from_squid_config,
)
from rapid_main.squid_transport import MotorTurningController, MotorVerticalController
from rapidpy_common.squid_stream import SquidReadTiming

from tests.acquisition_fakes import FakeClock, FakeSquidTransport, counts_for

CALIBRATION = (0.09, 0.106, 0.066)
ZERO, MEAS = -25886, -30607
FAST = SquidReadTiming(interval_s=0.0, latch_count_hold_s=0.0, latch_data_hold_s=0.0)


class Station:
    """Shared physical state: where the specimen is and how it is turned."""

    def __init__(self) -> None:
        self.position = 0.0
        self.angle = 0.0
        self.lock = threading.Lock()


class _Reading:
    def __init__(self, counts, dvm):
        self.counts, self.dvm = counts, dvm


class StreamClient:
    """2G stand-in whose signal depends on the station's height and angle."""

    def __init__(self, station: Station, *, fail=False, gate: threading.Event | None = None) -> None:
        self.station = station
        self.fail = fail
        self.gate = gate
        self.latches = 0

    def latch(self, axis="A", *, settle_s=0.0, count_hold_s=None, data_hold_s=None):
        if self.gate is not None:
            self.gate.wait(5.0)
        if self.fail:
            raise OSError("2G reply timed out")
        self.latches += 1
        time.sleep(0.002)
        return ("ALC", "ALD")

    def read_axis(self, axis, *, range_value=1.0):
        with self.station.lock:
            z, theta = self.station.position, math.radians(self.station.angle)
        coupling = math.exp(-0.5 * ((z - MEAS) / 900.0) ** 2)
        value = {"X": 2.0 * math.cos(theta + 0.5), "Y": 2.0 * math.sin(theta + 0.5), "Z": 1.0}[axis] * coupling
        return _Reading(0.0, -value)


class SlowVertical:
    """Lift that moves in steps, reporting encoder polls like wait_for_motor_stop."""

    def __init__(self, station: Station) -> None:
        self.station = station
        self._observer = None
        self.moves: list[int] = []

    @contextmanager
    def observe_positions(self, callback):
        self._observer = callback
        try:
            yield
        finally:
            self._observer = None

    def move_to(self, position, *, speed_index=0):
        self.moves.append(int(position))
        start = self.station.position
        for step in range(1, 13):
            with self.station.lock:
                self.station.position = start + (position - start) * step / 12
            if self._observer is not None:
                self._observer(self.station.position, t_monotonic=time.perf_counter())
            time.sleep(0.004)
        return MotionOutcome(target=float(position), actual=float(position), ok=True)

    def position(self):
        return int(self.station.position)


class SlowTurning:
    def __init__(self, station: Station, *, fail_at: float | None = None) -> None:
        self.station = station
        self.fail_at = fail_at
        self._observer = None

    @contextmanager
    def observe_positions(self, callback):
        self._observer = callback
        try:
            yield
        finally:
            self._observer = None

    def rotate_to(self, angle_deg):
        if self.fail_at is not None and float(angle_deg) == self.fail_at:
            return MotionOutcome(target=angle_deg, actual=self.station.angle, ok=False, detail="stall")
        start = self.station.angle
        for step in range(1, 13):
            with self.station.lock:
                self.station.angle = start + (angle_deg - start) * step / 12
            if self._observer is not None:
                self._observer(self.station.angle, t_monotonic=time.perf_counter())
            time.sleep(0.004)
        return MotionOutcome(target=angle_deg, actual=float(angle_deg), ok=True)

    def angle(self):
        return float(self.station.angle)

    def set_reference_angle(self, angle_deg):
        self.station.angle = float(angle_deg)


def _observations():
    return [
        counts_for((0.0, 0.0, 0.0), CALIBRATION),
        counts_for((1.0, 2.0, 3.0), CALIBRATION),
        counts_for((2.0, -1.0, 3.0), CALIBRATION),
        counts_for((-1.0, -2.0, 3.0), CALIBRATION),
        counts_for((-2.0, 1.0, 3.0), CALIBRATION),
        counts_for((0.0, 0.0, 0.0), CALIBRATION),
    ]


def _service(recorder=None, *, turning=None, station=None):
    station = station or Station()
    clock = FakeClock()
    service = BracketedAcquisitionService(
        FakeSquidTransport(_observations(), clock=clock),
        SlowVertical(station),
        turning or SlowTurning(station),
        config=AcquisitionConfig(zero_position=ZERO, measurement_position=MEAS, axis_calibration=CALIBRATION),
        clock=clock,
        id_factory=lambda prefix: f"{prefix}-mc",
        motion_capture=recorder,
    )
    return service, station


class MotionCaptureAcquisitionTests(unittest.TestCase):
    def test_without_capture_acquisition_is_unchanged(self):
        service, _ = _service()
        acquisition = service.acquire()
        self.assertEqual(acquisition.motion_traces, ())
        self.assertEqual(len(acquisition.block.observations), 6)

    def test_records_descent_three_turns_and_ascent(self):
        station = Station()
        recorder = MotionCaptureRecorder(StreamClient(station), MotionCaptureConfig(timing=FAST))
        service, _ = _service(recorder, station=station)
        acquisition = service.acquire()
        labels = [trace.label for trace in acquisition.motion_traces]
        self.assertEqual(
            labels,
            ["01-descent-measurement", "02-turn-90deg", "03-turn-180deg", "04-turn-270deg", "05-ascent-zero"],
        )
        for trace in acquisition.motion_traces:
            self.assertGreater(len(trace.samples), 3, trace.label)
            self.assertGreaterEqual(len(trace.positions), 10, trace.label)
            self.assertTrue(trace.motion["ok"])
            self.assertEqual(trace.errors, [])
        # Capture never touches the bracketed evidence or its reduction.
        baseline, _ = _service()
        expected = reduce_bracketed_measurement(baseline.acquire().block)
        actual = reduce_bracketed_measurement(acquisition.block)
        self.assertEqual(len(acquisition.block.observations), 6)
        self.assertEqual(actual.moment_emu, expected.moment_emu)
        self.assertEqual(actual.fischer_sd_deg, expected.fischer_sd_deg)

    def test_segments_can_be_switched_off(self):
        station = Station()
        recorder = MotionCaptureRecorder(
            StreamClient(station), MotionCaptureConfig(timing=FAST, capture_turns=False, capture_ascent=False)
        )
        service, _ = _service(recorder, station=station)
        labels = [trace.label for trace in service.acquire().motion_traces]
        self.assertEqual(labels, ["01-descent-measurement"])

    def test_stream_failure_is_recorded_and_block_still_completes(self):
        station = Station()
        recorder = MotionCaptureRecorder(StreamClient(station, fail=True), MotionCaptureConfig(timing=FAST))
        service, _ = _service(recorder, station=station)
        acquisition = service.acquire()
        self.assertEqual(len(acquisition.block.observations), 6)
        self.assertTrue(all(trace.stop_reason == "transport error" for trace in acquisition.motion_traces))
        self.assertIn("2G reply timed out", acquisition.motion_traces[0].errors[0])

    def test_link_that_cannot_be_released_fails_the_block(self):
        station = Station()
        gate = threading.Event()
        recorder = MotionCaptureRecorder(
            StreamClient(station, gate=gate), MotionCaptureConfig(timing=FAST), join_timeout_s=0.05
        )
        service, _ = _service(recorder, station=station)
        try:
            with self.assertRaises(MotionCaptureError):
                service.acquire()
        finally:
            gate.set()

    def test_motion_failure_still_closes_the_capture(self):
        station = Station()
        recorder = MotionCaptureRecorder(StreamClient(station), MotionCaptureConfig(timing=FAST))
        service, _ = _service(recorder, station=station, turning=SlowTurning(station, fail_at=180.0))
        with self.assertRaises(MotionVerificationError):
            service.acquire()
        # acquire() closed the block itself; the failed 180-degree turn is kept.
        traces = recorder.end_block(completed=False)
        self.assertEqual([t.label for t in traces], ["01-descent-measurement", "02-turn-90deg", "03-turn-180deg"])
        self.assertFalse(traces[-1].motion["ok"])
        self.assertEqual(len(recorder.analyses), 3)

    def test_sidecar_files_include_physical_fits(self):
        station = Station()
        with tempfile.TemporaryDirectory() as tmp:
            recorder = MotionCaptureRecorder(StreamClient(station), MotionCaptureConfig(timing=FAST, output_dir=tmp))
            service, _ = _service(recorder, station=station)
            service.acquire()
            folder = recorder.last_written
            self.assertIsNotNone(folder)
            self.assertEqual(len(list(folder.glob("*.csv"))), 5)
            payload = json.loads((folder / "analysis.json").read_text(encoding="utf-8"))
            self.assertTrue(payload["completed"])
            models = [segment["fit"].get("model") for segment in payload["segments"]]
            self.assertEqual(models, ["pass_through", "rotation", "rotation", "rotation", "pass_through"])
            # The fit physics is checked deterministically in test_signal_analysis;
            # here the integration must produce a finite peak fit per axis.
            # A loaded test machine may yield too few descent samples for a fit,
            # which is reported per axis rather than raised.
            descent_z = payload["segments"][0]["fit"]["axes"]["Z"]
            if "error" in descent_z:
                self.assertIn("eight", descent_z["error"])
            else:
                self.assertTrue(math.isfinite(descent_z["amplitude"]))
                self.assertGreaterEqual(descent_z["samples"], 8)
            turn = payload["segments"][1]["fit"]
            self.assertAlmostEqual(turn["angle_span_deg"], 90.0, delta=10.0)
            self.assertAlmostEqual(turn["horizontal_amplitude"], 2.0, delta=0.2)


class MotorObserverTests(unittest.TestCase):
    class _Client:
        class config:
            turning_motor_full_rotation = 2000

        def __init__(self):
            self.observer = None

        def set_position_observer(self, observer):
            self.observer = observer

    class _Axis:
        def __init__(self, motor_id):
            self.motor_id = motor_id

    def test_vertical_controller_forwards_only_its_axis(self):
        client = self._Client()
        controller = MotorVerticalController(client, self._Axis(3))
        seen = []
        with controller.observe_positions(lambda value, t_monotonic: seen.append((value, t_monotonic))):
            client.observer(self._Axis(3), 5.0, -1200)
            client.observer(self._Axis(2), 5.1, 77)
        self.assertEqual(seen, [(-1200.0, 5.0)])
        self.assertIsNone(client.observer)

    def test_turning_controller_reports_degrees(self):
        client = self._Client()
        controller = MotorTurningController(client, self._Axis(2))
        seen = []
        with controller.observe_positions(lambda value, t_monotonic: seen.append(value)):
            client.observer(self._Axis(2), 1.0, -500)
        self.assertEqual(len(seen), 1)
        self.assertAlmostEqual(abs(seen[0]) % 360, 90.0)

    def test_hardware_wait_loop_notifies_and_swallows_observer_errors(self):
        from rapidpy_common.hardware import MotorSerialClient

        client = MotorSerialClient.__new__(MotorSerialClient)
        calls = []
        client.set_position_observer(lambda axis, t, pos: calls.append(pos))
        client._notify_position("axis", 42)
        client.set_position_observer(lambda axis, t, pos: 1 / 0)
        client._notify_position("axis", 43)  # must not raise
        self.assertEqual(calls, [42])


class ConfigMappingTests(unittest.TestCase):
    def test_capture_is_off_by_default(self):
        self.assertIsNone(motion_capture_config_from_app_config(AppConfig()))

    def test_enabled_capture_maps_rate_axes_and_folder(self):
        config = AppConfig()
        config.squid.motion_capture_enabled = True
        config.squid.stream_interval_ms = 50
        config.squid.stream_axes = "zx"
        config.squid.stream_counts_every = 4
        config.squid.motion_capture_turns = False
        config.general.data_dir = str(Path(tempfile.gettempdir()) / "rapid-data")
        plan = motion_capture_config_from_app_config(config)
        self.assertEqual(plan.timing.interval_s, 0.05)
        self.assertEqual(plan.timing.axes, ("X", "Z"))
        self.assertEqual(plan.timing.counts_every, 4)
        self.assertFalse(plan.capture_turns)
        self.assertTrue(plan.output_dir.endswith("squid_motion_capture"))

    def test_timing_defaults_keep_vb6_latch_holds(self):
        timing = stream_timing_from_squid_config(AppConfig().squid)
        self.assertEqual((timing.latch_count_hold_s, timing.latch_data_hold_s), (0.1, 0.12))
        self.assertEqual(timing.interval_s, 0.5)


if __name__ == "__main__":
    unittest.main()


class Vb6SquidDefaultsTests(unittest.TestCase):
    """VB6 frmSQUID.Connect hardcodes 1200,N,8,1; the station INI sets ReadDelay=1."""

    def test_defaults_follow_vb6_station(self):
        from rapid_main.config import SquidConfig

        cfg = SquidConfig()
        self.assertEqual((cfg.port, cfg.baud, cfg.settle_time, cfg.samples_per_pos), ("COM1", 1200, 1.0, 1))

    def test_damaged_range_label_is_repaired_on_load(self):
        from rapid_main.config import SquidConfig

        self.assertEqual(SquidConfig(range_label="1�").range_label, "1×")
        self.assertEqual(SquidConfig(range_label="1Ã—").range_label, "1×")
        self.assertEqual(SquidConfig(range_label="100×").range_label, "100×")

    def test_legacy_import_sets_hardcoded_vb6_baud(self):
        from rapid_main.legacy_ini import import_vb6_ini

        with tempfile.TemporaryDirectory() as tmp:
            ini = Path(tmp) / "Paleomag.INI"
            ini.write_text("[COMPorts]\nCOMPortSquids= 1\n[MagnetometerCalibration]\nReadDelay= 1\n", encoding="utf-8")
            config = AppConfig()
            config.squid.baud = 9600
            import_vb6_ini(config, ini)
        self.assertEqual((config.squid.port, config.squid.baud, config.squid.settle_time), ("COM1", 1200, 1.0))
