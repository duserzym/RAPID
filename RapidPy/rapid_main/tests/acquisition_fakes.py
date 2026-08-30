"""Deterministic fakes for the bracketed SQUID acquisition state machine.

These mirror the 2G 581 / QuickSilver command surface closely enough to assert
ordering, evidence retention, and fault handling without hardware.
"""
from __future__ import annotations

from datetime import datetime, timedelta, timezone
import itertools

from rapid_main.acquisition import AxisReply, LatchResult, MotionOutcome

AxisPair = tuple[float, float]
ObservationCounts = tuple[AxisPair, AxisPair, AxisPair]


class FakeClock:
    """Wall/monotonic clock that only advances when told to."""

    def __init__(self, *, start: datetime | None = None, tick_s: float = 0.001) -> None:
        self._now = start or datetime(2026, 8, 29, 12, 0, 0, tzinfo=timezone.utc)
        self._monotonic = 1000.0
        self._tick_s = float(tick_s)
        self.slept: list[float] = []

    def now(self) -> datetime:
        self._now += timedelta(seconds=self._tick_s)
        self._monotonic += self._tick_s
        return self._now

    def monotonic(self) -> float:
        self._monotonic += self._tick_s
        return self._monotonic

    def sleep(self, seconds: float) -> None:
        self.slept.append(float(seconds))
        self._now += timedelta(seconds=float(seconds))
        self._monotonic += float(seconds)

    def advance(self, seconds: float) -> None:
        self._now += timedelta(seconds=float(seconds))
        self._monotonic += float(seconds)


class FakeSquidTransport:
    """Programmable 2G transport that hands back one observation at a time."""

    def __init__(
        self,
        observations: list[ObservationCounts],
        *,
        range_value: float = 1.0,
        mismatch_axis: str | None = None,
        read_error: Exception | None = None,
        error_on_read_index: int | None = None,
        latch_error: Exception | None = None,
        reset_error: Exception | None = None,
        blank_reply_axis: str | None = None,
        stale_seconds: float = 0.0,
        clock: FakeClock | None = None,
    ) -> None:
        self._observations = list(observations)
        self._range_value = float(range_value)
        self._mismatch_axis = mismatch_axis
        self._read_error = read_error
        self._error_on_read_index = error_on_read_index
        self._latch_error = latch_error
        self._reset_error = reset_error
        self._blank_reply_axis = blank_reply_axis
        self._stale_seconds = float(stale_seconds)
        self._clock = clock
        self._latch_ids = itertools.count(1)
        self._current_latch = ""
        self._observation_index = -1
        self._read_index = 0
        self.commands: list[str] = []
        self.ranges: list[tuple[str, str]] = []
        self.reset_count = 0
        self.latch_count = 0

    def clear_and_reset_counters(self, axis: str = "A") -> None:
        if self._reset_error is not None:
            raise self._reset_error
        self.reset_count += 1
        self.commands.append(f"{axis}CLP")
        self.commands.append(f"{axis}RC")

    def set_range(self, axis: str, range_label: str) -> None:
        self.ranges.append((axis, range_label))
        self.commands.append(f"{axis}CR{range_label}")

    def latch(self, axis: str = "A", *, settle: bool = False) -> LatchResult:
        if self._latch_error is not None:
            raise self._latch_error
        self.latch_count += 1
        self._current_latch = f"latch-{next(self._latch_ids):03d}"
        self._observation_index += 1
        self.commands.append(f"{axis}LC")
        self.commands.append(f"{axis}LD")
        if self._stale_seconds and self._clock is not None:
            self._clock.advance(self._stale_seconds)
        return LatchResult(latch_id=self._current_latch, commands=(f"{axis}LC", f"{axis}LD"))

    def read_axis(self, axis: str) -> AxisReply:
        if (
            self._read_error is not None
            and self._error_on_read_index is not None
            and self._read_index == self._error_on_read_index
        ):
            raise self._read_error
        self._read_index += 1
        index = min(self._observation_index, len(self._observations) - 1)
        if index < 0:
            raise AssertionError("read_axis called before latch")
        counts, dvm = self._observations[index]["XYZ".index(axis)]
        latch_id = self._current_latch
        if self._mismatch_axis == axis:
            latch_id = "latch-bogus"
        count_reply = "" if self._blank_reply_axis == axis else f"{counts:g}"
        self.commands.append(f"{axis}SC")
        self.commands.append(f"{axis}SD")
        return AxisReply(
            counts=float(counts),
            dvm=float(dvm),
            latch_id=latch_id,
            range_value=self._range_value,
            count_command=f"{axis}SC",
            count_reply=count_reply,
            data_command=f"{axis}SD",
            data_reply=f"{dvm:g}",
        )


class FakeVertical:
    """Lift axis that reports the commanded position unless told to fail."""

    def __init__(self, *, start: int = 0, fail_at_target: int | None = None, slop: int = 0) -> None:
        self._position = int(start)
        self._fail_at_target = fail_at_target
        self._slop = int(slop)
        self.moves: list[tuple[int, int]] = []

    def move_to(self, position: int, *, speed_index: int = 0) -> MotionOutcome:
        self.moves.append((int(position), int(speed_index)))
        if self._fail_at_target is not None and int(position) == int(self._fail_at_target):
            self._position = int(position) + 5000
            return MotionOutcome(
                target=float(position),
                actual=float(self._position),
                ok=False,
                detail="unacceptable slop on up/down motor",
            )
        self._position = int(position) + self._slop
        return MotionOutcome(target=float(position), actual=float(self._position), ok=True)

    def position(self) -> int:
        return int(self._position)


class FakeTurning:
    """Turning axis that reports the commanded angle unless told to fail."""

    def __init__(self, *, fail_at_angle: float | None = None) -> None:
        self._angle = 0.0
        self._reference = 0.0
        self._fail_at_angle = fail_at_angle
        self.angles: list[float] = []
        self.references: list[float] = []

    def rotate_to(self, angle_deg: float) -> MotionOutcome:
        self.angles.append(float(angle_deg))
        if self._fail_at_angle is not None and float(angle_deg) == float(self._fail_at_angle):
            return MotionOutcome(
                target=float(angle_deg),
                actual=float(self._angle),
                ok=False,
                detail="turn overshoot beyond 5 degrees",
            )
        self._angle = float(angle_deg)
        return MotionOutcome(target=float(angle_deg), actual=self._angle, ok=True)

    def angle(self) -> float:
        return float(self._angle)

    def set_reference_angle(self, angle_deg: float) -> None:
        self.references.append(float(angle_deg))
        self._angle = float(angle_deg)


def counts_for(vector: tuple[float, float, float], calibration: tuple[float, float, float]) -> ObservationCounts:
    """Build (counts, dvm) pairs whose calibrated value equals ``vector``.

    ``raw = -dvm - counts * range`` and ``calibrated = raw * calibration``.
    With ``counts = 0`` this reduces to ``dvm = -value / calibration``.
    """

    pairs = []
    for value, cal in zip(vector, calibration):
        pairs.append((0.0, -float(value) / float(cal)))
    return (pairs[0], pairs[1], pairs[2])


def counts_with_step(
    vector: tuple[float, float, float],
    calibration: tuple[float, float, float],
    *,
    axis: int,
    steps: int = 1,
) -> ObservationCounts:
    """Like :func:`counts_for` but with ``steps`` extra flux counts on ``axis``.

    In VB6 flux-counting mode ``rangeval`` is ``1``, so one counter increment
    moves the calibrated value by exactly the axis calibration factor.
    """

    pairs = list(counts_for(vector, calibration))
    counts, dvm = pairs[axis]
    pairs[axis] = (counts + float(steps), dvm)
    return (pairs[0], pairs[1], pairs[2])


class FakeRawSquidClient:
    """Stand-in for the extended ``updown_control.app.RawSquidClient``."""

    def __init__(self, observations, *, connected: bool = True) -> None:
        self.is_connected = connected
        self._observations = list(observations)
        self.calls: list[tuple[str, object]] = []
        self._index = -1

    def clear_and_reset(self, axis: str = "A") -> tuple[str, ...]:
        self.calls.append(("clear_and_reset", axis))
        return (f"{axis}CLP", f"{axis}RC")

    def set_range(self, axis: str = "A", range_label: str = "1") -> tuple[str, ...]:
        self.calls.append(("set_range", (axis, range_label)))
        return (f"{axis}CR{range_label}",)

    def latch(self, axis: str = "A", *, settle_s: float = 0.0) -> tuple[str, ...]:
        self.calls.append(("latch", (axis, settle_s)))
        self._index += 1
        return (f"{axis}LC", f"{axis}LD")

    def read_axis(self, axis: str, *, range_value: float = 1.0) -> "FakeAxisSample":
        self.calls.append(("read_axis", (axis, range_value)))
        index = min(max(self._index, 0), len(self._observations) - 1)
        counts, dvm = self._observations[index]["XYZ".index(axis)]
        return FakeAxisSample(
            axis=axis,
            counts=float(counts),
            dvm=float(dvm),
            range_value=float(range_value),
            count_command=f"{axis}SC",
            count_reply=f"{counts:g}",
            data_command=f"{axis}SD",
            data_reply=f"{dvm:g}",
        )


class FakeAxisSample:
    """Minimal duck-type of ``updown_control.app.SquidAxisSample``."""

    def __init__(
        self,
        *,
        axis: str,
        counts: float,
        dvm: float,
        range_value: float,
        count_command: str,
        count_reply: str,
        data_command: str,
        data_reply: str,
    ) -> None:
        self.axis = axis
        self.counts = counts
        self.dvm = dvm
        self.range_value = range_value
        self.count_command = count_command
        self.count_reply = count_reply
        self.data_command = data_command
        self.data_reply = data_reply

    @property
    def raw_value(self) -> float:
        return -self.dvm - self.counts * self.range_value


class FakeMotorSerialClient:
    """QuickSilver client stand-in for queue-backend composition tests."""

    def __init__(self, config=None) -> None:
        from rapidpy_common.hardware import MotorControllerConfig, MoveResult

        self.config = config or MotorControllerConfig()
        self._MoveResult = MoveResult
        self.is_connected = False
        self.positions: dict[int, int] = {1: 0, 2: 0, 3: 0, 4: 0}
        self.calls: list[tuple[str, object]] = []
        self.turn_failure_angle: float | None = None
        self.lift_failure_target: int | None = None

    def connect(self, port: str, baudrate: int = 57600, timeout: float = 0.35) -> None:
        self.calls.append(("connect", (port, baudrate)))
        self.is_connected = True

    def read_position(self, axis) -> int:
        return int(self.positions[axis.motor_id])

    def updown_move(self, axis, target: int, speed_index: int, wait_for_stop: bool = True):
        self.calls.append(("updown_move", (int(target), int(speed_index))))
        if self.lift_failure_target is not None and int(target) == int(self.lift_failure_target):
            return self._MoveResult(target=int(target), final_position=int(target) + 9999, success=False)
        self.positions[axis.motor_id] = int(target)
        return self._MoveResult(target=int(target), final_position=int(target), success=True)

    def turning_motor_rotate(self, axis, angle: float, wait_for_stop: bool = True):
        from rapidpy_common.hardware import convert_angle_to_pos

        self.calls.append(("turning_motor_rotate", float(angle)))
        if self.turn_failure_angle is not None and float(angle) == float(self.turn_failure_angle):
            return self._MoveResult(
                target=int(convert_angle_to_pos(angle, self.config.turning_motor_full_rotation)),
                final_position=int(self.positions[axis.motor_id]),
                success=False,
            )
        position = int(convert_angle_to_pos(angle, self.config.turning_motor_full_rotation))
        self.positions[axis.motor_id] = position
        return self._MoveResult(target=position, final_position=position, success=True)

    def relabel_pos(self, axis, pos: int, tolerance: int = 10, max_cycles: int = 20) -> None:
        self.calls.append(("relabel_pos", int(pos)))
        self.positions[axis.motor_id] = int(pos)

    def changer_motor_to_hole(self, axis, hole: float, wait_for_stop: bool = True):
        self.calls.append(("changer_motor_to_hole", float(hole)))
        return self._MoveResult(target=int(hole), final_position=int(hole), success=True)

    def sample_pickup(self, axis):
        self.calls.append(("sample_pickup", axis.name))
        return self._MoveResult(target=0, final_position=0, success=True)

    def sample_dropoff(self, axis, use_xy_table: bool = True):
        self.calls.append(("sample_dropoff", axis.name))
        return self._MoveResult(target=0, final_position=0, success=True)

    def home_to_top(self, axis):
        self.calls.append(("home_to_top", axis.name))
        return self._MoveResult(target=0, final_position=0, success=True)

    def halt(self, axis) -> str:
        self.calls.append(("halt", axis.name))
        return "halted"
