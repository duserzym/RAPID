"""Paced, continuous SQUID sampling for the 2G 581 serial magnetometer.

The bracketed measurement reads the SQUID only at rest: one latch, then the
counter and DVM of each axis.  This module reuses those same atomic client
operations (``latch`` / ``read_axis`` / ``read_axis_data``) to take a *stream*
of readings at an operator-chosen rate, so the signal can also be recorded
while the specimen travels down the borehole and while it turns between
orientations.

Rate is bounded by the serial link, not by software: every sample costs a
latch (two commands plus the legacy hold delays) and one or two query
round-trips per axis.  :func:`estimate_sample_period_s` reports that floor so
the UI can show what a requested rate will really achieve, and every trace
records the achieved rate rather than the requested one.

Nothing here is used to produce published scientific values.  Streams are
auxiliary evidence for noise characterisation and spectral analysis; the
bracketed acquisition remains the only source of accepted measurements.
"""

from __future__ import annotations

import csv
import io
import json
import math
import threading
import time
from dataclasses import asdict, dataclass, field
from datetime import datetime, timezone
from typing import Callable, Iterable, Protocol

AXES = ("X", "Y", "Z")

#: VB6 ``frmSQUID.LatchCount`` / ``LatchData`` pauses.
VB6_LATCH_COUNT_HOLD_S = 0.10
VB6_LATCH_DATA_HOLD_S = 0.12


class SquidStreamError(RuntimeError):
    """Raised when a stream cannot be configured or safely stopped."""


class StreamingSquidClient(Protocol):
    """The atomic 2G operations a stream needs (``updown_control.RawSquidClient``)."""

    def latch(self, axis: str = "A", *, settle_s: float = 0.0, **holds: float) -> tuple[str, ...]: ...

    def read_axis(self, axis: str, *, range_value: float = 1.0): ...


@dataclass(frozen=True)
class SquidReadTiming:
    """Operator-controlled SQUID read rate and per-sample read plan.

    ``interval_s`` is the requested period between sample *starts*; ``0``
    means "as fast as the link allows".  ``counts_every`` re-reads the flux
    counter on every Nth sample and reuses the last counter value in between,
    which roughly halves the per-axis cost.  Samples that reuse a counter are
    flagged so analysis can discard them if a flux jump is suspected.
    """

    interval_s: float = 0.5
    axes: tuple[str, ...] = AXES
    counts_every: int = 1
    latch_count_hold_s: float = VB6_LATCH_COUNT_HOLD_S
    latch_data_hold_s: float = VB6_LATCH_DATA_HOLD_S
    reply_timeout_s: float = 1.0

    def __post_init__(self) -> None:
        axes = tuple(str(axis).strip().upper() for axis in self.axes)
        if not axes or any(axis not in AXES for axis in axes) or len(set(axes)) != len(axes):
            raise SquidStreamError(f"Stream axes must be a non-empty subset of X/Y/Z, got {self.axes!r}")
        object.__setattr__(self, "axes", tuple(axis for axis in AXES if axis in axes))
        for name in ("interval_s", "latch_count_hold_s", "latch_data_hold_s"):
            value = float(getattr(self, name))
            if not math.isfinite(value) or value < 0.0 or value > 60.0:
                raise SquidStreamError(f"{name} must be within 0..60 s, got {value!r}")
        timeout = float(self.reply_timeout_s)
        if not math.isfinite(timeout) or not 0.05 <= timeout <= 10.0:
            raise SquidStreamError(f"reply_timeout_s must be within 0.05..10 s, got {timeout!r}")
        if int(self.counts_every) != self.counts_every or not 1 <= int(self.counts_every) <= 1000:
            raise SquidStreamError("counts_every must be an integer from 1 to 1000")

    @property
    def requested_rate_hz(self) -> float:
        return math.inf if self.interval_s <= 0 else 1.0 / self.interval_s

    def to_dict(self) -> dict:
        payload = asdict(self)
        payload["axes"] = list(self.axes)
        return payload

    @classmethod
    def from_dict(cls, payload: dict) -> "SquidReadTiming":
        known = {name: payload[name] for name in cls.__dataclass_fields__ if name in payload}
        if "axes" in known:
            known["axes"] = tuple(known["axes"])
        return cls(**known)


def estimate_query_s(baud: int, *, command_chars: int = 5, reply_chars: int = 12, turnaround_s: float = 0.004) -> float:
    """Approximate one 2G query round-trip at ``baud`` with 8N1 framing."""

    if baud <= 0:
        raise SquidStreamError("baud must be positive")
    return (command_chars + reply_chars) * 10.0 / float(baud) + turnaround_s


def estimate_sample_period_s(timing: SquidReadTiming, baud: int) -> float:
    """Lower bound on one streamed sample's duration for this read plan.

    Averages the counter re-read over ``counts_every`` samples.  Real links
    add controller latency, so the achieved period is reported per trace.
    """

    query = estimate_query_s(baud)
    latch_commands = 2 * 5 * 10.0 / float(baud)
    queries = len(timing.axes) * (1.0 + 1.0 / float(timing.counts_every))
    return timing.latch_count_hold_s + timing.latch_data_hold_s + latch_commands + queries * query


def effective_period_s(timing: SquidReadTiming, baud: int) -> float:
    """The period the stream will actually run at: requested or link-limited."""

    return max(float(timing.interval_s), estimate_sample_period_s(timing, baud))


@dataclass(frozen=True)
class SquidStreamSample:
    """One latched multi-axis sample.  Unread axes hold ``nan``."""

    index: int
    t_s: float
    duration_s: float
    raw: tuple[float, float, float]
    counts: tuple[float, float, float]
    dvm: tuple[float, float, float]
    counts_fresh: bool

    def as_row(self) -> list:
        return [
            self.index,
            f"{self.t_s:.6f}",
            f"{self.duration_s:.6f}",
            *(_fmt(v) for v in self.raw),
            *(_fmt(v) for v in self.counts),
            *(_fmt(v) for v in self.dvm),
            int(self.counts_fresh),
        ]


def _fmt(value: float) -> str:
    return "nan" if not math.isfinite(value) else repr(float(value))


@dataclass(frozen=True)
class PositionSample:
    """A motor encoder position observed while a capture was running."""

    t_s: float
    position: float


@dataclass
class SquidTrace:
    """A labelled continuous SQUID record with optional motion context.

    ``t_s`` values are seconds since ``started_monotonic``; motor positions
    observed during the same interval share that time base, so analysis can
    place every SQUID sample on the lift height or turn angle it was taken at.
    The time base is ``time.perf_counter`` because ``time.monotonic`` ticks
    only every ~15.6 ms on Windows, coarser than a fast stream.
    """

    label: str
    timing: SquidReadTiming
    range_value: float = 1.0
    started_iso: str = ""
    started_monotonic: float = 0.0
    stopped_monotonic: float = 0.0
    samples: list[SquidStreamSample] = field(default_factory=list)
    positions: list[PositionSample] = field(default_factory=list)
    motion: dict = field(default_factory=dict)
    errors: list[str] = field(default_factory=list)
    stop_reason: str = ""

    @property
    def duration_s(self) -> float:
        return max(0.0, self.stopped_monotonic - self.started_monotonic)

    @property
    def achieved_rate_hz(self) -> float:
        if len(self.samples) < 2:
            return 0.0
        span = self.samples[-1].t_s - self.samples[0].t_s
        return (len(self.samples) - 1) / span if span > 0 else 0.0

    def times(self) -> list[float]:
        return [sample.t_s for sample in self.samples]

    def axis_values(self, axis: str) -> list[float]:
        index = AXES.index(str(axis).upper())
        return [sample.raw[index] for sample in self.samples]

    def summary(self) -> dict:
        return {
            "label": self.label,
            "started_iso": self.started_iso,
            "duration_s": round(self.duration_s, 6),
            "samples": len(self.samples),
            "positions": len(self.positions),
            "requested_rate_hz": None if math.isinf(self.timing.requested_rate_hz) else self.timing.requested_rate_hz,
            "achieved_rate_hz": round(self.achieved_rate_hz, 6),
            "stop_reason": self.stop_reason,
            "errors": list(self.errors),
            "motion": dict(self.motion),
            "timing": self.timing.to_dict(),
            "range_value": self.range_value,
        }

    CSV_HEADER = (
        "index", "t_s", "duration_s",
        "raw_x", "raw_y", "raw_z",
        "counts_x", "counts_y", "counts_z",
        "dvm_x", "dvm_y", "dvm_z",
        "counts_fresh",
    )

    def to_csv(self) -> str:
        buffer = io.StringIO()
        writer = csv.writer(buffer, lineterminator="\n")
        buffer.write("# " + json.dumps(self.summary(), sort_keys=True) + "\n")
        writer.writerow(self.CSV_HEADER)
        for sample in self.samples:
            writer.writerow(sample.as_row())
        if self.positions:
            buffer.write("# positions\n")
            writer.writerow(("t_s", "position"))
            for position in self.positions:
                writer.writerow((f"{position.t_s:.6f}", _fmt(position.position)))
        return buffer.getvalue()

    @classmethod
    def from_csv(cls, text: str) -> "SquidTrace":
        lines = text.splitlines()
        if not lines or not lines[0].startswith("# "):
            raise SquidStreamError("Trace CSV is missing its summary header line")
        summary = json.loads(lines[0][2:])
        trace = cls(
            label=str(summary.get("label", "")),
            timing=SquidReadTiming.from_dict(summary.get("timing", {})),
            range_value=float(summary.get("range_value", 1.0)),
            started_iso=str(summary.get("started_iso", "")),
            motion=dict(summary.get("motion", {})),
            errors=list(summary.get("errors", [])),
            stop_reason=str(summary.get("stop_reason", "")),
        )
        section = "samples"
        for line in lines[1:]:
            if not line.strip():
                continue
            if line.startswith("# positions"):
                section = "positions"
                continue
            row = next(csv.reader([line]))
            if row[0] in ("index", "t_s"):
                continue
            if section == "samples":
                values = [float(v) for v in row[3:12]]
                trace.samples.append(
                    SquidStreamSample(
                        index=int(row[0]),
                        t_s=float(row[1]),
                        duration_s=float(row[2]),
                        raw=tuple(values[0:3]),  # type: ignore[arg-type]
                        counts=tuple(values[3:6]),  # type: ignore[arg-type]
                        dvm=tuple(values[6:9]),  # type: ignore[arg-type]
                        counts_fresh=bool(int(row[12])),
                    )
                )
            else:
                trace.positions.append(PositionSample(t_s=float(row[0]), position=float(row[1])))
        if trace.samples:
            trace.started_monotonic = 0.0
            trace.stopped_monotonic = float(summary.get("duration_s", trace.samples[-1].t_s))
        return trace


class SquidStreamSampler:
    """Take paced latched readings through an atomic 2G client.

    The sampler never raises out of :meth:`run`; a transport failure ends the
    stream and is recorded on the trace, because a stream is diagnostic and
    must not take a measurement down with it.  Callers that need the error
    can inspect ``trace.errors``.
    """

    def __init__(
        self,
        client: StreamingSquidClient,
        timing: SquidReadTiming,
        *,
        range_value: float = 1.0,
        monotonic: Callable[[], float] = time.perf_counter,
        sleep: Callable[[float], None] = time.sleep,
    ) -> None:
        self._client = client
        self._timing = timing
        self._range_value = float(range_value)
        self._monotonic = monotonic
        self._sleep = sleep
        self._last_counts = [math.nan, math.nan, math.nan]

    @property
    def timing(self) -> SquidReadTiming:
        return self._timing

    def _wait(self, stop_event: threading.Event, seconds: float) -> None:
        if self._sleep is time.sleep:
            stop_event.wait(seconds)
        else:
            self._sleep(seconds)

    def _latch(self) -> None:
        holds = {
            "count_hold_s": self._timing.latch_count_hold_s,
            "data_hold_s": self._timing.latch_data_hold_s,
        }
        try:
            self._client.latch("A", settle_s=0.0, **holds)
        except TypeError:
            # Clients without per-call hold overrides keep their fixed VB6 holds.
            self._client.latch("A", settle_s=0.0)

    def _read_data_only(self, axis: str) -> float:
        reader = getattr(self._client, "read_axis_data", None)
        if callable(reader):
            return float(reader(axis))
        return float(self._client.read_axis(axis, range_value=self._range_value).dvm)

    def sample_once(self, index: int, t0: float) -> SquidStreamSample:
        started = self._monotonic()
        refresh_counts = index % self._timing.counts_every == 0 or any(
            math.isnan(self._last_counts[AXES.index(axis)]) for axis in self._timing.axes
        )
        self._latch()
        raw = [math.nan, math.nan, math.nan]
        counts = [math.nan, math.nan, math.nan]
        dvm = [math.nan, math.nan, math.nan]
        for axis in self._timing.axes:
            slot = AXES.index(axis)
            if refresh_counts:
                reading = self._client.read_axis(axis, range_value=self._range_value)
                counts[slot] = float(reading.counts)
                dvm[slot] = float(reading.dvm)
                self._last_counts[slot] = counts[slot]
            else:
                counts[slot] = self._last_counts[slot]
                dvm[slot] = self._read_data_only(axis)
            # VB6 getVal: -data - count * range
            raw[slot] = -dvm[slot] - counts[slot] * self._range_value
        finished = self._monotonic()
        return SquidStreamSample(
            index=index,
            t_s=started - t0,
            duration_s=finished - started,
            raw=tuple(raw),  # type: ignore[arg-type]
            counts=tuple(counts),  # type: ignore[arg-type]
            dvm=tuple(dvm),  # type: ignore[arg-type]
            counts_fresh=refresh_counts,
        )

    def run(
        self,
        trace: SquidTrace,
        *,
        stop_event: threading.Event | None = None,
        max_samples: int | None = None,
        duration_s: float | None = None,
        on_sample: Callable[[SquidStreamSample], None] | None = None,
    ) -> SquidTrace:
        stop_event = stop_event or threading.Event()
        t0 = trace.started_monotonic or self._monotonic()
        trace.started_monotonic = t0
        if not trace.started_iso:
            trace.started_iso = datetime.now(timezone.utc).isoformat()
        index = 0
        next_start = t0
        try:
            while True:
                if stop_event.is_set():
                    trace.stop_reason = trace.stop_reason or "stopped"
                    break
                if max_samples is not None and index >= max_samples:
                    trace.stop_reason = "sample limit"
                    break
                now = self._monotonic()
                if duration_s is not None and now - t0 >= duration_s:
                    trace.stop_reason = "duration limit"
                    break
                wait = next_start - now
                if wait > 0:
                    # Wake in short slices so stop requests are honoured promptly.
                    self._wait(stop_event, min(wait, 0.05))
                    continue
                sample = self.sample_once(index, t0)
                trace.samples.append(sample)
                if on_sample is not None:
                    on_sample(sample)
                index += 1
                # Pace from this sample's start; a slow link simply runs back to
                # back at its own limit instead of bursting to "catch up".
                next_start = t0 + sample.t_s + self._timing.interval_s
        except Exception as exc:  # transport failures end the stream, never the caller
            trace.errors.append(f"{type(exc).__name__}: {exc}")
            trace.stop_reason = "transport error"
        trace.stopped_monotonic = self._monotonic()
        return trace


class BackgroundSquidCapture:
    """Run a :class:`SquidStreamSampler` on a worker thread for one interval.

    The capture owns the SQUID serial link while it runs.  :meth:`stop`
    always joins the worker before returning, so the caller can safely issue
    its own SQUID commands afterwards; if the worker cannot be joined within
    ``join_timeout_s`` the link is considered unsafe and
    :class:`SquidStreamError` is raised instead of returning.
    """

    def __init__(
        self,
        sampler: SquidStreamSampler,
        label: str,
        *,
        motion: dict | None = None,
        join_timeout_s: float | None = None,
        on_sample: Callable[[SquidStreamSample], None] | None = None,
        max_samples: int | None = None,
    ) -> None:
        self._sampler = sampler
        self._stop = threading.Event()
        timing = sampler.timing
        self._join_timeout_s = (
            float(join_timeout_s)
            if join_timeout_s is not None
            # Worst case for the in-flight sample: every query waits out its reply
            # deadline plus one serial read timeout before the stop flag is seen.
            else 2.0 + timing.latch_count_hold_s + timing.latch_data_hold_s
            + 2 * len(timing.axes) * (timing.reply_timeout_s + 1.0)
        )
        self.trace = SquidTrace(label=label, timing=timing, range_value=sampler._range_value, motion=dict(motion or {}))
        self._on_sample = on_sample
        self._max_samples = max_samples
        self._thread = threading.Thread(target=self._run, name=f"squid-capture-{label}", daemon=True)
        self._started = False

    def _run(self) -> None:
        self._sampler.run(self.trace, stop_event=self._stop, on_sample=self._on_sample, max_samples=self._max_samples)

    def start(self) -> "BackgroundSquidCapture":
        self.trace.started_monotonic = time.perf_counter()
        self.trace.started_iso = datetime.now(timezone.utc).isoformat()
        self._thread.start()
        self._started = True
        return self

    def mark_position(self, position: float, *, t_monotonic: float | None = None) -> None:
        """Record a motor position on the capture's time base (thread-safe append)."""

        t = time.perf_counter() if t_monotonic is None else float(t_monotonic)
        self.trace.positions.append(PositionSample(t_s=t - self.trace.started_monotonic, position=float(position)))

    @property
    def running(self) -> bool:
        return self._thread.is_alive()

    def stop(self, reason: str = "motion complete") -> SquidTrace:
        if not self._started:
            return self.trace
        if not self.trace.stop_reason:
            self.trace.stop_reason = reason
        self._stop.set()
        self._thread.join(self._join_timeout_s)
        if self._thread.is_alive():
            raise SquidStreamError(
                f"SQUID capture '{self.trace.label}' did not release the serial link within "
                f"{self._join_timeout_s:.1f} s; the link is not safe to reuse."
            )
        return self.trace

    def __enter__(self) -> "BackgroundSquidCapture":
        return self.start()

    def __exit__(self, exc_type, exc, tb) -> bool:  # noqa: ANN001 - context protocol
        self.stop("motion failed" if exc_type is not None else "motion complete")
        return False


def iter_axis_series(trace: SquidTrace, axes: Iterable[str] = AXES):
    """Yield ``(axis, times, values)`` for axes that were actually read."""

    times = trace.times()
    for axis in axes:
        if axis.upper() in trace.timing.axes:
            yield axis.upper(), times, trace.axis_values(axis)
