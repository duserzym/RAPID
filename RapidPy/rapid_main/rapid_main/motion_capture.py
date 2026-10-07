"""Continuous SQUID capture while the specimen moves inside a bracketed block.

The bracketed acquisition reads the SQUID only with the specimen at rest.  The
specimen also spends seconds travelling down the borehole into the pickup
coils and turning 90 degrees between orientations; this module records the
SQUID throughout those motions so the extra information can be used for
spectral characterisation, filtering, and model fits
(:mod:`rapidpy_common.signal_analysis`).

Safety and evidence boundaries
------------------------------

* Capture is **opt-in** (``SquidConfig.motion_capture_enabled``) and only runs
  while a lift or turn is in progress.  The motors use a different serial port,
  so the stream and the motion never share a link.
* Each capture is stopped and its worker thread joined *before* the
  acquisition issues its next SQUID command.  A worker that cannot release the
  link fails the block (:class:`MotionCaptureError`) rather than risking
  interleaved serial traffic.
* A stream failure is recorded on the trace and never fails the block.
* Traces are auxiliary evidence: they do not enter the six bracketed
  observations, the reduction, holder correction, or published results.
  Sidecar files are written under their own folder, outside the scientific
  bundle.
"""

from __future__ import annotations

import json
import math
import os
import tempfile
from contextlib import contextmanager, nullcontext
from dataclasses import dataclass, field
from datetime import datetime, timezone
from pathlib import Path
from typing import Iterator

from rapidpy_common.signal_analysis import (
    SignalAnalysisError,
    fit_pass_through,
    fit_rotation,
    interpolate_motion,
    noise_summary,
)
from rapidpy_common.squid_stream import (
    BackgroundSquidCapture,
    SquidReadTiming,
    SquidStreamError,
    SquidStreamSampler,
    SquidTrace,
)

#: Segment names recorded inside :meth:`BracketedAcquisitionService.acquire`.
SEGMENT_DESCENT = "descent"
SEGMENT_TURN = "turn"
SEGMENT_ASCENT = "ascent"


class MotionCaptureError(RuntimeError):
    """Raised when a capture cannot release the SQUID link safely."""


@dataclass(frozen=True)
class MotionCaptureConfig:
    timing: SquidReadTiming = field(default_factory=SquidReadTiming)
    capture_descent: bool = True
    capture_turns: bool = True
    capture_ascent: bool = True
    output_dir: str = ""

    def wants(self, segment: str) -> bool:
        return {
            SEGMENT_DESCENT: self.capture_descent,
            SEGMENT_TURN: self.capture_turns,
            SEGMENT_ASCENT: self.capture_ascent,
        }.get(segment, False)


def stream_timing_from_squid_config(squid_config) -> SquidReadTiming:
    """Map the persisted SQUID settings onto a validated read plan."""

    axes = tuple(ch for ch in str(getattr(squid_config, "stream_axes", "XYZ")).upper() if ch in "XYZ") or ("X", "Y", "Z")
    return SquidReadTiming(
        interval_s=max(0.0, float(getattr(squid_config, "stream_interval_ms", 500))) / 1000.0,
        axes=axes,
        counts_every=int(getattr(squid_config, "stream_counts_every", 1)),
        latch_count_hold_s=float(getattr(squid_config, "stream_latch_count_hold_ms", 100)) / 1000.0,
        latch_data_hold_s=float(getattr(squid_config, "stream_latch_data_hold_ms", 120)) / 1000.0,
    )


def motion_capture_config_from_app_config(config) -> MotionCaptureConfig | None:
    """Return the capture plan, or ``None`` when capture is switched off."""

    squid = config.squid
    if not bool(getattr(squid, "motion_capture_enabled", False)):
        return None
    output_dir = str(getattr(squid, "motion_capture_dir", "") or "").strip()
    if not output_dir:
        data_dir = str(getattr(config.general, "data_dir", "") or "").strip()
        base = Path(data_dir) if data_dir else Path.home() / ".rapidpy"
        output_dir = str(base / "squid_motion_capture")
    return MotionCaptureConfig(
        timing=stream_timing_from_squid_config(squid),
        capture_descent=bool(getattr(squid, "motion_capture_descent", True)),
        capture_turns=bool(getattr(squid, "motion_capture_turns", True)),
        capture_ascent=bool(getattr(squid, "motion_capture_ascent", True)),
        output_dir=output_dir,
    )


class _Segment:
    """Handle yielded to the motion wrapper so it can attach the outcome."""

    def __init__(self, capture: BackgroundSquidCapture | None) -> None:
        self.capture = capture

    def set_outcome(self, actual: object, ok: bool) -> None:
        if self.capture is not None:
            self.capture.trace.motion["actual"] = _jsonable(actual)
            self.capture.trace.motion["ok"] = bool(ok)


class MotionCaptureRecorder:
    """Owns per-block motion captures for one bracketed acquisition service."""

    def __init__(
        self,
        client: object,
        config: MotionCaptureConfig,
        *,
        range_value: float = 1.0,
        join_timeout_s: float | None = None,
    ) -> None:
        self._client = client
        self._config = config
        self._range_value = float(range_value)
        self._join_timeout_s = join_timeout_s
        self._block_id = ""
        self._traces: list[SquidTrace] = []
        self._analyses: list[dict] = []
        self.last_written: Path | None = None

    @property
    def config(self) -> MotionCaptureConfig:
        return self._config

    def begin_block(self, block_id: str) -> None:
        self._block_id = str(block_id)
        self._traces = []
        self._analyses = []

    @contextmanager
    def segment(self, segment: str, name: str, *, target: object, mover: object | None = None) -> Iterator[_Segment]:
        if not self._config.wants(segment) or not self._block_id:
            yield _Segment(None)
            return
        sequence = len(self._traces) + 1
        sampler = SquidStreamSampler(self._client, self._config.timing, range_value=self._range_value)
        capture = BackgroundSquidCapture(
            sampler,
            f"{sequence:02d}-{segment}-{name}",
            motion={"segment": segment, "name": name, "target": _jsonable(target), "block_id": self._block_id},
            join_timeout_s=self._join_timeout_s,
        )
        observe = getattr(mover, "observe_positions", None)
        observing = observe(capture.mark_position) if callable(observe) else nullcontext()
        capture.start()
        failed = False
        try:
            with observing:
                yield _Segment(capture)
        except BaseException:
            failed = True
            raise
        finally:
            try:
                trace = capture.stop("motion failed" if failed else "motion complete")
            except SquidStreamError as exc:
                raise MotionCaptureError(str(exc)) from exc
            self._traces.append(trace)

    def end_block(self, *, completed: bool) -> tuple[SquidTrace, ...]:
        traces = tuple(self._traces)
        if not self._block_id:
            return traces
        self._analyses = [analyze_motion_trace(trace) for trace in traces]
        if self._config.output_dir and traces:
            try:
                self.last_written = write_block_captures(
                    Path(self._config.output_dir), self._block_id, traces, self._analyses, completed=completed
                )
            except OSError:
                # Diagnostic sidecars must never fail a measurement block.
                self.last_written = None
        self._block_id = ""
        return traces

    @property
    def analyses(self) -> tuple[dict, ...]:
        return tuple(self._analyses)


def _jsonable(value: object) -> object:
    if isinstance(value, float) and not math.isfinite(value):
        return None
    if isinstance(value, (int, float, str, bool)) or value is None:
        return value
    return str(value)


def analyze_motion_trace(trace: SquidTrace) -> dict:
    """Spectral summary per axis plus the physical fit for the segment type."""

    result: dict = {"label": trace.label, "summary": trace.summary(), "axes": {}, "fit": None}
    times = trace.times()
    for axis in trace.timing.axes:
        values = trace.axis_values(axis)
        try:
            summary = noise_summary(times, values)
        except SignalAnalysisError as exc:
            result["axes"][axis] = {"error": str(exc)}
            continue
        result["axes"][axis] = {
            "sample_rate_hz": summary.sample_rate_hz,
            "std": summary.std,
            "detrended_std": summary.detrended_std,
            "drift_per_s": summary.drift_per_s,
            "noise_density": summary.noise_density,
            "lines": [line.__dict__ for line in summary.lines],
        }

    segment = trace.motion.get("segment")
    try:
        path = _motion_path(trace)
        if segment == SEGMENT_TURN and {"X", "Y"} <= set(trace.timing.axes):
            fit = fit_rotation(
                path,
                trace.axis_values("X"),
                trace.axis_values("Y"),
                trace.axis_values("Z") if "Z" in trace.timing.axes else None,
            )
            result["fit"] = {
                "model": "rotation",
                "angle_span_deg": fit.angle_span_deg,
                "horizontal_amplitude": fit.horizontal_amplitude,
                "horizontal_sigma": fit.horizontal_sigma,
                "horizontal_phase_deg": fit.horizontal_phase_deg,
                "rotation_sense": fit.rotation_sense,
                "xy_consistency": fit.xy_consistency,
                "axes": {axis: harmonic.__dict__ for axis, harmonic in fit.axes.items()},
            }
        elif segment in (SEGMENT_DESCENT, SEGMENT_ASCENT):
            fits = {}
            for axis in trace.timing.axes:
                try:
                    fits[axis] = fit_pass_through(path, trace.axis_values(axis)).__dict__
                except SignalAnalysisError as exc:
                    fits[axis] = {"error": str(exc)}
            result["fit"] = {"model": "pass_through", "axes": fits}
    except SignalAnalysisError as exc:
        result["fit"] = {"error": str(exc)}
    return result


def _motion_path(trace: SquidTrace):
    """Lift height (motor counts) or turn angle (degrees) for each sample."""

    if len(trace.positions) >= 2:
        return interpolate_motion(
            trace.times(),
            [p.t_s for p in trace.positions],
            [p.position for p in trace.positions],
        )
    raise SignalAnalysisError("no encoder positions were observed during this motion")


def _atomic_write(path: Path, text: str) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    fd, tmp = tempfile.mkstemp(prefix=f".{path.name}.", dir=str(path.parent))
    try:
        with os.fdopen(fd, "w", encoding="utf-8", newline="") as handle:
            handle.write(text)
        os.replace(tmp, path)
    except BaseException:
        try:
            os.unlink(tmp)
        except OSError:
            pass
        raise


def write_block_captures(
    root: Path,
    block_id: str,
    traces: tuple[SquidTrace, ...],
    analyses: list[dict],
    *,
    completed: bool,
) -> Path:
    """Write one folder per block: a CSV per motion plus ``analysis.json``."""

    stamp = datetime.now(timezone.utc).strftime("%Y%m%dT%H%M%SZ")
    safe_block = "".join(ch if ch.isalnum() or ch in "-_" else "_" for ch in block_id)
    folder = root / f"{stamp}_{safe_block}"
    for trace in traces:
        safe_label = "".join(ch if ch.isalnum() or ch in "-_" else "_" for ch in trace.label)
        _atomic_write(folder / f"{safe_label}.csv", trace.to_csv())
    payload = {
        "block_id": block_id,
        "completed": bool(completed),
        "written_iso": datetime.now(timezone.utc).isoformat(),
        "note": "Auxiliary motion capture; not part of the bracketed measurement or published results.",
        "segments": analyses,
    }
    _atomic_write(folder / "analysis.json", json.dumps(_sanitize(payload), indent=2, default=_jsonable))
    return folder


def _sanitize(value: object) -> object:
    if isinstance(value, dict):
        return {k: _sanitize(v) for k, v in value.items()}
    if isinstance(value, (list, tuple)):
        return [_sanitize(v) for v in value]
    if isinstance(value, float) and not math.isfinite(value):
        return None
    return value
