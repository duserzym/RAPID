"""Live SQUID stream: operator-controlled read rate, filtering and spectrum.

Opened from the SQUID communication dialog, so it runs inside that dialog's
exclusive ``squid`` ownership lease.  The stream runs on a worker thread
(:class:`rapidpy_common.squid_stream.BackgroundSquidCapture`) and the UI only
redraws from snapshots on a timer, so a slow serial link never blocks the
window.  The same view opens saved motion-capture CSV files and shows the
descent / rotation fits for them.
"""

from __future__ import annotations

import math
import random
import time
from pathlib import Path
from typing import Callable

from PySide6 import QtCore, QtGui, QtWidgets

from rapid_main.glass_theme import set_semantic_status
from rapidpy_common.glass import fit_workspace_window
from rapidpy_common.signal_analysis import (
    SignalAnalysisError,
    detrend,
    fit_pass_through,
    fit_rotation,
    interpolate_motion,
    lowpass,
    noise_summary,
    notch,
    resample_uniform,
    welch_psd,
)
from rapidpy_common.squid_stream import (
    AXES,
    BackgroundSquidCapture,
    SquidReadTiming,
    SquidStreamError,
    SquidStreamSampler,
    SquidTrace,
    effective_period_s,
    estimate_sample_period_s,
)

try:
    import numpy as np
    import pyqtgraph as pg
except ImportError:  # pragma: no cover - packaged app ships both
    np = None  # type: ignore[assignment]
    pg = None  # type: ignore[assignment]

AXIS_COLORS = {"X": "#7A0219", "Y": "#0F766E", "Z": "#B7791F"}
AXIS_CHOICES = (("X, Y and Z", "XYZ"), ("Z only (fastest)", "Z"), ("X and Y", "XY"))
MAX_LIVE_SAMPLES = 200_000


class _SimulatedStreamClient:
    """NO-COMM practice signal: drift, white noise and a 0.8 Hz line.  Always labelled."""

    def __init__(self) -> None:
        self._t0 = time.perf_counter()
        self._rng = random.Random(4)

    def latch(self, axis="A", *, settle_s=0.0, count_hold_s=None, data_hold_s=None):
        time.sleep((count_hold_s or 0.0) + (data_hold_s or 0.0))
        return ("ALC", "ALD")

    def _value(self, axis: str) -> float:
        t = time.perf_counter() - self._t0
        base = {"X": 0.40, "Y": -0.25, "Z": 1.10}[axis]
        return base + 0.002 * t + 0.03 * math.sin(2 * math.pi * 0.8 * t) + self._rng.gauss(0.0, 0.01)

    def read_axis(self, axis, *, range_value=1.0):
        time.sleep(0.004)
        return type("Reading", (), {"counts": 0.0, "dvm": -self._value(axis)})()

    def read_axis_data(self, axis):
        time.sleep(0.002)
        return -self._value(axis)


class _SampleBridge(QtCore.QObject):
    finished = QtCore.Signal()


class SquidStreamDialog(QtWidgets.QDialog):
    def __init__(
        self,
        parent: QtWidgets.QWidget | None = None,
        *,
        client_provider: Callable[[], object] | None = None,
        squid_config: object | None = None,
        simulated: bool = False,
    ) -> None:
        super().__init__(parent)
        self.setObjectName("glassDialog")
        self.setWindowTitle("SQUID Live Stream & Spectrum")
        self.setAccessibleName("SQUID live stream and spectrum")
        self.setWindowFlags(self.windowFlags() & ~QtCore.Qt.WindowContextHelpButtonHint)
        self._preferred_size = (1180, 760)
        self.setMinimumSize(860, 560)
        self.resize(*self._preferred_size)
        self._client_provider = client_provider
        self._cfg = squid_config
        self._simulated = bool(simulated)
        self._capture: BackgroundSquidCapture | None = None
        self._trace: SquidTrace | None = None
        self._bridge = _SampleBridge(self)
        self._bridge.finished.connect(self._on_stream_finished)
        self._refresh = QtCore.QTimer(self)
        self._refresh.setInterval(250)
        self._refresh.timeout.connect(self._redraw)
        self._build_ui()
        self._load_config()
        self._update_estimate()
        self._set_status("Idle — choose a read rate and press Start." + (" (simulated)" if self._simulated else ""), "neutral")

    def _fit_to_screen(self, screen: QtGui.QScreen | None = None) -> None:
        """Bounds-guard hook: plots need more than the compact default width."""
        self._preferred_size = fit_workspace_window(self, self._preferred_size, screen=screen)

    # ── layout ───────────────────────────────────────────────────────────
    def _build_ui(self) -> None:
        root = QtWidgets.QHBoxLayout(self)
        root.setContentsMargins(16, 14, 16, 14)
        root.setSpacing(14)

        controls_scroll = QtWidgets.QScrollArea()
        controls_scroll.setObjectName("dialogScroll")
        controls_scroll.setWidgetResizable(True)
        controls_scroll.setFrameShape(QtWidgets.QFrame.Shape.NoFrame)
        controls_scroll.setHorizontalScrollBarPolicy(QtCore.Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        controls_scroll.setFixedWidth(330)
        host = QtWidgets.QWidget()
        controls = QtWidgets.QVBoxLayout(host)
        controls.setContentsMargins(0, 0, 6, 0)
        controls.setSpacing(10)

        title = QtWidgets.QLabel("SQUID Live Stream")
        title.setObjectName("dialogTitle")
        controls.addWidget(title)

        rate = QtWidgets.QGroupBox("Read rate")
        form = QtWidgets.QFormLayout(rate)
        form.setLabelAlignment(QtCore.Qt.AlignRight)
        self._interval = QtWidgets.QSpinBox()
        self._interval.setRange(0, 60_000)
        self._interval.setSingleStep(50)
        self._interval.setSuffix(" ms")
        self._interval.setSpecialValueText("As fast as link allows")
        self._interval.setAccessibleName("Requested interval between SQUID samples")
        form.addRow("Sample every:", self._interval)
        self._axes = QtWidgets.QComboBox()
        for label, value in AXIS_CHOICES:
            self._axes.addItem(label, value)
        self._axes.setAccessibleName("Axes to read")
        form.addRow("Axes:", self._axes)
        self._counts_every = QtWidgets.QSpinBox()
        self._counts_every.setRange(1, 100)
        self._counts_every.setPrefix("every ")
        self._counts_every.setSuffix(" samples")
        self._counts_every.setToolTip(
            "Re-read the flux counter on every Nth sample and reuse it in between.\n"
            "1 reads the counter every time (safest). Larger values roughly halve the\n"
            "per-axis cost but cannot see a flux jump until the next counter read."
        )
        form.addRow("Counter read:", self._counts_every)
        self._count_hold = QtWidgets.QSpinBox()
        self._count_hold.setRange(0, 1000)
        self._count_hold.setSuffix(" ms")
        self._count_hold.setToolTip("Pause after LatchCount. VB6 uses 100 ms.")
        form.addRow("Latch count hold:", self._count_hold)
        self._data_hold = QtWidgets.QSpinBox()
        self._data_hold.setRange(0, 1000)
        self._data_hold.setSuffix(" ms")
        self._data_hold.setToolTip("Pause after LatchData. VB6 uses 120 ms.")
        form.addRow("Latch data hold:", self._data_hold)
        self._estimate = QtWidgets.QLabel()
        self._estimate.setObjectName("guidanceText")
        self._estimate.setWordWrap(True)
        form.addRow(self._estimate)
        controls.addWidget(rate)
        for widget in (self._interval, self._counts_every, self._count_hold, self._data_hold):
            widget.valueChanged.connect(self._update_estimate)
        self._axes.currentIndexChanged.connect(self._update_estimate)

        run_row = QtWidgets.QHBoxLayout()
        self._start = QtWidgets.QPushButton("Start")
        self._start.setObjectName("accent")
        self._start.clicked.connect(self._start_stream)
        self._stop = QtWidgets.QPushButton("Stop")
        self._stop.setEnabled(False)
        self._stop.clicked.connect(self._stop_stream)
        run_row.addWidget(self._start)
        run_row.addWidget(self._stop)
        controls.addLayout(run_row)

        filters = QtWidgets.QGroupBox("Display filter")
        filter_form = QtWidgets.QFormLayout(filters)
        filter_form.setLabelAlignment(QtCore.Qt.AlignRight)
        self._lowpass = QtWidgets.QDoubleSpinBox()
        self._lowpass.setRange(0.0, 100.0)
        self._lowpass.setDecimals(2)
        self._lowpass.setSingleStep(0.1)
        self._lowpass.setSuffix(" Hz")
        self._lowpass.setSpecialValueText("Off")
        self._lowpass.setToolTip("Zero-phase low-pass applied to the overlay and analysis.")
        filter_form.addRow("Low-pass:", self._lowpass)
        self._notch = QtWidgets.QDoubleSpinBox()
        self._notch.setRange(0.0, 100.0)
        self._notch.setDecimals(2)
        self._notch.setSuffix(" Hz")
        self._notch.setSpecialValueText("Off")
        self._notch.setToolTip("Remove one interference line (pick it from the spectrum).")
        filter_form.addRow("Notch:", self._notch)
        controls.addWidget(filters)
        for widget in (self._lowpass, self._notch):
            widget.valueChanged.connect(self._redraw)

        file_row = QtWidgets.QHBoxLayout()
        self._save = QtWidgets.QPushButton("Save CSV…")
        self._save.clicked.connect(self._save_trace)
        self._save.setEnabled(False)
        self._open = QtWidgets.QPushButton("Open capture…")
        self._open.setToolTip("Open a live or motion-capture CSV and show its spectrum and fits.")
        self._open.clicked.connect(self._open_trace)
        file_row.addWidget(self._save)
        file_row.addWidget(self._open)
        controls.addLayout(file_row)

        self._apply_defaults = QtWidgets.QPushButton("Use these rate settings by default")
        self._apply_defaults.setToolTip("Store the read-rate settings for live streams and motion capture.")
        self._apply_defaults.clicked.connect(self._store_config)
        self._apply_defaults.setEnabled(self._cfg is not None)
        controls.addWidget(self._apply_defaults)

        self._status = QtWidgets.QLabel()
        self._status.setWordWrap(True)
        controls.addWidget(self._status)
        controls.addStretch(1)
        controls_scroll.setWidget(host)
        root.addWidget(controls_scroll)

        right = QtWidgets.QVBoxLayout()
        right.setSpacing(10)
        if pg is None or np is None:
            missing = QtWidgets.QLabel("Live plotting requires numpy and pyqtgraph. Install the app requirements.")
            missing.setWordWrap(True)
            right.addWidget(missing, 1)
            self._graph = None
        else:
            self._graph = pg.GraphicsLayoutWidget()
            self._graph.setBackground((255, 255, 255, 170))
            self._time_plot = self._graph.addPlot(row=0, col=0)
            self._time_plot.setLabel("bottom", "Time", units="s")
            self._time_plot.setLabel("left", "Raw SQUID (counts + DVM)")
            self._time_plot.showGrid(x=True, y=True, alpha=0.18)
            self._time_plot.addLegend(offset=(8, 8))
            self._psd_plot = self._graph.addPlot(row=1, col=0)
            self._psd_plot.setLabel("bottom", "Frequency", units="Hz")
            self._psd_plot.setLabel("left", "PSD (units²/Hz)")
            self._psd_plot.setLogMode(x=False, y=True)
            self._psd_plot.getAxis("left").enableAutoSIPrefix(False)
            self._psd_plot.showGrid(x=True, y=True, alpha=0.18)
            self._raw_curves = {}
            self._filtered_curves = {}
            self._psd_curves = {}
            for axis in AXES:
                color = AXIS_COLORS[axis]
                self._raw_curves[axis] = self._time_plot.plot(pen=pg.mkPen(color, width=1), name=f"{axis} raw")
                self._filtered_curves[axis] = self._time_plot.plot(pen=pg.mkPen(color, width=2.5, style=QtCore.Qt.PenStyle.DashLine))
                self._psd_curves[axis] = self._psd_plot.plot(pen=pg.mkPen(color, width=1.6), name=axis)
            right.addWidget(self._graph, 3)
        self._analysis = QtWidgets.QPlainTextEdit()
        self._analysis.setReadOnly(True)
        self._analysis.setAccessibleName("Stream analysis")
        self._analysis.setMinimumHeight(140)
        right.addWidget(self._analysis, 1)
        root.addLayout(right, 1)

    # ── configuration ────────────────────────────────────────────────────
    def _load_config(self) -> None:
        cfg = self._cfg
        self._interval.setValue(int(getattr(cfg, "stream_interval_ms", 500)))
        axes = str(getattr(cfg, "stream_axes", "XYZ")).upper()
        index = next((i for i, (_, value) in enumerate(AXIS_CHOICES) if value == axes), 0)
        self._axes.setCurrentIndex(index)
        self._counts_every.setValue(int(getattr(cfg, "stream_counts_every", 1)))
        self._count_hold.setValue(int(getattr(cfg, "stream_latch_count_hold_ms", 100)))
        self._data_hold.setValue(int(getattr(cfg, "stream_latch_data_hold_ms", 120)))

    def _store_config(self) -> None:
        cfg = self._cfg
        if cfg is None:
            return
        cfg.stream_interval_ms = int(self._interval.value())
        cfg.stream_axes = str(self._axes.currentData())
        cfg.stream_counts_every = int(self._counts_every.value())
        cfg.stream_latch_count_hold_ms = int(self._count_hold.value())
        cfg.stream_latch_data_hold_ms = int(self._data_hold.value())
        self._set_status("Read-rate settings stored; save the SQUID settings to keep them.", "ready")

    def timing(self) -> SquidReadTiming:
        return SquidReadTiming(
            interval_s=self._interval.value() / 1000.0,
            axes=tuple(str(self._axes.currentData())),
            counts_every=int(self._counts_every.value()),
            latch_count_hold_s=self._count_hold.value() / 1000.0,
            latch_data_hold_s=self._data_hold.value() / 1000.0,
        )

    def _baud(self) -> int:
        try:
            return int(getattr(self._cfg, "baud", 1200) or 1200)
        except (TypeError, ValueError):
            return 1200

    def _update_estimate(self) -> None:
        try:
            timing = self.timing()
        except SquidStreamError as exc:
            self._estimate.setText(str(exc))
            return
        baud = self._baud()
        floor = estimate_sample_period_s(timing, baud)
        period = effective_period_s(timing, baud)
        limited = timing.interval_s < floor
        self._estimate.setText(
            f"At {baud} baud one sample needs ≈ {floor * 1000:.0f} ms, so the stream will run at "
            f"≈ {1.0 / period:.2f} Hz{' (link-limited)' if limited else ''}. "
            "The achieved rate is measured and shown while streaming."
        )

    # ── streaming ────────────────────────────────────────────────────────
    def _resolve_client(self) -> object:
        if self._simulated:
            return _SimulatedStreamClient()
        if self._client_provider is None:
            raise SquidStreamError("No SQUID client is available.")
        client = self._client_provider()
        if client is None:
            raise SquidStreamError("The SQUID backend did not provide a raw 2G client.")
        return client

    def _start_stream(self) -> None:
        if self._capture is not None:
            return
        try:
            timing = self.timing()
            client = self._resolve_client()
        except Exception as exc:
            self._set_status(f"Cannot start: {exc}", "error")
            return
        range_value = 1.0
        sampler = SquidStreamSampler(client, timing, range_value=range_value)
        label = "live-simulated" if self._simulated else "live"
        self._capture = BackgroundSquidCapture(
            sampler,
            label,
            max_samples=MAX_LIVE_SAMPLES,
            on_sample=None,
        )
        self._trace = self._capture.trace
        self._capture.start()
        self._watch_thread()
        self._start.setEnabled(False)
        self._stop.setEnabled(True)
        self._save.setEnabled(False)
        self._refresh.start()
        self._set_status("Streaming…" + (" (simulated)" if self._simulated else ""), "active")

    def _watch_thread(self) -> None:
        capture = self._capture

        def poll() -> None:
            if capture is not self._capture:
                return
            if capture is not None and not capture.running:
                self._bridge.finished.emit()
            else:
                QtCore.QTimer.singleShot(200, poll)

        QtCore.QTimer.singleShot(200, poll)

    def _stop_stream(self) -> None:
        capture = self._capture
        if capture is None:
            return
        try:
            capture.stop("operator stop")
        except SquidStreamError as exc:
            self._set_status(str(exc), "error")
            return
        self._finish(capture)

    def _on_stream_finished(self) -> None:
        capture = self._capture
        if capture is None:
            return
        try:
            capture.stop(capture.trace.stop_reason or "finished")
        except SquidStreamError as exc:
            self._set_status(str(exc), "error")
            return
        self._finish(capture)

    def _finish(self, capture: BackgroundSquidCapture) -> None:
        self._capture = None
        self._refresh.stop()
        self._start.setEnabled(True)
        self._stop.setEnabled(False)
        self._save.setEnabled(bool(capture.trace.samples))
        self._redraw()
        trace = capture.trace
        if trace.errors:
            self._set_status(f"Stream ended: {trace.errors[-1]}", "error")
        else:
            self._set_status(
                f"Stopped ({trace.stop_reason}). {len(trace.samples)} samples at {trace.achieved_rate_hz:.2f} Hz.", "ready"
            )

    # ── files ────────────────────────────────────────────────────────────
    def _save_trace(self) -> None:
        if self._trace is None or not self._trace.samples:
            return
        default = Path.home() / f"squid_stream_{time.strftime('%Y%m%d_%H%M%S')}.csv"
        path, _ = QtWidgets.QFileDialog.getSaveFileName(self, "Save SQUID stream", str(default), "CSV (*.csv)")
        if not path:
            return
        try:
            Path(path).write_text(self._trace.to_csv(), encoding="utf-8")
        except OSError as exc:
            self._set_status(f"Save failed: {exc}", "error")
            return
        self._set_status(f"Saved {Path(path).name}", "ready")

    def _open_trace(self) -> None:
        if self._capture is not None:
            self._set_status("Stop the live stream before opening a file.", "warning")
            return
        path, _ = QtWidgets.QFileDialog.getOpenFileName(self, "Open SQUID capture", str(Path.home()), "CSV (*.csv)")
        if path:
            self.load_trace_file(Path(path))

    def load_trace_file(self, path: Path) -> None:
        try:
            self._trace = SquidTrace.from_csv(Path(path).read_text(encoding="utf-8"))
        except (OSError, ValueError, SquidStreamError) as exc:
            self._set_status(f"Cannot open {Path(path).name}: {exc}", "error")
            return
        self._save.setEnabled(False)
        self._redraw()
        self._set_status(f"Opened {Path(path).name} ({len(self._trace.samples)} samples).", "ready")

    # ── drawing and analysis ─────────────────────────────────────────────
    def _filtered(self, t, y):
        grid, uniform, dt = resample_uniform(t, y)
        fs = 1.0 / dt
        out = uniform
        if self._notch.value() > 0 and self._notch.value() < fs / 2:
            out = notch(out, fs, self._notch.value())
        if self._lowpass.value() > 0 and self._lowpass.value() < fs / 2:
            out = lowpass(out, fs, self._lowpass.value())
        return grid, out, fs

    def _redraw(self) -> None:
        trace = self._trace
        if trace is None:
            return
        samples = list(trace.samples)  # snapshot; the worker keeps appending
        if not samples:
            return
        times = [s.t_s for s in samples]
        lines = [
            f"{trace.label}: {len(samples)} samples, achieved "
            f"{(len(samples) - 1) / (times[-1] - times[0]) if len(samples) > 1 and times[-1] > times[0] else 0.0:.2f} Hz"
        ]
        filters = []
        if self._lowpass.value() > 0:
            filters.append(f"low-pass {self._lowpass.value():g} Hz")
        if self._notch.value() > 0:
            filters.append(f"notch {self._notch.value():g} Hz")
        if filters:
            lines.append("Filter: " + ", ".join(filters))
        for axis in AXES:
            slot = AXES.index(axis)
            values = [s.raw[slot] for s in samples]
            visible = axis in trace.timing.axes
            if self._graph is not None:
                self._raw_curves[axis].setData(times if visible else [], values if visible else [])
            if not visible:
                if self._graph is not None:
                    self._filtered_curves[axis].setData([], [])
                    self._psd_curves[axis].setData([], [])
                continue
            try:
                # Statistics and spectrum describe the raw signal (that is what a
                # filter is chosen from); the filter only shapes the overlay.
                summary = noise_summary(times, values)
                grid, raw_uniform, _dt = resample_uniform(times, values)
                freqs, psd = welch_psd(raw_uniform, summary.sample_rate_hz)
                _grid, filtered, _fs = self._filtered(times, values)
            except SignalAnalysisError:
                if self._graph is not None:
                    self._filtered_curves[axis].setData([], [])
                    self._psd_curves[axis].setData([], [])
                continue
            if self._graph is not None:
                show_filter = bool(filters)
                self._filtered_curves[axis].setData(grid if show_filter else [], filtered if show_filter else [])
                positive = psd > 0
                self._psd_curves[axis].setData(freqs[positive][1:], psd[positive][1:])
            line_text = ", ".join(f"{line.frequency_hz:.3g} Hz (×{line.prominence:.0f})" for line in summary.lines[:3]) or "none"
            filtered_text = f", filtered σ={float(detrend(filtered, 1).std()):.4g}" if filters else ""
            lines.append(
                f"{axis}: σ={summary.detrended_std:.4g} (detrended){filtered_text}, drift={summary.drift_per_s:.3g}/s, "
                f"noise density={summary.noise_density:.3g}/√Hz, lines: {line_text}"
            )
        lines.extend(self._motion_fit_lines(trace))
        self._analysis.setPlainText("\n".join(lines))

    def _motion_fit_lines(self, trace: SquidTrace) -> list[str]:
        segment = trace.motion.get("segment")
        if segment is None:
            return []
        if len(trace.positions) < 2:
            return [f"{segment}: no encoder positions recorded; motion fit unavailable."]
        try:
            path = interpolate_motion(trace.times(), [p.t_s for p in trace.positions], [p.position for p in trace.positions])
            if segment == "turn" and {"X", "Y"} <= set(trace.timing.axes):
                fit = fit_rotation(path, trace.axis_values("X"), trace.axis_values("Y"))
                return [
                    f"Rotation fit over {fit.angle_span_deg:.0f}°: horizontal amplitude {fit.horizontal_amplitude:.4g} "
                    f"± {fit.horizontal_sigma:.2g}, phase {fit.horizontal_phase_deg:.1f}°, X/Y consistency {fit.xy_consistency:.3f}"
                ]
            if segment in ("descent", "ascent"):
                out = []
                for axis in trace.timing.axes:
                    fit = fit_pass_through(path, trace.axis_values(axis))
                    out.append(
                        f"Pass-through {axis}: peak {fit.amplitude:.4g} at {fit.center:.0f} counts "
                        f"(width {fit.width:.0f}, SNR {fit.snr:.1f})"
                    )
                return out
        except SignalAnalysisError as exc:
            return [f"{segment} fit unavailable: {exc}"]
        return []

    def _set_status(self, text: str, level: str) -> None:
        if self._simulated and level in {"neutral", "ready", "active"}:
            level = "simulated"
        set_semantic_status(self._status, text, level, accessible_name="SQUID stream status")

    def done(self, result: int) -> None:  # noqa: D401 - Qt override
        """Always release the serial link before the dialog closes."""
        if self._capture is not None:
            try:
                self._capture.stop("dialog closed")
            except SquidStreamError as exc:
                self._set_status(str(exc), "error")
                return
            self._capture = None
        self._refresh.stop()
        super().done(result)
