from __future__ import annotations

from collections import deque
from datetime import datetime
import json
from pathlib import Path
from typing import Optional

from PySide6 import QtCore, QtGui, QtWidgets

from rapid_main.analysis import ReadingCycleStatistics, reading_cycle_statistics
from rapid_main.calibration_registry import CalibrationRegistry, CalibrationRegistryError
from rapid_main.data_model import MeasurementStep, SpecimenMeta
from rapid_main.specimen_metadata import resolve_specimen_meta, validate_specimen_provenance
from rapid_main.device_ownership import DeviceOwnershipError
from rapid_main.dialogs.plots import build_quicklook_summary, write_quicklook_json
from rapid_main.hardware_contracts import MeasurementBackend, NoCommBackend
from rapid_main.io.measurement_bundle import SIMULATED_SUBDIR, validate_measurement_output
from rapid_main.specimen_paths import specimen_run_directory
from rapid_main.measurement_worker import MeasurementWorker, StepResult
from rapid_main.printing import print_widget_snapshot

try:
    import pyqtgraph as pg
except ImportError:  # pragma: no cover - optional dependency on operator environment
    pg = None


class MeasurementTraceWidget(QtWidgets.QWidget):
    """Dependency-free quicklook used when pyqtgraph is unavailable."""

    def __init__(self, parent: QtWidgets.QWidget | None = None) -> None:
        super().__init__(parent)
        self.setObjectName("measurementTraceFallback")
        self.setAccessibleName("Live measurement quicklook")
        self.setMinimumHeight(310)
        self._series: dict[str, list[float]] = {
            "step": [],
            "moment": [],
            "dec": [],
            "inc": [],
        }

    def set_series(
        self,
        steps: list[float],
        moments: list[float],
        declinations: list[float],
        inclinations: list[float],
    ) -> None:
        self._series = {
            "step": list(steps),
            "moment": list(moments),
            "dec": list(declinations),
            "inc": list(inclinations),
        }
        self.update()

    def clear(self) -> None:
        self.set_series([], [], [], [])

    @staticmethod
    def _points(values: list[float], rect: QtCore.QRectF) -> list[QtCore.QPointF]:
        if not values:
            return []
        low = min(values)
        high = max(values)
        span = high - low
        if span <= 0.0:
            span = 1.0
            low -= 0.5
        divisor = max(1, len(values) - 1)
        return [
            QtCore.QPointF(
                rect.left() + (idx / divisor) * rect.width(),
                rect.bottom() - ((value - low) / span) * rect.height(),
            )
            for idx, value in enumerate(values)
        ]

    def paintEvent(self, event: QtGui.QPaintEvent) -> None:  # type: ignore[override]
        super().paintEvent(event)
        painter = QtGui.QPainter(self)
        painter.setRenderHint(QtGui.QPainter.RenderHint.Antialiasing, True)
        painter.setPen(QtCore.Qt.PenStyle.NoPen)
        painter.setBrush(QtGui.QColor(255, 255, 255, 150))
        painter.drawRoundedRect(QtCore.QRectF(self.rect()).adjusted(1, 1, -1, -1), 14, 14)

        bounds = QtCore.QRectF(self.rect()).adjusted(42, 28, -18, -32)
        upper = QtCore.QRectF(bounds.left(), bounds.top(), bounds.width(), bounds.height() * 0.44)
        lower = QtCore.QRectF(
            bounds.left(),
            bounds.top() + bounds.height() * 0.56,
            bounds.width(),
            bounds.height() * 0.44,
        )
        grid_pen = QtGui.QPen(QtGui.QColor(122, 2, 25, 34), 1)
        painter.setPen(grid_pen)
        for rect in (upper, lower):
            painter.drawRect(rect)
            for index in range(1, 4):
                y = rect.top() + rect.height() * index / 4
                painter.drawLine(QtCore.QPointF(rect.left(), y), QtCore.QPointF(rect.right(), y))

        painter.setPen(QtGui.QColor("#6F6265"))
        painter.drawText(QtCore.QRectF(8, upper.top(), 32, 24), QtCore.Qt.AlignmentFlag.AlignRight, "|M|")
        painter.drawText(QtCore.QRectF(8, lower.top(), 32, 24), QtCore.Qt.AlignmentFlag.AlignRight, "°")
        painter.drawText(
            QtCore.QRectF(bounds.left(), bounds.bottom() + 5, bounds.width(), 22),
            QtCore.Qt.AlignmentFlag.AlignCenter,
            "Treatment step",
        )

        if not self._series["step"]:
            painter.setPen(QtGui.QColor("#887A7D"))
            painter.drawText(self.rect(), QtCore.Qt.AlignmentFlag.AlignCenter, "Waiting for accepted measurement steps")
            painter.end()
            return

        for key, rect, color in (
            ("moment", upper, "#7A0219"),
            ("dec", lower, "#31566D"),
            ("inc", lower, "#D69F00"),
        ):
            points = self._points(self._series[key], rect)
            painter.setPen(QtGui.QPen(QtGui.QColor(color), 2.2))
            if len(points) == 1:
                painter.drawEllipse(points[0], 3.5, 3.5)
            elif points:
                painter.drawPolyline(QtGui.QPolygonF(points))
        painter.end()


class MeasurementPanel(QtWidgets.QWidget):
    """Live measurement display panel.

    Maps to: frmMeasure (Measurement Window) in VB6.

    Layout:
        Top strip  — current sample, depth, step, coordinates selector
        Col A (22%) — flow controls, coordinate frame selector
        Col B (42%) — live SQUID readings card + measurement stats card
        Col C (36%) — live moment and direction quicklook
        Bottom     — quality warning banners (hidden until data arrives)
    """

    sample_run_finished = QtCore.Signal(bool, str)
    _TRACE_LIMIT = 1200

    def __init__(self, parent: QtWidgets.QWidget | None = None) -> None:
        super().__init__(parent)
        self._worker: Optional[MeasurementWorker] = None
        self._pending_run_finished: bool | None = None
        self._worker_settle_timer = QtCore.QTimer(self)
        self._worker_settle_timer.setSingleShot(True)
        self._worker_settle_timer.timeout.connect(self._settle_measurement_worker)
        self._leases: list[object] = []
        self._current_sample = "UNKNOWN"
        self._current_depth = "—"
        self._current_treatment = "—"
        self._current_output_dir: Path | None = None
        self._current_run_simulated = False
        self._plot_traces: dict[str, deque[float]] = {
            "step": deque(maxlen=self._TRACE_LIMIT),
            "moment": deque(maxlen=self._TRACE_LIMIT),
            "dec": deque(maxlen=self._TRACE_LIMIT),
            "inc": deque(maxlen=self._TRACE_LIMIT),
            "susceptibility": deque(maxlen=self._TRACE_LIMIT),
        }
        self._completed_steps: list[MeasurementStep] = []
        self._plot_curves: dict[str, object] = {}
        self._plot_graph: object | None = None
        self._plot_fallback: MeasurementTraceWidget | None = None
        self._last_run_error = False
        root = QtWidgets.QVBoxLayout(self)
        root.setContentsMargins(0, 0, 0, 0)
        root.setSpacing(0)

        root.addWidget(self._build_sample_strip())

        body = QtWidgets.QHBoxLayout()
        body.setContentsMargins(14, 10, 14, 10)
        body.setSpacing(10)
        body.addWidget(self._build_controls_card(), 22)
        body.addWidget(self._build_readings_card(), 42)
        body.addWidget(self._build_plot_card(), 36)
        root.addLayout(body, 1)

        root.addWidget(self._build_warning_strip())

    # ── Sample info strip ─────────────────────────────────────────────────────
    def _build_sample_strip(self) -> QtWidgets.QFrame:
        strip = QtWidgets.QFrame()
        strip.setObjectName("header")
        strip.setFixedHeight(44)
        hl = QtWidgets.QHBoxLayout(strip)
        hl.setContentsMargins(18, 0, 18, 0)
        hl.setSpacing(16)

        def _pill(name: str, default: str) -> QtWidgets.QLabel:
            lbl = QtWidgets.QLabel(default)
            lbl.setObjectName("valuePill")
            setattr(self, name, lbl)
            return lbl

        for label_text, attr, default in [
            ("Sample:", "_meas_sample", "—"),
            ("Depth:",  "_meas_depth",  "—"),
            ("Step:",   "_meas_step",   "— / —"),
            ("Treatment:", "_meas_treat", "—"),
        ]:
            hl.addWidget(QtWidgets.QLabel(label_text))
            hl.addWidget(_pill(attr, default))
        hl.addStretch()
        return strip

    # ── Controls ──────────────────────────────────────────────────────────────
    def _build_controls_card(self) -> QtWidgets.QFrame:
        card = QtWidgets.QFrame()
        card.setObjectName("card")
        cl = QtWidgets.QVBoxLayout(card)
        cl.setContentsMargins(14, 14, 14, 14)
        cl.setSpacing(8)

        hdr = QtWidgets.QLabel("CONTROLS")
        hdr.setObjectName("sectionHdr")
        cl.addWidget(hdr)

        self._start_btn = QtWidgets.QPushButton("▶  Start / Resume")
        self._start_btn.setObjectName("accent")
        self._pause_btn = QtWidgets.QPushButton("⏸  Pause Run")
        self._halt_btn  = QtWidgets.QPushButton("■  Halt Run")
        self._halt_btn.setStyleSheet(
            "QPushButton { color: #b91c1c; }"
            "QPushButton:hover { background: rgba(220,38,38,0.10); }"
        )
        self._print_btn = QtWidgets.QPushButton("🖨  Print")
        self._print_btn.clicked.connect(self._print_current_view)
        self._plots_btn = QtWidgets.QPushButton("📈  Show Plots")
        self._plots_btn.clicked.connect(self._show_plots)
        self._start_btn.clicked.connect(self._on_start)
        self._pause_btn.clicked.connect(self._on_pause)
        self._halt_btn.clicked.connect(self._on_halt)

        for btn in (self._start_btn, self._pause_btn, self._halt_btn,
                    self._print_btn, self._plots_btn):
            cl.addWidget(btn)

        cl.addSpacing(10)

        # Coordinate frame
        coord_hdr = QtWidgets.QLabel("COORDINATES")
        coord_hdr.setObjectName("sectionHdr")
        cl.addWidget(coord_hdr)

        self._coord_core = QtWidgets.QRadioButton("Core")
        self._coord_geo  = QtWidgets.QRadioButton("Geographic")
        self._coord_bed  = QtWidgets.QRadioButton("Bedding")
        self._coord_core.setChecked(True)
        for rb in (self._coord_core, self._coord_geo, self._coord_bed):
            cl.addWidget(rb)

        cl.addSpacing(10)
        disp_hdr = QtWidgets.QLabel("DISPLAY")
        disp_hdr.setObjectName("sectionHdr")
        cl.addWidget(disp_hdr)
        self._chk_susc   = QtWidgets.QCheckBox("Susceptibility")
        self._chk_moment = QtWidgets.QCheckBox("Moment magnitude")
        self._chk_moment.setChecked(True)
        cl.addWidget(self._chk_susc)
        cl.addWidget(self._chk_moment)

        cl.addStretch()

        hide_btn = QtWidgets.QPushButton("↩  Back to Dashboard")
        hide_btn.clicked.connect(lambda: self._goto("dashboard"))
        cl.addWidget(hide_btn)
        return card

    # ── Live readings + stats ─────────────────────────────────────────────────
    def _build_readings_card(self) -> QtWidgets.QFrame:
        outer = QtWidgets.QFrame()
        outer.setObjectName("card")
        ov = QtWidgets.QVBoxLayout(outer)
        ov.setContentsMargins(14, 14, 14, 14)
        ov.setSpacing(10)

        # ── Raw SQUID ──
        squid_hdr = QtWidgets.QLabel("RAW SQUID  (A/m × 10⁻⁷)")
        squid_hdr.setObjectName("sectionHdr")
        ov.addWidget(squid_hdr)

        squid_grid = QtWidgets.QGridLayout()
        squid_grid.setSpacing(6)
        for col, axis in enumerate(("X", "Y", "Z")):
            lbl = QtWidgets.QLabel(axis)
            lbl.setObjectName("readLbl")
            lbl.setAlignment(QtCore.Qt.AlignCenter)
            val = QtWidgets.QLabel("—")
            val.setObjectName("valueMonospace")
            val.setAlignment(QtCore.Qt.AlignCenter)
            val.setMinimumWidth(90)
            squid_grid.addWidget(lbl, 0, col)
            squid_grid.addWidget(val, 1, col)
            setattr(self, f"_squid_{axis.lower()}", val)
        ov.addLayout(squid_grid)

        ov.addWidget(_hline())

        # ── Calculated ──
        calc_hdr = QtWidgets.QLabel("CALCULATED VALUES")
        calc_hdr.setObjectName("sectionHdr")
        ov.addWidget(calc_hdr)

        calc_grid = QtWidgets.QGridLayout()
        calc_grid.setSpacing(6)
        calc_fields = [
            ("Dec (°)",       "_calc_dec"),
            ("Inc (°)",       "_calc_inc"),
            ("Moment (A·m²)", "_calc_moment"),
            ("CSD",           "_calc_csd"),
        ]
        for i, (label, attr) in enumerate(calc_fields):
            r, c = divmod(i, 2)
            lbl = QtWidgets.QLabel(label)
            lbl.setObjectName("readLbl")
            val = QtWidgets.QLabel("—")
            val.setObjectName("valueMonospace")
            val.setAlignment(QtCore.Qt.AlignCenter)
            calc_grid.addWidget(lbl, r * 2,     c)
            calc_grid.addWidget(val, r * 2 + 1, c)
            setattr(self, attr, val)
        ov.addLayout(calc_grid)

        ov.addWidget(_hline())

        # ── Stats ──
        stats_hdr = QtWidgets.QLabel("MEASUREMENT STATS")
        stats_hdr.setObjectName("sectionHdr")
        ov.addWidget(stats_hdr)

        stats_grid = QtWidgets.QGridLayout()
        stats_grid.setSpacing(4)
        stats_grid.setColumnMinimumWidth(1, 80)
        stats_grid.setColumnMinimumWidth(3, 80)

        def _stat(r: int, c: int, label: str, attr: str) -> None:
            lbl = QtWidgets.QLabel(label)
            lbl.setObjectName("readLbl")
            val = QtWidgets.QLabel("—")
            val.setObjectName("valuePill")
            val.setAlignment(QtCore.Qt.AlignCenter)
            stats_grid.addWidget(lbl, r, c)
            stats_grid.addWidget(val, r, c + 1)
            setattr(self, attr, val)

        # 4-Position deltas
        for col, axis in enumerate(("X", "Y", "Z")):
            stats_grid.addWidget(_axis_hdr(f"Δ{axis} (emu)"), 0, col * 2, 1, 2)
        for col, axis in enumerate(("X", "Y", "Z")):
            _stat(1, col * 2, "", f"_delta_{axis.lower()}")
        for col, axis in enumerate(("X", "Y", "Z")):
            stats_grid.addWidget(_axis_hdr(f"Δ{axis}/M"), 2, col * 2, 1, 2)
        for col, axis in enumerate(("X", "Y", "Z")):
            _stat(3, col * 2, "", f"_ratio_{axis.lower()}")

        stats_grid.addWidget(_hline_w(), 4, 0, 1, 6)

        _stat(5, 0, "Avg Moment", "_avg_moment")
        _stat(5, 2, "Avg Dec",    "_avg_dec")
        _stat(5, 4, "Avg Inc",    "_avg_inc")
        _stat(6, 0, "CSD",        "_avg_csd")
        _stat(7, 0, "Sig/Drift",  "_sig_drift")
        _stat(7, 2, "Sig/Holder", "_sig_holder")
        _stat(7, 4, "Sig/Induced","_sig_induced")
        _stat(8, 0, "Holder",     "_holder_id")
        _stat(8, 2, "Holder |M|", "_holder_moment")
        _stat(8, 4, "Holder age", "_holder_age")

        ov.addLayout(stats_grid)
        ov.addStretch()
        return outer

    # ── Live plot ─────────────────────────────────────────────────────────────
    def _build_plot_card(self) -> QtWidgets.QFrame:
        card = QtWidgets.QFrame()
        card.setObjectName("card")
        cl = QtWidgets.QVBoxLayout(card)
        cl.setContentsMargins(14, 14, 14, 14)
        cl.setSpacing(8)

        hdr = QtWidgets.QLabel("MOMENT vs. TREATMENT STEP")
        hdr.setObjectName("sectionHdr")
        cl.addWidget(hdr)

        if pg is None:
            self._plot_fallback = MeasurementTraceWidget()
            self._plot_fallback.setToolTip(
                "Built-in live quicklook. Install pyqtgraph for interactive pan and zoom."
            )
            cl.addWidget(self._plot_fallback, 1)
        else:
            self._plot_graph = pg.GraphicsLayoutWidget()
            self._plot_graph.setBackground("#fffdf8")
            self._plot_graph.setMinimumHeight(310)
            moment_plot = self._plot_graph.addPlot(row=0, col=0, title="Moment trace")
            zij_plot = self._plot_graph.addPlot(row=1, col=0, title="Live Zijderveld")
            moment_plot.showGrid(x=True, y=True, alpha=0.16)
            zij_plot.showGrid(x=True, y=True, alpha=0.16)
            moment_plot.setLabel("left", "Moment", units="A·m²")
            moment_plot.setLabel("bottom", "Step")
            zij_plot.setLabel("left", "Inc", units="°")
            zij_plot.setLabel("bottom", "Dec", units="°")
            moment_plot.addLegend(offset=(8, 8))
            zij_plot.addLegend(offset=(8, 8))
            self._plot_curves["moment"] = moment_plot.plot(
                pen=pg.mkPen("#7a0219", width=2),
                name="Moment",
            )
            self._plot_curves["susceptibility"] = moment_plot.plot(
                pen=pg.mkPen("#d69f00", width=1, style=QtCore.Qt.PenStyle.DotLine),
                name="Susceptibility",
            )
            self._plot_curves["zij"] = zij_plot.plot(
                pen=pg.mkPen("#31566d", width=2),
                symbol="+",
                symbolSize=8,
                symbolPen=pg.mkPen("#0f766e"),
                symbolBrush=pg.mkBrush("#0f766e"),
            )
            cl.addWidget(self._plot_graph, 1)

        stats_btn = QtWidgets.QPushButton("📊  Show Stats Window")
        stats_btn.clicked.connect(self._show_plots)
        cl.addWidget(stats_btn)
        return card

    # ── Quality warning strip ─────────────────────────────────────────────────
    def _build_warning_strip(self) -> QtWidgets.QWidget:
        w = QtWidgets.QWidget()
        hl = QtWidgets.QHBoxLayout(w)
        hl.setContentsMargins(14, 4, 14, 8)
        hl.setSpacing(8)

        self._warn_orange = QtWidgets.QLabel(
            "⚠  Noise is 1–5× the moment — measurement quality may be poor"
        )
        self._warn_orange.setObjectName("warnOrange")
        self._warn_orange.hide()

        self._warn_red = QtWidgets.QLabel(
            "⛔  Noise > 5× the moment — consider re-measuring manually"
        )
        self._warn_red.setObjectName("warnRed")
        self._warn_red.hide()

        hl.addWidget(self._warn_orange)
        hl.addWidget(self._warn_red)
        hl.addStretch()
        return w

    # ── Dialog launchers ──────────────────────────────────────────────────────
    def _show_plots(self) -> None:
        from rapid_main.dialogs import PlotsDialog  # avoid circular at module load
        dialog = PlotsDialog(self)
        self._apply_completed_steps_to_plots(dialog)
        dialog.exec()

    def _print_current_view(self) -> bool:
        sample = self._current_sample or "UNKNOWN"
        return print_widget_snapshot(
            self,
            self,
            title=f"RAPID measurement — {sample}",
        )

    def _apply_completed_steps_to_plots(self, dialog: object) -> bool:
        if not self._completed_steps or not hasattr(dialog, "set_data"):
            return False
        dialog.set_data(
            [step.sdx for step in self._completed_steps],
            [step.sdy for step in self._completed_steps],
            [step.sdz for step in self._completed_steps],
            [step.demag_label for step in self._completed_steps],
            simulated=self._current_run_simulated,
            provenance_statement=(
                "Current run used a simulated backend; not hardware evidence."
                if self._current_run_simulated
                else "Current measurement run."
            ),
        )
        return True

    # ── Worker control ────────────────────────────────────────────────────────

    def _on_start(self) -> None:
        self.start_measurement_for_sample(self._current_sample)

    def start_measurement_for_sample(
        self,
        sample: str,
        *,
        queue_run: bool = False,
        owner: str | None = None,
    ) -> bool:
        """Start a full run for a given specimen name.

        Parameters
        ----------
        sample:
            Specimen name to measure.
        queue_run:
            True when the call originates from queue automation. Queue-driven
            calls should not cancel an active queue run.
        """
        owner = owner or ("queue_workflow" if queue_run else "measurement_panel")
        if self._worker is not None and self._worker.is_paused:
            self._worker.resume()
            return True
        if self._worker is not None and self._worker.isRunning():
            return False  # already running

        mw = self.window()
        if hasattr(mw, "cancel_queue_run") and not queue_run:
            mw.cancel_queue_run("Manual run started.")

        labels = getattr(mw, "_sequence_labels", [])
        queue_source = None
        queue_metadata = None
        queue_provenance = None
        if queue_run and hasattr(mw, 'queue_measurement_labels'):
            try:
                labels = mw.queue_measurement_labels(sample)
                if hasattr(mw, 'queue_measurement_source'):
                    queue_source = mw.queue_measurement_source(sample)
                if hasattr(mw, 'queue_measurement_metadata'):
                    queue_metadata = mw.queue_measurement_metadata(sample)
                if hasattr(mw, 'queue_measurement_provenance'):
                    queue_provenance = mw.queue_measurement_provenance(sample)
                if queue_provenance is not None:
                    if queue_metadata is None:
                        raise ValueError('Queue source provenance requires its original resolved metadata.')
                    validate_specimen_provenance(queue_provenance, queue_metadata.meta)
            except Exception as exc:
                mw.set_status('Queue sequence cannot start: ' + str(exc))
                return False
        if not labels:
            QtWidgets.QMessageBox.warning(
                self, "No Sequence",
                "No measurement sequence is loaded.\n"
                "Go to the Sequence panel and generate or load a step list first.",
            )
            return False

        cfg = getattr(mw, "config", None)
        backend = _resolve_measurement_backend(mw, NoCommBackend()) or NoCommBackend()
        self._current_run_simulated = bool(getattr(backend, "simulated", False))
        op = cfg.general.operator if cfg else ""
        out = Path(cfg.general.data_dir) if cfg and cfg.general.data_dir else Path.home() / "RAPID_data"
        try:
            resolution = queue_metadata or resolve_specimen_meta(sample,
                sample_dir=(queue_source[0].parent if queue_source else (cfg.general.sample_dir if cfg else None)),
                data_dir=(None if queue_source else (cfg.general.data_dir if cfg else None)),
                registrations=(queue_source[1] if queue_source else getattr(mw, "sample_registrations", None)))
            meta = resolution.meta
            run_output_dir = specimen_run_directory(out, meta.name, queue_source[0] if queue_source else None)
            validate_measurement_output(run_output_dir, meta.name, simulated=self._current_run_simulated)
        except (OSError, ValueError) as exc:
            QtWidgets.QMessageBox.warning(self, "Invalid Specimen Path", str(exc))
            return False
        if resolution.defaulted_fields:
            self._on_preflight_warning("Specimen metadata defaulted for: "
                + ", ".join(resolution.defaulted_fields)
                + ". Check the specimen header or sample index before archiving.")
        try:
            self._last_run_error = False
            self._clear_measurement_plot()
            self._acquire_ownerships(
                mw,
                owner=owner,
                queue_run=queue_run,
                needs_susceptibility=any(
                    str(label).strip().upper() == "SUSC" for label in labels
                ),
            )
        except DeviceOwnershipError as exc:
            QtWidgets.QMessageBox.warning(self, "Device Busy", str(exc))
            return False

        self.set_specimen_context(sample)

        self._current_output_dir = (
            run_output_dir / SIMULATED_SUBDIR
            if self._current_run_simulated
            else run_output_dir
        )
        run_id = f"{meta.name}-{datetime.now().strftime('%Y%m%dT%H%M%S')}"
        routine_context = None
        thermal_context = None
        rockmag_plan = getattr(mw, "_rockmag_routine_plan", None)
        if rockmag_plan is not None:
            routine_context = rockmag_plan.to_artifact(
                run_context=run_id,
                operator=op,
                notes="Compiled sequence executed by MeasurementWorker.",
            )
        thermal_plan = getattr(mw, "_thermal_routine_plan", None)
        if thermal_plan is not None:
            thermal_context = thermal_plan.to_artifact(
                run_context=run_id,
                operator=op,
                notes=(
                    "Manual-external thermal plan carried into MeasurementWorker. "
                    "No specimen furnace control is implied."
                ),
            )
        try:
            calibration_records = CalibrationRegistry.default().provenance_refs()
        except CalibrationRegistryError as exc:
            calibration_records = []
            self._on_preflight_warning(
                f"Calibration registry could not be verified: {exc}. "
                "The measurement bundle will contain no approved calibration record IDs."
            )
        if not calibration_records:
            self._on_preflight_warning(
                "No active, unexpired calibration records are registered. The measurement "
                "bundle will preserve this as an empty calibration-record list."
            )

        self._worker = MeasurementWorker(
            meta=meta,
            labels=labels,
            output_dir=run_output_dir,
            backend=backend,
            operator=op,
            samples_per_position=max(
                1,
                int(getattr(getattr(cfg, "squid", None), "samples_per_pos", 1)),
            ),
            run_id=run_id,
            calibration_records=calibration_records,
            communication_sources=(
                getattr(mw, "_vacuum_backend", None),
                getattr(mw, "_susceptibility_backend", None),
            ),
            routine_context=routine_context,
            thermal_context=thermal_context,
            specimen_provenance=queue_provenance,
            parent=self,
        )
        self._worker.step_started.connect(self._on_step_started)
        self._worker.step_complete.connect(self._on_step_complete)
        self._worker.run_finished.connect(self._on_run_finished)
        self._worker.finished.connect(self._settle_measurement_worker)
        self._worker.preflight_warning.connect(self._on_preflight_warning)
        self._worker.error_occurred.connect(self._on_error)
        self._worker.phase_changed.connect(self._on_phase_changed)
        self._worker.start()

        if hasattr(mw, "start_run"):
            mw.start_run(labels)
        if hasattr(mw, "set_flow_state"):
            mw.set_flow_state("running")
        return True

    def toggle_pause(self) -> None:
        """Public entrypoint for toolbar/menu pause toggles."""
        self._on_pause()

    def halt_run(self) -> None:
        """Public entrypoint for pause/stop controls from other widgets."""
        self._on_halt()

    def is_active(self) -> bool:
        """Whether a worker exists and is currently running or paused."""
        return self._worker is not None

    def _on_pause(self) -> None:
        if self._worker is not None:
            if self._worker.is_paused:
                self._worker.resume()
                self._pause_btn.setText("⏸  Pause Run")
                if hasattr(self.window(), "set_flow_state"):
                    self.window().set_flow_state("running")
            else:
                self._worker.pause()
                self._pause_btn.setText("▶  Resume Run")
                if hasattr(self.window(), "set_flow_state"):
                    self.window().set_flow_state("paused")

    def _on_halt(self) -> None:
        if self._worker is not None:
            if hasattr(self.window(), "set_flow_state"):
                self.window().set_flow_state("halted")
            self._worker.halt()

    def set_specimen_context(
        self,
        sample: str,
        depth: str = "—",
        treatment: str = "—",
    ) -> None:
        """Update live measurement panel context for the active specimen."""
        self._current_sample = sample or "UNKNOWN"
        self._current_depth = depth or "—"
        self._current_treatment = treatment or "—"
        self._meas_sample.setText(self._current_sample)
        self._meas_depth.setText(self._current_depth)
        self._meas_treat.setText(self._current_treatment)

        mw = self.window()
        if hasattr(mw, "set_current_sample"):
            mw.set_current_sample(self._current_sample)

    @QtCore.Slot(int, str)
    def _on_step_started(self, idx: int, label: str) -> None:
        mw = self.window()
        total = len(self._worker._labels) if self._worker is not None else len(getattr(mw, "_sequence_labels", []))
        self._meas_step.setText(f"{idx + 1} / {total}")
        self._meas_treat.setText(label)
        if hasattr(mw, "set_step"):
            mw.set_step(label)
        if hasattr(mw, "advance_step") and idx > 0:
            mw.advance_step()

    @QtCore.Slot(object)
    def _on_step_complete(self, result: StepResult) -> None:
        step = result.step
        self._completed_steps.append(step)
        # Update raw SQUID display (scaled to 1e-7)
        scale = 1e7
        self.update_squid(
            f"{step.sdx * scale:.4f}",
            f"{step.sdy * scale:.4f}",
            f"{step.sdz * scale:.4f}",
        )
        # Update calculated values
        moment_Am2 = step.magn_moment_Am2()
        self.update_calculated(
            dec=f"{step.gdec:.1f}",
            inc=f"{step.ginc:.1f}",
            moment=f"{moment_Am2:.3e}",
            csd=f"{step.error_angle:.1f}°",
        )
        self._update_measurement_stats(
            step,
            result.cycle_stats,
            block_result=getattr(result, "block_result", None),
            holder_status=getattr(result, "holder_status", None),
            collection_stats=getattr(result, "collection_stats", None),
        )
        self._plot_step_complete(result, moment_Am2=moment_Am2)

    def _update_measurement_stats(
        self,
        step: MeasurementStep,
        cycle_stats: ReadingCycleStatistics | None,
        block_result: object | None = None,
        holder_status: object | None = None,
        collection_stats: object | None = None,
    ) -> None:
        """Render VB6-style cycle quality values from the current SQUID reads."""

        stats = cycle_stats or reading_cycle_statistics([(step.sdx, step.sdy, step.sdz)])
        moment = float(step.moment)
        self._avg_moment.setText(f"{moment:.3e} emu")
        self._avg_dec.setText(f"{step.sdec:.1f} deg")
        self._avg_inc.setText(f"{step.sinc:.1f} deg")
        self._avg_csd.setText(
            "N/A" if stats.directional_spread_deg is None else f"{stats.directional_spread_deg:.2f} deg"
        )

        ratios: list[float] = []
        for axis, delta in zip(("x", "y", "z"), stats.axis_ranges):
            getattr(self, f"_delta_{axis}").setText(f"{delta:.3e}")
            ratio = None if moment <= 0.0 else delta / moment
            getattr(self, f"_ratio_{axis}").setText("N/A" if ratio is None else f"{ratio:.2f}")
            if ratio is not None:
                ratios.append(ratio)

        if stats.signal_to_drift is None:
            self._sig_drift.setText("N/A")
        elif stats.signal_to_drift == float("inf"):
            self._sig_drift.setText("stable")
        else:
            self._sig_drift.setText(f"{stats.signal_to_drift:.2f}")

        self._update_holder_stats(collection_stats or block_result, holder_status)
        if collection_stats is not None:
            self._avg_csd.setText(f"{collection_stats.fischer_sd_deg:.2f} deg")
            self._avg_csd.setToolTip(f"Fischer circular SD across {collection_stats.position_count} positions "
                                   f"in {collection_stats.block_count} accepted blocks.")
            self._sig_drift.setText(f"{collection_stats.sig_drift:.2f}")
            self._sig_holder.setToolTip("Collection mean moment / mean block holder magnitude.")
            self._sig_induced.setToolTip("Collection mean moment / mean rotational asymmetry magnitude.")
        else:
            self._avg_csd.setToolTip(f"RMS directional spread across {stats.count} SQUID sample(s).")

        worst_ratio = max(ratios, default=0.0)
        self.set_warning("red" if worst_ratio > 5.0 else "orange" if worst_ratio > 1.0 else "")

    def _update_holder_stats(self, block_result: object | None, holder_status: object | None) -> None:
        """Render VB6 Sig/Holder and Sig/Induced plus holder identity and age.

        These stay ``N/A`` until a backend supplies a real bracketed block and
        an installed holder correction, so an operator can never read a stale
        or absent holder as a valid one.
        """

        if block_result is not None:
            self._sig_holder.setText(f"{float(block_result.sig_holder):.2f}")
            self._sig_holder.setToolTip("Average moment / holder moment for the last block.")
            self._sig_induced.setText(f"{float(block_result.sig_induced):.2f}")
            self._sig_induced.setToolTip("Average moment / rotational asymmetry for the last block.")
        else:
            self._sig_holder.setText("N/A")
            self._sig_holder.setToolTip(
                "Requires a bracketed measurement block from the active hardware backend."
            )
            self._sig_induced.setText("N/A")
            self._sig_induced.setToolTip(
                "Requires a bracketed measurement block from the active hardware backend."
            )

        if holder_status is None or not getattr(holder_status, "present", False):
            self._holder_id.setText("None")
            self._holder_id.setToolTip(
                getattr(holder_status, "reason", "") or "No holder measurement recorded."
            )
            self._holder_moment.setText("N/A")
            self._holder_age.setText("N/A")
            return

        valid = bool(getattr(holder_status, "valid", False))
        identity = str(getattr(holder_status, "holder_id", "") or "unknown")
        self._holder_id.setText(identity if valid else f"{identity} (invalid)")
        tooltip = str(getattr(holder_status, "record_version", "") or identity)
        reason = str(getattr(holder_status, "reason", "") or "")
        if getattr(holder_status, "simulated", False):
            tooltip = f"SIMULATED holder correction. {tooltip}"
        if reason:
            tooltip = f"{tooltip} - {reason}"
        self._holder_id.setToolTip(tooltip)

        magnitude = getattr(holder_status, "magnitude_emu", None)
        self._holder_moment.setText("N/A" if magnitude is None else f"{float(magnitude):.3e} emu")
        asymmetry = getattr(holder_status, "asymmetry_ratio", None)
        if asymmetry is not None:
            self._holder_moment.setToolTip(f"Induced/holder asymmetry ratio {float(asymmetry):.3f}")

        age = getattr(holder_status, "age_seconds", None)
        if age is None:
            self._holder_age.setText("N/A")
        elif age < 3600:
            self._holder_age.setText(f"{float(age) / 60.0:.0f} min")
        else:
            self._holder_age.setText(f"{float(age) / 3600.0:.1f} h")

    def _plot_step_complete(self, result: StepResult, *, moment_Am2: float) -> None:
        self._plot_traces["step"].append(float(result.step_idx + 1))
        self._plot_traces["moment"].append(moment_Am2)
        self._plot_traces["dec"].append(float(result.step.gdec))
        self._plot_traces["inc"].append(float(result.step.ginc))
        self._plot_traces["susceptibility"].append(float(result.susceptibility))
        self._refresh_measurement_plot()

    def _refresh_measurement_plot(self) -> None:
        steps = list(self._plot_traces["step"])
        if self._plot_fallback is not None:
            self._plot_fallback.set_series(
                steps,
                list(self._plot_traces["moment"]),
                list(self._plot_traces["dec"]),
                list(self._plot_traces["inc"]),
            )
        if not self._plot_curves:
            return
        if not steps:
            for curve in self._plot_curves.values():
                curve.setData([], [])
            return
        self._plot_curves["moment"].setData(
            steps,
            list(self._plot_traces["moment"]),
            skipFiniteCheck=True,
        )
        self._plot_curves["susceptibility"].setData(
            steps,
            list(self._plot_traces["susceptibility"]),
            skipFiniteCheck=True,
        )
        self._plot_curves["zij"].setData(
            list(self._plot_traces["dec"]),
            list(self._plot_traces["inc"]),
            skipFiniteCheck=True,
        )

    def _clear_measurement_plot(self) -> None:
        self._completed_steps.clear()
        for trace in self._plot_traces.values():
            trace.clear()
        for curve in self._plot_curves.values():
            curve.setData([], [])
        if self._plot_fallback is not None:
            self._plot_fallback.clear()

    def _write_quicklook_sidecar(self) -> Path | None:
        if self._current_output_dir is None or not self._completed_steps:
            return None
        summary = build_quicklook_summary(
            [step.sdx for step in self._completed_steps],
            [step.sdy for step in self._completed_steps],
            [step.sdz for step in self._completed_steps],
            [step.demag_label for step in self._completed_steps],
        )
        summary["provenance"] = {
            "kind": "simulated" if self._current_run_simulated else "measurement",
            "simulated": self._current_run_simulated,
            "statement": (
                "Current run used a simulated backend; not hardware evidence."
                if self._current_run_simulated
                else "Measurement quicklook data."
            ),
        }
        written = write_quicklook_json(
            self._current_output_dir / "quicklook.json",
            summary,
        )
        self._refresh_quicklook_artifact_index(written)
        return written

    def _refresh_quicklook_artifact_index(self, quicklook_path: Path) -> None:
        """Reconcile the worker index after the UI publishes its quicklook."""

        index_path = quicklook_path.parent / "artifact_index.json"
        if not index_path.exists():
            return
        payload = json.loads(index_path.read_text(encoding="utf-8"))
        artifacts = payload.get("artifacts")
        if not isinstance(artifacts, list):
            raise ValueError("Artifact index does not contain an artifacts list.")
        matching = [
            entry
            for entry in artifacts
            if isinstance(entry, dict) and entry.get("name") == "quicklook_summary"
        ]
        if len(matching) != 1:
            raise ValueError(
                "Artifact index must contain exactly one quicklook_summary entry."
            )
        entry = matching[0]
        entry["relative_path"] = quicklook_path.name
        entry["exists"] = True
        entry["size_bytes"] = quicklook_path.stat().st_size
        write_quicklook_json(index_path, payload)

    @QtCore.Slot(bool)
    def _on_run_finished(self, aborted: bool) -> None:
        self._pending_run_finished = aborted
        self._settle_measurement_worker()

    @QtCore.Slot()
    def _settle_measurement_worker(self) -> None:
        if self._pending_run_finished is None:
            return
        if self._worker is not None and self._worker.isRunning():
            self._worker_settle_timer.start(10)
            return
        self._worker_settle_timer.stop()
        aborted, self._pending_run_finished = self._pending_run_finished, None
        sample_name = self._current_sample
        if self._worker is not None:
            self._worker.deleteLater()
        self._worker = None
        self._pause_btn.setText("⏸  Pause Run")
        self._release_measurement_ownership()
        sidecar_error: Exception | None = None
        if not aborted:
            try:
                self._write_quicklook_sidecar()
            except Exception as exc:  # pragma: no cover - filesystem failure path
                sidecar_error = exc
        mw = self.window()
        if hasattr(mw, "stop_run"):
            mw.stop_run()
        msg = "Run halted by user." if aborted else "Sequence complete!"
        if sidecar_error is not None:
            msg = f"{msg} Quicklook sidecar not written: {sidecar_error}"
        if hasattr(mw, "set_status"):
            mw.set_status(msg)
        self.sample_run_finished.emit(aborted, sample_name)

    @QtCore.Slot(str)
    def _on_error(self, msg: str) -> None:
        self._last_run_error = True
        QtWidgets.QMessageBox.critical(self, "Measurement Error", msg)

    def take_last_run_error(self) -> bool:
        """Return and clear whether the most recent run reported an error."""
        had_error = self._last_run_error
        self._last_run_error = False
        return had_error

    @QtCore.Slot(str)
    def _on_phase_changed(self, phase: str) -> None:
        mw = self.window()
        if hasattr(mw, "set_flow_state"):
            if phase in {
                "running",
                "paused",
                "halted",
                "idle",
                "preflight",
                "loading",
                "treating",
                "positioning",
                "measuring",
                "validating",
                "saving",
                "returning",
                "complete",
                "error",
            }:
                mw.set_flow_state(phase)

    def _acquire_ownerships(
        self,
        window: object,
        *,
        owner: str,
        queue_run: bool,
        needs_susceptibility: bool = False,
    ) -> None:
        """Acquire runtime ownership for sample measurement and queue movement."""
        try:
            if hasattr(window, "acquire_measurement_device"):
                lease = window.acquire_measurement_device(owner)  # type: ignore[arg-type]
                self._leases.append(lease)

            if queue_run and hasattr(window, "acquire_device"):
                lease = window.acquire_device("changer", owner)  # type: ignore[arg-type]
                self._leases.append(lease)

            if needs_susceptibility and hasattr(window, "acquire_device"):
                lease = window.acquire_device("susceptibility", owner)  # type: ignore[arg-type]
                self._leases.append(lease)
        except Exception:
            self._release_measurement_ownership()
            raise

    def _release_measurement_ownership(self) -> None:
        for lease in reversed(self._leases):
            try:
                lease.release()
            except Exception:
                pass
        self._leases.clear()

    @QtCore.Slot(str)
    def _on_preflight_warning(self, msg: str) -> None:
        mw = self.window()
        if hasattr(mw, "set_status"):
            mw.set_status(f"Preflight warning: {msg}")

    # ── Public update API ─────────────────────────────────────────────────────
    def update_squid(self, x: str, y: str, z: str) -> None:
        self._squid_x.setText(x)
        self._squid_y.setText(y)
        self._squid_z.setText(z)

    def update_calculated(self, dec: str, inc: str, moment: str, csd: str) -> None:
        self._calc_dec.setText(dec)
        self._calc_inc.setText(inc)
        self._calc_moment.setText(moment)
        self._calc_csd.setText(csd)

    def set_warning(self, level: str) -> None:
        """level: '' | 'orange' | 'red'"""
        self._warn_orange.setVisible(level == "orange")
        self._warn_red.setVisible(level == "red")

    # ── Helpers ───────────────────────────────────────────────────────────────
    def _goto(self, key: str) -> None:
        mw = self.window()
        if hasattr(mw, "navigate_to"):
            mw.navigate_to(key)


def _hline() -> QtWidgets.QFrame:
    f = QtWidgets.QFrame()
    f.setFrameShape(QtWidgets.QFrame.HLine)
    f.setStyleSheet("color: rgba(122,2,25,0.12); margin: 2px 0;")
    return f


def _hline_w() -> QtWidgets.QWidget:
    w = QtWidgets.QWidget()
    w.setFixedHeight(1)
    w.setStyleSheet("background: rgba(122,2,25,0.12);")
    return w


def _axis_hdr(text: str) -> QtWidgets.QLabel:
    lbl = QtWidgets.QLabel(text)
    lbl.setObjectName("readLbl")
    lbl.setAlignment(QtCore.Qt.AlignCenter)
    return lbl


def _resolve_measurement_backend(
    window: object | None,
    fallback: MeasurementBackend,
) -> MeasurementBackend:
    """Resolve an active measurement backend from an owning window.

    This safely handles older call patterns and keeps startup flows resilient.
    """
    if window is not None:
        provider = getattr(window, "measurement_backend", None)
        if callable(provider):
            candidate = provider()
            required = ("read_squid", "set_demag_step", "read_susceptibility", "preflight", "is_available")
            if all(hasattr(candidate, attr) for attr in required):
                return candidate
    return fallback


