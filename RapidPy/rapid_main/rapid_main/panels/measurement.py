from __future__ import annotations

import math
from pathlib import Path
from typing import Optional

from PySide6 import QtCore, QtWidgets

from rapid_main.data_model import SpecimenMeta
from rapid_main.measurement_worker import MeasurementWorker, NoCommBackend, StepResult


class MeasurementPanel(QtWidgets.QWidget):
    """Live measurement display panel.

    Maps to: frmMeasure (Measurement Window) in VB6.

    Layout:
        Top strip  — current sample, depth, step, coordinates selector
        Col A (22%) — flow controls, coordinate frame selector
        Col B (42%) — live SQUID readings card + measurement stats card
        Col C (36%) — moment vs step plot placeholder
        Bottom     — quality warning banners (hidden until data arrives)
    """

    def __init__(self, parent: QtWidgets.QWidget | None = None) -> None:
        super().__init__(parent)
        self._worker: Optional[MeasurementWorker] = None
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

        ov.addLayout(stats_grid)
        ov.addStretch()
        return outer

    # ── Plot placeholder ──────────────────────────────────────────────────────
    def _build_plot_card(self) -> QtWidgets.QFrame:
        card = QtWidgets.QFrame()
        card.setObjectName("card")
        cl = QtWidgets.QVBoxLayout(card)
        cl.setContentsMargins(14, 14, 14, 14)
        cl.setSpacing(8)

        hdr = QtWidgets.QLabel("MOMENT vs. TREATMENT STEP")
        hdr.setObjectName("sectionHdr")
        cl.addWidget(hdr)

        placeholder = QtWidgets.QLabel(
            "Live Zijderveld / moment plot\n\n"
            "(pyqtgraph PlotWidget — Phase 2)"
        )
        placeholder.setAlignment(QtCore.Qt.AlignCenter)
        placeholder.setStyleSheet(
            "QLabel { border: 2px dashed rgba(122,2,25,0.18); "
            "border-radius: 14px; color: #c4b7b3; font-size: 13px; }"
        )
        cl.addWidget(placeholder, 1)

        stats_btn = QtWidgets.QPushButton("📊  Show Stats Window")
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
        PlotsDialog(self).exec()

    # ── Worker control ────────────────────────────────────────────────────────

    def _on_start(self) -> None:
        if self._worker is not None and self._worker.is_paused:
            self._worker.resume()
            return
        if self._worker is not None and self._worker.isRunning():
            return  # already running

        mw = self.window()
        labels = getattr(mw, "_sequence_labels", [])
        if not labels:
            QtWidgets.QMessageBox.warning(
                self, "No Sequence",
                "No measurement sequence is loaded.\n"
                "Go to the Sequence panel and generate or load a step list first.",
            )
            return

        cfg    = getattr(mw, "config", None)
        op     = cfg.general.operator if cfg else ""
        out    = Path(cfg.general.data_dir) if cfg and cfg.general.data_dir else Path.home() / "RAPID_data"
        meta   = SpecimenMeta(
            name=getattr(mw, "_current_sample", "UNKNOWN"),
            comment="",
            sample="",
            site="",
            location="",
        )

        self._worker = MeasurementWorker(
            meta=meta,
            labels=labels,
            output_dir=out / meta.name,
            backend=NoCommBackend(),
            operator=op,
            parent=self,
        )
        self._worker.step_started.connect(self._on_step_started)
        self._worker.step_complete.connect(self._on_step_complete)
        self._worker.run_finished.connect(self._on_run_finished)
        self._worker.error_occurred.connect(self._on_error)
        self._worker.start()

        if hasattr(mw, "start_run"):
            mw.start_run(labels)

    def _on_pause(self) -> None:
        if self._worker is not None:
            if self._worker.is_paused:
                self._worker.resume()
                self._pause_btn.setText("⏸  Pause Run")
            else:
                self._worker.pause()
                self._pause_btn.setText("▶  Resume Run")

    def _on_halt(self) -> None:
        if self._worker is not None:
            self._worker.halt()

    @QtCore.Slot(int, str)
    def _on_step_started(self, idx: int, label: str) -> None:
        mw = self.window()
        total = len(getattr(mw, "_sequence_labels", []))
        self._meas_step.setText(f"{idx + 1} / {total}")
        self._meas_treat.setText(label)
        if hasattr(mw, "set_step"):
            mw.set_step(label)
        if hasattr(mw, "advance_step") and idx > 0:
            mw.advance_step()

    @QtCore.Slot(object)
    def _on_step_complete(self, result: StepResult) -> None:
        step = result.step
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

    @QtCore.Slot(bool)
    def _on_run_finished(self, aborted: bool) -> None:
        self._worker = None
        self._pause_btn.setText("⏸  Pause Run")
        mw = self.window()
        if hasattr(mw, "stop_run"):
            mw.stop_run()
        msg = "Run halted by user." if aborted else "Sequence complete!"
        if hasattr(mw, "set_status"):
            mw.set_status(msg)

    @QtCore.Slot(str)
    def _on_error(self, msg: str) -> None:
        QtWidgets.QMessageBox.critical(self, "Measurement Error", msg)

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
