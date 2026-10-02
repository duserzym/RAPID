from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path

from PySide6 import QtCore, QtWidgets

from rapid_main.rockmag import (
    RockmagRoutinePlan,
    compile_rockmag_routine,
    rockmag_af_demag,
    rockmag_the_works,
)


@dataclass
class SequenceConfig:
    do_nrm: bool = False
    nrm_3axis: bool = False
    do_rrm: bool = False
    rrm_step: float = 0.5
    rrm_max: float = 2.0
    rrm_af_field: float = 100.0
    rrm_do_negative: bool = False
    do_arm: bool = False
    arm_step: float = 5.0
    arm_max: float = 100.0
    arm_af_field: float = 100.0
    do_irm_af: bool = False
    irm_log_factor: float = 0.25
    irm_min_step: float = 5.0
    irm_af_max: float = 800.0
    irm_irm_max: float = 1200.0
    do_backfield: bool = False
    do_susceptibility: bool = False


class SequencePanel(QtWidgets.QWidget):
    """Experiment sequence builder panel.

    Maps to: frmRockmagRoutine in VB6.
    Left side: preset buttons + individual step controls.
    Right side: live text preview of the configured sequence.
    """

    def __init__(self, parent: QtWidgets.QWidget | None = None) -> None:
        super().__init__(parent)
        self._cfg = SequenceConfig()
        self._compiled_routine_labels: list[str] | None = None
        self._compiled_routine_plan: RockmagRoutinePlan | None = None
        self._applying_preset = False
        self._current_path: Path | None = None
        self._dirty = False

        root = QtWidgets.QHBoxLayout(self)
        root.setContentsMargins(16, 12, 16, 16)
        root.setSpacing(12)

        left_panel = QtWidgets.QWidget()
        lv = QtWidgets.QVBoxLayout(left_panel)
        lv.setContentsMargins(0, 0, 0, 0)
        lv.setSpacing(12)
        lv.addWidget(self._build_presets_card())
        lv.addWidget(self._build_steps_card())
        lv.addStretch()
        root.addWidget(left_panel, 3)
        root.addWidget(self._build_preview_card(), 2)
        self._rebuild_preview()

    # ── Presets ───────────────────────────────────────────────────────────────
    def _build_presets_card(self) -> QtWidgets.QFrame:
        card = QtWidgets.QFrame()
        card.setObjectName("card")
        cl = QtWidgets.QVBoxLayout(card)
        cl.setContentsMargins(18, 14, 18, 14)
        cl.setSpacing(8)

        hdr = QtWidgets.QLabel("PRESETS")
        hdr.setObjectName("sectionHdr")
        cl.addWidget(hdr)

        note = QtWidgets.QLabel("Click a preset to auto-fill all step parameters:")
        note.setStyleSheet("color: #7a6f6e; font-size: 12px;")
        cl.addWidget(note)

        row = QtWidgets.QHBoxLayout()
        row.setContentsMargins(0, 0, 0, 0)
        row.setSpacing(10)

        hw_btn = QtWidgets.QPushButton("Hawaiian AF Preset")
        hw_btn.setMinimumHeight(52)
        hw_btn.setSizePolicy(
            QtWidgets.QSizePolicy.Policy.Expanding,
            QtWidgets.QSizePolicy.Policy.Fixed,
        )
        hw_btn.setToolTip("Hawaiian Standard AF\n(25, 50, 100, 200, 400, 800 mT)")
        hw_btn.clicked.connect(self._preset_hawaiian)

        rw_btn = QtWidgets.QPushButton('Rockmag the Works')
        rw_btn.setMinimumHeight(52)
        rw_btn.setObjectName("accent")
        rw_btn.setSizePolicy(
            QtWidgets.QSizePolicy.Policy.Expanding,
            QtWidgets.QSizePolicy.Policy.Fixed,
        )
        rw_btn.setToolTip("NRM + ARM + IRM + AF/IRM")
        rw_btn.clicked.connect(self._preset_works)

        row.addWidget(hw_btn, 1)
        row.addWidget(rw_btn, 1)
        cl.addLayout(row)
        return card

    # ── Step configuration ────────────────────────────────────────────────────
    def _build_steps_card(self) -> QtWidgets.QFrame:
        card = QtWidgets.QFrame()
        card.setObjectName("card")
        cl = QtWidgets.QVBoxLayout(card)
        cl.setContentsMargins(18, 14, 18, 18)
        cl.setSpacing(10)

        hdr = QtWidgets.QLabel("MEASUREMENT STEPS")
        hdr.setObjectName("sectionHdr")
        cl.addWidget(hdr)
        cl.addSpacing(10)

        # ── NRM ──
        self._chk_nrm = QtWidgets.QCheckBox("Measure and AF demagnetize NRM")
        self._chk_nrm.toggled.connect(self._rebuild_preview)
        cl.addWidget(self._chk_nrm)

        nrm_extra = QtWidgets.QWidget()
        nrm_extra_l = QtWidgets.QHBoxLayout(nrm_extra)
        nrm_extra_l.setContentsMargins(22, 0, 0, 0)
        self._chk_nrm_3axis = QtWidgets.QCheckBox("along all three axes")
        self._chk_nrm_3axis.setEnabled(False)
        nrm_extra_l.addWidget(self._chk_nrm_3axis)
        nrm_extra_l.addStretch()
        self._chk_nrm.toggled.connect(self._chk_nrm_3axis.setEnabled)
        self._chk_nrm.toggled.connect(self._rebuild_preview)
        self._chk_nrm_3axis.toggled.connect(self._rebuild_preview)
        cl.addWidget(nrm_extra)
        cl.addSpacing(6)

        # ── RRM ──
        self._chk_rrm = QtWidgets.QCheckBox("RRM  — rps step")
        self._chk_rrm.toggled.connect(self._rebuild_preview)
        rrm_row, self._rrm_step, self._rrm_max, self._rrm_af = self._param_row(
            self._chk_rrm,
            [("step (rps)", 0.5, 0.01, 10.0),
             ("to (rps)",   2.0, 0.01, 10.0),
             ("AF field (mT)", 100.0, 1.0, 1200.0)],
        )
        self._chk_rrm_neg = QtWidgets.QCheckBox("and negative rotations")
        self._chk_rrm_neg.setContentsMargins(22, 0, 0, 0)
        self._chk_rrm_neg.setEnabled(False)
        self._chk_rrm.toggled.connect(self._chk_rrm_neg.setEnabled)
        self._chk_rrm_neg.toggled.connect(self._rebuild_preview)
        cl.addWidget(self._chk_rrm)
        cl.addWidget(rrm_row)
        cl.addWidget(self._chk_rrm_neg)
        cl.addSpacing(6)

        # ── ARM ──
        self._chk_arm = QtWidgets.QCheckBox("ARM  — step size (G)")
        self._chk_arm.toggled.connect(self._rebuild_preview)
        arm_row, self._arm_step, self._arm_max, self._arm_af = self._param_row(
            self._chk_arm,
            [("step (G)",     5.0,   0.1, 500.0),
             ("to (G)",       100.0, 0.1, 2000.0),
             ("in AF (mT)",   100.0, 1.0, 1200.0)],
        )
        cl.addWidget(self._chk_arm)
        cl.addWidget(arm_row)
        cl.addSpacing(6)

        # ── AF / IRM ──
        self._chk_irm = QtWidgets.QCheckBox("AF / IRM  — log step factor")
        self._chk_irm.toggled.connect(self._rebuild_preview)
        irm_row, self._irm_log, self._irm_min, self._irm_af_max, self._irm_irm_max = \
            self._param_row(
                self._chk_irm,
                [("log factor",   0.25, 0.01, 2.0),
                 ("min step (G)", 5.0,  0.1,  100.0),
                 ("AF max (mT)",  800.0, 1.0,  2000.0),
                 ("IRM max (G)",  1200.0, 1.0, 5000.0)],
            )
        cl.addWidget(self._chk_irm)
        cl.addWidget(irm_row)
        cl.addSpacing(6)

        # ── DC Backfield ──
        self._chk_backfield = QtWidgets.QCheckBox(
            "DC Backfield Demag  (via backfield IRM)"
        )
        self._chk_backfield.toggled.connect(self._rebuild_preview)
        cl.addWidget(self._chk_backfield)
        cl.addSpacing(4)

        # ── Susceptibility ──
        self._chk_susc = QtWidgets.QCheckBox("Measure Susceptibility")
        self._chk_susc.toggled.connect(self._rebuild_preview)
        cl.addWidget(self._chk_susc)

        return card

    def _param_row(
        self,
        parent_chk: QtWidgets.QCheckBox,
        params: list[tuple[str, float, float, float]],
    ) -> tuple[QtWidgets.QWidget, ...]:
        row = QtWidgets.QWidget()
        row_layout = QtWidgets.QVBoxLayout(row)
        row_layout.setContentsMargins(18, 2, 6, 2)
        row_layout.setSpacing(6)
        spins = []
        for label, default, lo, hi in params:
            field = QtWidgets.QWidget()
            field_layout = QtWidgets.QHBoxLayout(field)
            field_layout.setContentsMargins(0, 0, 0, 0)
            field_layout.setSpacing(8)

            lbl = QtWidgets.QLabel(f"{label}:")
            lbl.setObjectName("readLbl")
            lbl.setAlignment(
                QtCore.Qt.AlignmentFlag.AlignLeft | QtCore.Qt.AlignmentFlag.AlignVCenter
            )
            lbl.setSizePolicy(
                QtWidgets.QSizePolicy.Policy.Preferred,
                QtWidgets.QSizePolicy.Policy.Fixed,
            )
            label_min_width = lbl.fontMetrics().horizontalAdvance(f"{label}:")
            lbl.setMinimumWidth(int(label_min_width) + 12)
            spin = QtWidgets.QDoubleSpinBox()
            spin.setRange(lo, hi)
            spin.setValue(default)
            spin.setAlignment(QtCore.Qt.AlignmentFlag.AlignRight)
            spin.setSizePolicy(
                QtWidgets.QSizePolicy.Policy.MinimumExpanding,
                QtWidgets.QSizePolicy.Policy.Fixed,
            )
            # Reserve enough room for the widest configured value (no visible clipping).
            sample_text = spin.textFromValue(hi)
            content_width = spin.fontMetrics().horizontalAdvance(sample_text)
            spin.setMinimumWidth(max(170, content_width + 68))
            spin.setMaximumWidth(max(260, content_width + 92))
            spin.setContentsMargins(0, 0, 0, 0)
            spin.setDecimals(3)
            spin.setButtonSymbols(QtWidgets.QAbstractSpinBox.ButtonSymbols.NoButtons)
            spin.setFixedHeight(30)
            spin.setEnabled(False)
            spin.valueChanged.connect(self._rebuild_preview)
            parent_chk.toggled.connect(spin.setEnabled)
            field_layout.addWidget(lbl)
            field_layout.addWidget(spin, 1)
            row_layout.addWidget(field)
            spins.append(spin)
        return (row, *spins)

    # ── Preview ───────────────────────────────────────────────────────────────
    def _build_preview_card(self) -> QtWidgets.QFrame:
        card = QtWidgets.QFrame()
        card.setObjectName("card")
        cl = QtWidgets.QVBoxLayout(card)
        cl.setContentsMargins(18, 14, 18, 14)
        cl.setSpacing(8)

        self._preview_header = QtWidgets.QLabel("SEQUENCE PREVIEW")
        self._preview_header.setObjectName("sectionHdr")
        cl.addWidget(self._preview_header)

        self._preview = QtWidgets.QPlainTextEdit()
        self._preview.setObjectName("console")
        self._preview.setReadOnly(True)
        self._preview.setLineWrapMode(QtWidgets.QPlainTextEdit.LineWrapMode.WidgetWidth)
        self._preview.setHorizontalScrollBarPolicy(QtCore.Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        cl.addWidget(self._preview, 1)

        btn_row = QtWidgets.QHBoxLayout()
        gen_btn = QtWidgets.QPushButton("▶  Generate Step List")
        gen_btn.setObjectName("accent")
        gen_btn.clicked.connect(self._generate)
        btn_row.addWidget(gen_btn)

        save_btn = QtWidgets.QPushButton("💾  Save…")
        save_btn.clicked.connect(self._save_sequence)
        btn_row.addWidget(save_btn)

        load_btn = QtWidgets.QPushButton("📂  Load…")
        load_btn.clicked.connect(self._load_sequence)
        btn_row.addWidget(load_btn)
        cl.addLayout(btn_row)

        self._step_count_lbl = QtWidgets.QLabel("0 steps configured")
        self._step_count_lbl.setStyleSheet("color: #9a8885; font-size: 12px;")
        cl.addWidget(self._step_count_lbl)
        return card

    # ── Step label generation ─────────────────────────────────────────────────

    def generate_labels(self) -> list[str]:
        """Compute the flat list of step labels from the current configuration."""
        if self._compiled_routine_labels is not None:
            return list(self._compiled_routine_labels)

        labels: list[str] = []

        if self._chk_nrm.isChecked():
            if self._chk_nrm_3axis.isChecked():
                labels += ["NRM-X", "NRM-Y", "NRM-Z"]
            else:
                labels.append("NRM")

        if self._chk_rrm.isChecked():
            step = self._rrm_step.value()
            top  = self._rrm_max.value()
            v    = step
            while v <= top + 1e-9:
                labels.append(f"RRM{v:.2f}".rstrip("0").rstrip("."))
                v += step
            if self._chk_rrm_neg.isChecked():
                v = step
                while v <= top + 1e-9:
                    labels.append(f"RRM-{v:.2f}".rstrip("0").rstrip("."))
                    v += step

        if self._chk_arm.isChecked():
            step = self._arm_step.value()
            top  = self._arm_max.value()
            v    = step
            while v <= top + 1e-9:
                labels.append(f"ARM{int(round(v))}" if v == round(v) else f"ARM{v:.2f}".rstrip("0").rstrip("."))
                v += step

        if self._chk_irm.isChecked():
            log_factor = self._irm_log.value()
            min_step   = self._irm_min.value()
            af_max     = self._irm_af_max.value()
            irm_max    = self._irm_irm_max.value()
            # AF demagnetisation steps (log-spaced)
            labels += _log_steps("AF", min_step, af_max, log_factor)
            # IRM acquisition steps (log-spaced)
            labels += _log_steps("IRM", min_step, irm_max, log_factor)

        if self._chk_backfield.isChecked():
            labels.append("IRM-BF")

        if self._chk_susc.isChecked():
            labels.append("SUSC")

        return labels

    @QtCore.Slot()
    def _rebuild_preview(self) -> None:
        if not self._applying_preset and self.sender() is not None:
            self._compiled_routine_labels = None
            self._compiled_routine_plan = None
            self._set_dirty(True)
        labels = self.generate_labels()
        self._render_labels(labels)

    def _render_labels(self, labels: list[str], *, loaded: bool = False) -> None:
        lines = ["Configured measurement sequence:\n"]
        if labels:
            for lbl in labels:
                lines.append(f"  {lbl}")
        else:
            lines.append("  (no steps selected — choose a preset or check steps above)")
        self._preview.setPlainText("\n".join(lines))
        n = len(labels)
        suffix = "loaded from file" if loaded else "configured"
        self._step_count_lbl.setText(f"{n} step{'s' if n != 1 else ''} {suffix}")
        self._update_document_caption()
        # Push labels to MainWindow runtime estimator
        mw = self.window()
        if hasattr(mw, "load_sequence_labels"):
            mw.load_sequence_labels(labels)
        if hasattr(mw, "set_rockmag_routine_plan"):
            mw.set_rockmag_routine_plan(self._compiled_routine_plan)

    def _set_dirty(self, dirty: bool) -> None:
        self._dirty = bool(dirty)
        if hasattr(self, "_preview_header"):
            self._update_document_caption()

    def _update_document_caption(self) -> None:
        name = self._current_path.name if self._current_path is not None else "Untitled"
        marker = " • Unsaved" if self._dirty else ""
        self._preview_header.setText(f"SEQUENCE PREVIEW — {name}{marker}")

    def has_unsaved_changes(self) -> bool:
        return self._dirty

    def confirm_discard_changes(self) -> bool:
        if not self._dirty:
            return True
        name = self._current_path.name if self._current_path is not None else "Untitled sequence"
        answer = QtWidgets.QMessageBox.question(
            self,
            "Unsaved sequence",
            f"{name} has unsaved changes. Discard them?",
            QtWidgets.QMessageBox.StandardButton.Discard
            | QtWidgets.QMessageBox.StandardButton.Cancel,
            QtWidgets.QMessageBox.StandardButton.Cancel,
        )
        return answer == QtWidgets.QMessageBox.StandardButton.Discard

    def _generate(self) -> None:
        labels = self.generate_labels()
        if not labels:
            QtWidgets.QMessageBox.information(
                self, "Generate", "No steps configured."
            )
            return
        QtWidgets.QMessageBox.information(
            self, "Step List Generated",
            f"{len(labels)} steps ready.\n\n"
            "The sequence has been sent to the status bar estimator.\n"
            "Use 'Save…' to export the step list to a file.",
        )

    # ── Save / Load ───────────────────────────────────────────────────────────

    def _save_sequence(self) -> bool:
        labels = self.generate_labels()
        if not labels:
            QtWidgets.QMessageBox.warning(self, "Save Sequence", "No steps to save.")
            return False
        path, _ = QtWidgets.QFileDialog.getSaveFileName(
            self,
            "Save Sequence",
            str(self._current_path or ""),
            "JSON sequence (*.json);;Text file (*.txt)",
        )
        if not path:
            return False
        return self.save_to_path(Path(path))

    def save_to_path(self, path: Path) -> bool:
        from rapid_main.io.sequence_io import save_sequence_json, save_sequence_txt
        labels = self.generate_labels()
        try:
            if path.suffix.lower() == ".json":
                save_sequence_json(path, labels)
            else:
                save_sequence_txt(path, labels)
        except (OSError, ValueError) as exc:
            QtWidgets.QMessageBox.critical(
                self, "Save Sequence", f"Unable to save sequence:\n\n{exc}"
            )
            return False
        self._current_path = path
        self._set_dirty(False)
        self._step_count_lbl.setText(
            f"{len(labels)} step{'s' if len(labels) != 1 else ''} saved"
        )
        return True

    def _load_sequence(self) -> None:
        path, _ = QtWidgets.QFileDialog.getOpenFileName(
            self, "Load Sequence", "", "JSON sequence (*.json);;Text file (*.txt);;All files (*)"
        )
        if not path:
            return
        if not self.confirm_discard_changes():
            return
        from rapid_main.io.sequence_io import SequenceFormatError, load_sequence_strict
        try:
            labels = load_sequence_strict(Path(path))
        except SequenceFormatError as exc:
            QtWidgets.QMessageBox.critical(
                self, "Load Sequence", f"Unable to load sequence:\n\n{exc}"
            )
            return
        self.load_labels(labels, source_path=Path(path))

    def load_labels(self, labels: list[str], *, source_path: Path | None = None) -> None:
        """Install imported labels as the active executable and saveable sequence."""
        self._compiled_routine_labels = list(labels)
        self._compiled_routine_plan = None
        self._current_path = source_path
        self._set_dirty(False)
        self._render_labels(list(labels), loaded=source_path is not None)

    def clear_for_new_session(self) -> None:
        self._applying_preset = True
        try:
            for checkbox in (
                self._chk_nrm,
                self._chk_nrm_3axis,
                self._chk_rrm,
                self._chk_rrm_neg,
                self._chk_arm,
                self._chk_irm,
                self._chk_backfield,
                self._chk_susc,
            ):
                checkbox.setChecked(False)
        finally:
            self._applying_preset = False
        self._compiled_routine_labels = None
        self._compiled_routine_plan = None
        self._current_path = None
        self._set_dirty(False)
        self._rebuild_preview()

    # ── Preset loaders ────────────────────────────────────────────────────────
    def _preset_hawaiian(self) -> None:
        plan = compile_rockmag_routine(
            rockmag_af_demag(
                (25.0, 50.0, 100.0, 200.0, 400.0, 800.0),
                name="Hawaiian AF Preset",
            )
        )
        self._compiled_routine_plan = plan
        self._compiled_routine_labels = plan.to_queue_labels()
        self._applying_preset = True
        try:
            for chk in (self._chk_nrm, self._chk_rrm, self._chk_arm,
                        self._chk_irm, self._chk_backfield, self._chk_susc):
                chk.setChecked(False)
            self._chk_nrm.setChecked(True)
        finally:
            self._applying_preset = False
        self._current_path = None
        self._set_dirty(True)
        self._rebuild_preview()

    def _preset_works(self) -> None:
        plan = compile_rockmag_routine(rockmag_the_works())
        self._compiled_routine_plan = plan
        self._compiled_routine_labels = plan.to_queue_labels()
        self._applying_preset = True
        try:
            self._chk_nrm.setChecked(True)
            self._chk_arm.setChecked(True)
            self._chk_irm.setChecked(True)
            self._chk_backfield.setChecked(True)
            self._chk_susc.setChecked(True)
        finally:
            self._applying_preset = False
        self._current_path = None
        self._set_dirty(True)
        self._rebuild_preview()


# ── Module-level helpers ──────────────────────────────────────────────────────

def _log_steps(prefix: str, min_val: float, max_val: float, log_factor: float) -> list[str]:
    """Generate logarithmically-spaced step labels.

    Starting from *min_val*, each step multiplies by (1 + *log_factor*).
    The final step at *max_val* is always included.
    Returns list of strings like ``["AF5", "AF7", "AF10", ..., "AF800"]``.
    """
    if min_val <= 0 or max_val <= 0 or log_factor <= 0:
        return []
    steps: list[float] = []
    v = min_val
    while v < max_val * (1 - 1e-9):
        steps.append(v)
        v *= (1.0 + log_factor)
    steps.append(max_val)
    labels: list[str] = []
    for s in steps:
        if s == round(s):
            labels.append(f"{prefix}{int(round(s))}")
        else:
            labels.append(f"{prefix}{s:.1f}".rstrip("0").rstrip("."))
    return labels

