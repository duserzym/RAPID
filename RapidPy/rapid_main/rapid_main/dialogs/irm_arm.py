from __future__ import annotations
import threading

from PySide6 import QtCore, QtWidgets

from rapid_main.diagnostic_services import IrmArmBackend, IrmArmNoCommBackend
from rapid_main.glass_theme import set_semantic_status


class _TreatmentTask(QtCore.QThread):
    result = QtCore.Signal(str, bool)

    def __init__(self, operation, parent):
        super().__init__(parent)
        self.operation = operation

    def run(self):
        try:
            message = self.operation()
        except Exception as exc:
            self.result.emit(str(exc), False)
        else:
            self.result.emit(str(message), True)


class IrmArmDialog(QtWidgets.QDialog):
    """IRM / ARM control dialog — replaces VB6 frmIRMARM."""

    def __init__(
        self,
        parent: QtWidgets.QWidget | None = None,
        backend: IrmArmBackend | None = None,
        manual_arm=None,
    ) -> None:
        super().__init__(parent)
        self.setObjectName("glassDialog")
        self.setWindowTitle("IRM / ARM Control")
        self.setAccessibleName("IRM and ARM field control")
        self.setMinimumWidth(340)
        self.setWindowFlags(self.windowFlags() & ~QtCore.Qt.WindowContextHelpButtonHint)
        self._backend = backend or IrmArmNoCommBackend()
        self._manual_arm = manual_arm
        self._task = None
        self._cancel = threading.Event()
        self._close_pending = False
        self._build_ui()

    # ── UI ─────────────────────────────────────────────────────────────────
    def _build_ui(self) -> None:
        vl = QtWidgets.QVBoxLayout(self)
        vl.setContentsMargins(20, 16, 20, 16)
        vl.setSpacing(12)

        hdr = QtWidgets.QLabel("IRM / ARM Control")
        hdr.setObjectName("dialogTitle")
        hdr.setAccessibleName("IRM and ARM control title")
        vl.addWidget(hdr)

        # ── Mode selector ────────────────────────────────────────────────
        mode_row = QtWidgets.QHBoxLayout()
        mode_lbl = QtWidgets.QLabel("Mode:")
        mode_lbl.setObjectName("dialogSubtitle")
        self._mode = QtWidgets.QComboBox()
        self._mode.addItems(["IRM (Isothermal Remanence)", "ARM (Anhysteretic Remanence)"])
        self._mode.setAccessibleName("Field treatment mode")
        self._mode.currentIndexChanged.connect(self._on_mode_changed)
        mode_row.addWidget(mode_lbl)
        mode_row.addWidget(self._mode, 1)
        vl.addLayout(mode_row)
        self._sample_row = QtWidgets.QWidget()
        sample_layout = QtWidgets.QFormLayout(self._sample_row)
        self._sample_id = QtWidgets.QLineEdit()
        self._sample_id.setAccessibleName("Manual treatment specimen identity")
        sample_layout.addRow("Specimen:",self._sample_id)
        self._sample_row.setVisible(self._manual_arm is not None)
        vl.addWidget(self._sample_row)

        # ── IRM settings ─────────────────────────────────────────────────
        self._irm_grp = QtWidgets.QGroupBox("IRM Settings")
        fl = QtWidgets.QFormLayout(self._irm_grp)
        fl.setSpacing(8)
        fl.setLabelAlignment(QtCore.Qt.AlignRight)
        fl.setRowWrapPolicy(QtWidgets.QFormLayout.RowWrapPolicy.WrapLongRows)

        self._irm_field = QtWidgets.QDoubleSpinBox()
        self._irm_field.setRange(-2000 if self._manual_arm is not None else 0, 2000)
        self._irm_field.setValue(100)
        self._irm_field.setSuffix(" mT")
        self._irm_field.setSingleStep(10)
        self._irm_field.setAccessibleName("IRM peak direct-current field in millitesla")
        fl.addRow("Peak DC field:", self._irm_field)

        self._irm_axis = QtWidgets.QComboBox()
        self._irm_axis.addItems(["Z (up-axis)", "X", "Y"])
        self._irm_axis.setAccessibleName("IRM magnetization axis")
        fl.addRow("Magnetise axis:", self._irm_axis)

        self._irm_ramp = QtWidgets.QComboBox()
        self._irm_ramp.addItems(["Slow (60 s)", "Medium (30 s)", "Fast (10 s)"])
        self._irm_ramp.setAccessibleName("IRM ramp speed")
        fl.addRow("Ramp speed:", self._irm_ramp)
        if self._manual_arm is not None:
            self._irm_ramp.setEnabled(False)
            self._irm_ramp.setToolTip("Pulse charging uses calibrated capacitor readback and bounded deadlines.")

        vl.addWidget(self._irm_grp)

        # ── ARM settings ─────────────────────────────────────────────────
        self._arm_grp = QtWidgets.QGroupBox("ARM Settings")
        fl2 = QtWidgets.QFormLayout(self._arm_grp)
        fl2.setSpacing(8)
        fl2.setLabelAlignment(QtCore.Qt.AlignRight)
        fl2.setRowWrapPolicy(QtWidgets.QFormLayout.RowWrapPolicy.WrapLongRows)

        self._arm_peak_af = QtWidgets.QDoubleSpinBox()
        self._arm_peak_af_spin = self._arm_peak_af
        self._arm_peak_af.setRange(0, 200)
        self._arm_peak_af.setValue(80)
        self._arm_peak_af.setSuffix(" mT")
        self._arm_peak_af.setAccessibleName("ARM peak alternating field in millitesla")
        fl2.addRow("Peak AF field:", self._arm_peak_af)

        self._arm_bias = QtWidgets.QDoubleSpinBox()
        self._arm_bias.setRange(0, 100)
        self._arm_bias.setValue(0.05)
        self._arm_bias.setDecimals(3)
        self._arm_bias.setSingleStep(0.005)
        self._arm_bias.setSuffix(" mT")
        self._arm_bias.setAccessibleName("ARM bias field in millitesla")
        fl2.addRow("Bias field:", self._arm_bias)

        vl.addWidget(self._arm_grp)
        self._arm_grp.setVisible(False)

        # ── Status ───────────────────────────────────────────────────────
        self._status_lbl = QtWidgets.QLabel()
        self._status_lbl.setWordWrap(True)
        try:
            connected = bool(self._backend.is_connected())
        except Exception:
            connected = False
        try:
            backend_status = self._backend.status()
        except Exception as exc:
            backend_status = f"IRM/ARM status read failed: {exc}"
        self._set_status(
            backend_status,
            "ready"
            if connected
            else ("error" if "failed" in backend_status else "unavailable"),
        )
        vl.addWidget(self._status_lbl)

        # ── Control buttons ──────────────────────────────────────────────
        ctrl_row = QtWidgets.QHBoxLayout()
        self._apply_btn = QtWidgets.QPushButton("Apply Field")
        self._apply_btn.setObjectName("accent")
        self._apply_btn.setAccessibleName("Apply configured IRM or ARM field")
        self._apply_btn.setAccessibleDescription(
            "Runs the selected field treatment through the configured backend."
        )
        self._apply_btn.clicked.connect(self._apply)

        self._reset_btn = QtWidgets.QPushButton("Zero Field")
        self._reset_btn.setAccessibleName("Reset IRM and ARM field to zero")
        self._reset_btn.clicked.connect(self._reset)

        self._close_btn = QtWidgets.QPushButton("Close")
        self._close_btn.setAccessibleName("Close IRM and ARM control")
        self._close_btn.clicked.connect(self.close)
        self._cancel_btn = QtWidgets.QPushButton("Cancel Treatment")
        self._cancel_btn.clicked.connect(self._cancel_treatment)
        self._cancel_btn.setEnabled(False)

        self._apply_btn.setEnabled(connected)
        self._reset_btn.setEnabled(connected)

        ctrl_row.addWidget(self._apply_btn)
        ctrl_row.addWidget(self._reset_btn)
        ctrl_row.addWidget(self._cancel_btn)
        ctrl_row.addStretch()
        ctrl_row.addWidget(self._close_btn)
        vl.addLayout(ctrl_row)

    def _set_status(self, text: str, level: str) -> None:
        message = str(text).strip() or "No backend status available"
        if self._backend.simulated and level in {"neutral", "ready", "active"}:
            level = "simulated"
            if "simulat" not in message.lower():
                message = f"{message} (no-communication simulation; not hardware evidence)"
        set_semantic_status(
            self._status_lbl,
            message,
            level,
            accessible_name="IRM and ARM backend status",
        )

    def _on_mode_changed(self, idx: int) -> None:
        self._irm_grp.setVisible(idx == 0)
        self._arm_grp.setVisible(idx == 1)

    def _apply(self) -> None:
        if self._task is not None:
            return
        if self._mode.currentIndex()==0 and self._manual_arm is not None:
            sample,field,axis = self._sample_id.text(),float(self._irm_field.value()),self._irm_axis.currentText()
            self._start_task(lambda:self._manual_arm.apply_irm(sample_id=sample,max_field_mT=field,axis=axis,should_cancel=self._cancel.is_set))
            return
        if self._mode.currentIndex() == 1 and self._manual_arm is not None:
            sample = self._sample_id.text()
            peak, bias = float(self._arm_peak_af.value()), float(self._arm_bias.value())
            self._start_task(lambda: self._manual_arm.apply(sample_id=sample, peak_af_mT=peak,
                             bias_mT=bias, should_cancel=self._cancel.is_set))
            return
        try:
            if self._mode.currentIndex() == 0:
                msg = self._backend.apply_irm(
                    max_field_mT=float(self._irm_field.value()),
                    axis=self._irm_axis.currentText(),
                    ramp_label=self._irm_ramp.currentText(),
                    steps=10,
                )
            else:
                msg = self._backend.apply_arm(
                    peak_af_mT=float(self._arm_peak_af.value()),
                    bias_mT=float(self._arm_bias.value()),
                )
        except Exception as exc:
            msg = str(exc)
            level = "error"
        else:
            level = "ready"
        self._set_status(msg, level)

    def _reset(self) -> None:
        if self._task is not None:
            return
        if self._manual_arm is not None:
            self._start_task(self._manual_arm.reset)
            return
        try:
            msg = self._backend.reset_field()
        except Exception as exc:
            msg = str(exc)
            level = "error"
        else:
            level = "ready"
        self._set_status(msg, level)

    def _start_task(self, operation):
        self._cancel.clear()
        self._apply_btn.setEnabled(False)
        self._reset_btn.setEnabled(False)
        self._mode.setEnabled(False)
        self._arm_grp.setEnabled(False)
        self._irm_grp.setEnabled(False)
        self._sample_row.setEnabled(False)
        self._cancel_btn.setEnabled(True)
        self._set_status("Treatment running. Cancel waits for field cleanup and safe return.", "active")
        self._task = _TreatmentTask(operation, self)
        self._task.result.connect(self._task_result)
        self._task.finished.connect(self._task_finished)
        self._task.start()

    @QtCore.Slot(str, bool)
    def _task_result(self, message, ok):
        self._set_status(message, "ready" if ok else "error")

    def _task_finished(self):
        task, self._task = self._task, None
        task.deleteLater()
        self._apply_btn.setEnabled(True)
        self._reset_btn.setEnabled(True)
        self._mode.setEnabled(True)
        self._arm_grp.setEnabled(True)
        self._irm_grp.setEnabled(True)
        self._sample_row.setEnabled(True)
        self._cancel_btn.setEnabled(False)
        if self._close_pending:
            super().reject()

    def _cancel_treatment(self):
        self._cancel.set()
        self._set_status("Cancellation requested; waiting for cleanup and evidence publication.", "active")

    def reject(self):
        if self._task is not None:
            self._close_pending = True
            self._cancel_treatment()
            return
        super().reject()

    def accept(self):
        if self._task is not None:
            self._close_pending = True
            self._cancel_treatment()
            return
        super().accept()

    def closeEvent(self, event):
        if self._task is not None:
            self._close_pending = True
            self._cancel_treatment()
            event.ignore()
            return
        super().closeEvent(event)
