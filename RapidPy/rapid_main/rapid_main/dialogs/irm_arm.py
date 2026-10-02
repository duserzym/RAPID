from __future__ import annotations

from PySide6 import QtCore, QtWidgets

from rapid_main.diagnostic_services import IrmArmBackend, IrmArmNoCommBackend
from rapid_main.glass_theme import set_semantic_status


class IrmArmDialog(QtWidgets.QDialog):
    """IRM / ARM control dialog — replaces VB6 frmIRMARM."""

    def __init__(
        self,
        parent: QtWidgets.QWidget | None = None,
        backend: IrmArmBackend | None = None,
    ) -> None:
        super().__init__(parent)
        self.setObjectName("glassDialog")
        self.setWindowTitle("IRM / ARM Control")
        self.setAccessibleName("IRM and ARM field control")
        self.setMinimumWidth(340)
        self.setWindowFlags(self.windowFlags() & ~QtCore.Qt.WindowContextHelpButtonHint)
        self._backend = backend or IrmArmNoCommBackend()
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

        # ── IRM settings ─────────────────────────────────────────────────
        self._irm_grp = QtWidgets.QGroupBox("IRM Settings")
        fl = QtWidgets.QFormLayout(self._irm_grp)
        fl.setSpacing(8)
        fl.setLabelAlignment(QtCore.Qt.AlignRight)
        fl.setRowWrapPolicy(QtWidgets.QFormLayout.RowWrapPolicy.WrapLongRows)

        self._irm_field = QtWidgets.QDoubleSpinBox()
        self._irm_field.setRange(0, 2000)
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

        self._apply_btn.setEnabled(connected)
        self._reset_btn.setEnabled(connected)

        ctrl_row.addWidget(self._apply_btn)
        ctrl_row.addWidget(self._reset_btn)
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
        try:
            msg = self._backend.reset_field()
        except Exception as exc:
            msg = str(exc)
            level = "error"
        else:
            level = "ready"
        self._set_status(msg, level)
