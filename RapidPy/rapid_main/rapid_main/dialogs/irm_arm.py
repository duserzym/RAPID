from __future__ import annotations

from PySide6 import QtCore, QtWidgets


class IrmArmDialog(QtWidgets.QDialog):
    """IRM / ARM control dialog — replaces VB6 frmIRMARM."""

    def __init__(self, parent: QtWidgets.QWidget | None = None) -> None:
        super().__init__(parent)
        self.setWindowTitle("IRM / ARM Control")
        self.setMinimumWidth(440)
        self.setWindowFlags(self.windowFlags() & ~QtCore.Qt.WindowContextHelpButtonHint)
        self._build_ui()

    # ── UI ─────────────────────────────────────────────────────────────────
    def _build_ui(self) -> None:
        vl = QtWidgets.QVBoxLayout(self)
        vl.setContentsMargins(20, 16, 20, 16)
        vl.setSpacing(12)

        hdr = QtWidgets.QLabel("IRM / ARM Control")
        hdr.setStyleSheet("font-size: 14px; font-weight: 700; color: #7A0219;")
        vl.addWidget(hdr)

        # ── Mode selector ────────────────────────────────────────────────
        mode_row = QtWidgets.QHBoxLayout()
        mode_lbl = QtWidgets.QLabel("Mode:")
        mode_lbl.setStyleSheet("color: #9a8885;")
        self._mode = QtWidgets.QComboBox()
        self._mode.addItems(["IRM (Isothermal Remanence)", "ARM (Anhysteretic Remanence)"])
        self._mode.currentIndexChanged.connect(self._on_mode_changed)
        mode_row.addWidget(mode_lbl)
        mode_row.addWidget(self._mode, 1)
        vl.addLayout(mode_row)

        # ── IRM settings ─────────────────────────────────────────────────
        self._irm_grp = QtWidgets.QGroupBox("IRM Settings")
        fl = QtWidgets.QFormLayout(self._irm_grp)
        fl.setSpacing(8)
        fl.setLabelAlignment(QtCore.Qt.AlignRight)

        self._irm_field = QtWidgets.QDoubleSpinBox()
        self._irm_field.setRange(0, 2000)
        self._irm_field.setValue(100)
        self._irm_field.setSuffix(" mT")
        self._irm_field.setSingleStep(10)
        fl.addRow("Peak DC field:", self._irm_field)

        self._irm_axis = QtWidgets.QComboBox()
        self._irm_axis.addItems(["Z (up-axis)", "X", "Y"])
        fl.addRow("Magnetise axis:", self._irm_axis)

        self._irm_ramp = QtWidgets.QComboBox()
        self._irm_ramp.addItems(["Slow (60 s)", "Medium (30 s)", "Fast (10 s)"])
        fl.addRow("Ramp speed:", self._irm_ramp)

        vl.addWidget(self._irm_grp)

        # ── ARM settings ─────────────────────────────────────────────────
        self._arm_grp = QtWidgets.QGroupBox("ARM Settings")
        fl2 = QtWidgets.QFormLayout(self._arm_grp)
        fl2.setSpacing(8)
        fl2.setLabelAlignment(QtCore.Qt.AlignRight)

        self._arm_peak_af = QtWidgets.QDoubleSpinBox()
        self._arm_peak_af_spin = self._arm_peak_af
        self._arm_peak_af.setRange(0, 200)
        self._arm_peak_af.setValue(80)
        self._arm_peak_af.setSuffix(" mT")
        fl2.addRow("Peak AF field:", self._arm_peak_af)

        self._arm_bias = QtWidgets.QDoubleSpinBox()
        self._arm_bias.setRange(0, 100)
        self._arm_bias.setValue(0.05)
        self._arm_bias.setDecimals(3)
        self._arm_bias.setSingleStep(0.005)
        self._arm_bias.setSuffix(" mT")
        fl2.addRow("Bias field:", self._arm_bias)

        vl.addWidget(self._arm_grp)
        self._arm_grp.setVisible(False)

        # ── Status ───────────────────────────────────────────────────────
        self._status_lbl = QtWidgets.QLabel("Ready — no hardware connected (Phase 3)")
        self._status_lbl.setStyleSheet("color: #9a8885; font-size: 11px;")
        vl.addWidget(self._status_lbl)

        # ── Control buttons ──────────────────────────────────────────────
        ctrl_row = QtWidgets.QHBoxLayout()
        self._apply_btn = QtWidgets.QPushButton("Apply Field")
        self._apply_btn.setObjectName("accent")
        self._apply_btn.clicked.connect(self._apply)

        self._reset_btn = QtWidgets.QPushButton("Reset to Zero")
        self._reset_btn.clicked.connect(self._reset)

        close_btn = QtWidgets.QPushButton("Close")
        close_btn.clicked.connect(self.close)

        ctrl_row.addWidget(self._apply_btn)
        ctrl_row.addWidget(self._reset_btn)
        ctrl_row.addStretch()
        ctrl_row.addWidget(close_btn)
        vl.addLayout(ctrl_row)

    def _on_mode_changed(self, idx: int) -> None:
        self._irm_grp.setVisible(idx == 0)
        self._arm_grp.setVisible(idx == 1)

    def _apply(self) -> None:
        """Phase 3: send ramp command to hardware."""
        self._status_lbl.setText("⚠  Hardware not connected (Phase 3)")
        self._status_lbl.setStyleSheet("color: #b45309; font-size: 11px;")

    def _reset(self) -> None:
        """Phase 3: ramp field down to zero."""
        self._status_lbl.setText("⚠  Hardware not connected (Phase 3)")
        self._status_lbl.setStyleSheet("color: #b45309; font-size: 11px;")
