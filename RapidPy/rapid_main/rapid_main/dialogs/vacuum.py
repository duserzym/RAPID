from __future__ import annotations

from PySide6 import QtCore, QtWidgets


class VacuumDialog(QtWidgets.QDialog):
    """Vacuum pressure monitor — replaces VB6 frmVacuum."""

    def __init__(self, parent: QtWidgets.QWidget | None = None) -> None:
        super().__init__(parent)
        self.setWindowTitle("Vacuum Monitor")
        self.setMinimumWidth(380)
        self.setWindowFlags(self.windowFlags() & ~QtCore.Qt.WindowContextHelpButtonHint)
        self._build_ui()

        # Simulated live refresh timer (Phase 3 will poll real hardware)
        self._timer = QtCore.QTimer(self)
        self._timer.timeout.connect(self._refresh)
        self._timer.start(2000)

    # ── Public API ─────────────────────────────────────────────────────────
    def set_pressure(self, mtorr: float) -> None:
        self._pressure_lbl.setText(f"{mtorr:.3f}")
        warn = mtorr > float(self._warn_spin.value())
        self._pressure_lbl.setStyleSheet(
            f"font-size: 36px; font-weight: 800; "
            f"color: {'#b91c1c' if warn else '#15803d'};"
        )
        self._status_lbl.setText("⚠ Above threshold" if warn else "✓ Within range")
        self._status_lbl.setStyleSheet(
            f"color: {'#b91c1c' if warn else '#15803d'}; font-size: 12px;"
        )

    # ── UI ─────────────────────────────────────────────────────────────────
    def _build_ui(self) -> None:
        vl = QtWidgets.QVBoxLayout(self)
        vl.setContentsMargins(20, 16, 20, 16)
        vl.setSpacing(14)

        hdr = QtWidgets.QLabel("Vacuum Pressure")
        hdr.setStyleSheet("font-size: 14px; font-weight: 700; color: #7A0219;")
        vl.addWidget(hdr)

        # Live pressure readout
        read_frame = QtWidgets.QFrame()
        read_frame.setStyleSheet(
            "QFrame { background: rgba(122,2,25,0.04); border: 1px solid rgba(122,2,25,0.12);"
            " border-radius: 10px; }"
        )
        rl = QtWidgets.QVBoxLayout(read_frame)
        rl.setAlignment(QtCore.Qt.AlignCenter)
        rl.setContentsMargins(16, 14, 16, 14)

        self._pressure_lbl = QtWidgets.QLabel("—")
        self._pressure_lbl.setAlignment(QtCore.Qt.AlignCenter)
        self._pressure_lbl.setStyleSheet("font-size: 36px; font-weight: 800; color: #6b7280;")

        unit_lbl = QtWidgets.QLabel("mTorr")
        unit_lbl.setAlignment(QtCore.Qt.AlignCenter)
        unit_lbl.setStyleSheet("color: #9a8885; font-size: 13px;")

        self._status_lbl = QtWidgets.QLabel("Not connected (Phase 3)")
        self._status_lbl.setAlignment(QtCore.Qt.AlignCenter)
        self._status_lbl.setStyleSheet("color: #6b7280; font-size: 12px;")

        rl.addWidget(self._pressure_lbl)
        rl.addWidget(unit_lbl)
        rl.addWidget(self._status_lbl)
        vl.addWidget(read_frame)

        # Settings
        grp = QtWidgets.QGroupBox("Threshold")
        fl = QtWidgets.QFormLayout(grp)
        fl.setSpacing(8)
        fl.setLabelAlignment(QtCore.Qt.AlignRight)

        self._target_spin = QtWidgets.QDoubleSpinBox()
        self._target_spin.setRange(0.0, 100.0)
        self._target_spin.setValue(5.0)
        self._target_spin.setSuffix(" mTorr")
        fl.addRow("Target pressure:", self._target_spin)

        self._warn_spin = QtWidgets.QDoubleSpinBox()
        self._warn_spin.setRange(0.0, 200.0)
        self._warn_spin.setValue(20.0)
        self._warn_spin.setSuffix(" mTorr")
        fl.addRow("Warning threshold:", self._warn_spin)

        vl.addWidget(grp)

        # Pump control
        pump_row = QtWidgets.QHBoxLayout()
        self._pump_btn = QtWidgets.QPushButton("⏻  Pump On")
        self._pump_btn.setCheckable(True)
        self._pump_btn.toggled.connect(self._on_pump_toggle)
        pump_row.addWidget(self._pump_btn)
        pump_row.addStretch()
        vl.addLayout(pump_row)

        # Close button
        close_btn = QtWidgets.QPushButton("Close")
        close_btn.clicked.connect(self.close)
        btn_row = QtWidgets.QHBoxLayout()
        btn_row.addStretch()
        btn_row.addWidget(close_btn)
        vl.addLayout(btn_row)

    def _on_pump_toggle(self, on: bool) -> None:
        self._pump_btn.setText("⏹  Pump Off" if on else "⏻  Pump On")

    def _refresh(self) -> None:
        """Phase 3: poll hardware. For now display stub."""
        pass

    def closeEvent(self, event: "QtCore.QEvent") -> None:  # type: ignore[override]
        self._timer.stop()
        super().closeEvent(event)
