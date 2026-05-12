from __future__ import annotations

from PySide6 import QtCore, QtWidgets


class SquidCommDialog(QtWidgets.QDialog):
    """SQUID serial communication settings & test — replaces VB6 frmSquid."""

    def __init__(self, parent: QtWidgets.QWidget | None = None) -> None:
        super().__init__(parent)
        self.setWindowTitle("SQUID Communication Settings")
        self.setMinimumWidth(420)
        self.setWindowFlags(self.windowFlags() & ~QtCore.Qt.WindowContextHelpButtonHint)
        self._build_ui()

    # ── UI ─────────────────────────────────────────────────────────────────
    def _build_ui(self) -> None:
        vl = QtWidgets.QVBoxLayout(self)
        vl.setContentsMargins(20, 16, 20, 16)
        vl.setSpacing(14)

        hdr = QtWidgets.QLabel("SQUID Magnetometer — Serial Port")
        hdr.setStyleSheet("font-size: 14px; font-weight: 700; color: #7A0219;")
        vl.addWidget(hdr)

        # ── Port settings ────────────────────────────────────────────────
        grp_port = QtWidgets.QGroupBox("Connection")
        fl = QtWidgets.QFormLayout(grp_port)
        fl.setSpacing(8)
        fl.setLabelAlignment(QtCore.Qt.AlignRight)

        self._port = QtWidgets.QComboBox()
        self._port.setEditable(True)
        self._port.addItems([f"COM{i}" for i in range(1, 9)])
        fl.addRow("Serial port:", self._port)

        self._baud = QtWidgets.QComboBox()
        self._baud.addItems(["1200", "2400", "4800", "9600", "19200"])
        self._baud.setCurrentText("9600")
        fl.addRow("Baud rate:", self._baud)

        self._parity = QtWidgets.QComboBox()
        self._parity.addItems(["None", "Even", "Odd"])
        fl.addRow("Parity:", self._parity)

        vl.addWidget(grp_port)

        # ── Measurement settings ─────────────────────────────────────────
        grp_meas = QtWidgets.QGroupBox("Measurement")
        fl2 = QtWidgets.QFormLayout(grp_meas)
        fl2.setSpacing(8)
        fl2.setLabelAlignment(QtCore.Qt.AlignRight)

        self._range = QtWidgets.QComboBox()
        self._range.addItems(["1×", "10×", "100×", "1000×"])
        fl2.addRow("Sensitivity range:", self._range)

        self._samples = QtWidgets.QSpinBox()
        self._samples.setRange(1, 64)
        self._samples.setValue(8)
        fl2.addRow("Samples per position:", self._samples)

        self._settle = QtWidgets.QDoubleSpinBox()
        self._settle.setRange(0.1, 10.0)
        self._settle.setValue(1.0)
        self._settle.setSuffix(" s")
        self._settle.setSingleStep(0.1)
        fl2.addRow("Settle time:", self._settle)

        vl.addWidget(grp_meas)

        # ── Test connection ──────────────────────────────────────────────
        test_row = QtWidgets.QHBoxLayout()
        self._test_btn = QtWidgets.QPushButton("Test Connection")
        self._test_btn.clicked.connect(self._test_connection)
        self._status_lbl = QtWidgets.QLabel("Not tested")
        self._status_lbl.setStyleSheet("color: #9a8885; font-size: 11px;")
        test_row.addWidget(self._test_btn)
        test_row.addWidget(self._status_lbl, 1)
        vl.addLayout(test_row)

        # ── Buttons ──────────────────────────────────────────────────────
        btns = QtWidgets.QDialogButtonBox(
            QtWidgets.QDialogButtonBox.Ok | QtWidgets.QDialogButtonBox.Cancel
        )
        btns.accepted.connect(self.accept)
        btns.rejected.connect(self.reject)
        vl.addWidget(btns)

    def _test_connection(self) -> None:
        """Stub — Phase 3 will open the serial port and send a query."""
        self._status_lbl.setText("⚠  Hardware not connected (Phase 3)")
        self._status_lbl.setStyleSheet("color: #b45309; font-size: 11px;")
