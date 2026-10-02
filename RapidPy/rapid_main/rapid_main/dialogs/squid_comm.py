from __future__ import annotations

from PySide6 import QtCore, QtWidgets

from rapid_main.diagnostic_services import SquidBackend, SquidNoCommBackend
from rapid_main.glass_theme import set_semantic_status


class SquidCommDialog(QtWidgets.QDialog):
    """SQUID serial communication settings & test — replaces VB6 frmSquid."""

    def __init__(
        self,
        parent: QtWidgets.QWidget | None = None,
        backend: SquidBackend | None = None,
    ) -> None:
        super().__init__(parent)
        self._backend: SquidBackend = backend or SquidNoCommBackend()
        self.setObjectName("glassDialog")
        self.setWindowTitle("SQUID Communication Settings")
        self.setAccessibleName("SQUID communication settings")
        self.setMinimumWidth(340)
        self.setWindowFlags(self.windowFlags() & ~QtCore.Qt.WindowContextHelpButtonHint)
        self._build_ui()

    def _build_ui(self) -> None:
        vl = QtWidgets.QVBoxLayout(self)
        vl.setContentsMargins(20, 16, 20, 16)
        vl.setSpacing(14)

        hdr = QtWidgets.QLabel("SQUID Magnetometer — Serial Port")
        hdr.setObjectName("dialogTitle")
        hdr.setAccessibleName("SQUID communication settings title")
        vl.addWidget(hdr)

        self._settings_scroll = QtWidgets.QScrollArea()
        self._settings_scroll.setObjectName("dialogScroll")
        self._settings_scroll.setWidgetResizable(True)
        self._settings_scroll.setFrameShape(QtWidgets.QFrame.Shape.NoFrame)
        self._settings_scroll.setHorizontalScrollBarPolicy(
            QtCore.Qt.ScrollBarPolicy.ScrollBarAlwaysOff
        )
        self._settings_scroll.setAccessibleName("SQUID connection and measurement settings")
        settings_host = QtWidgets.QWidget()
        settings_layout = QtWidgets.QVBoxLayout(settings_host)
        settings_layout.setContentsMargins(0, 0, 0, 0)
        settings_layout.setSpacing(12)

        grp_port = QtWidgets.QGroupBox("Connection")
        fl = QtWidgets.QFormLayout(grp_port)
        fl.setSpacing(8)
        fl.setLabelAlignment(QtCore.Qt.AlignRight)
        fl.setRowWrapPolicy(QtWidgets.QFormLayout.RowWrapPolicy.WrapLongRows)

        self._port = QtWidgets.QComboBox()
        self._port.setEditable(True)
        self._port.addItems([f"COM{i}" for i in range(1, 9)])
        self._port.setAccessibleName("SQUID serial port")
        fl.addRow("Serial port:", self._port)

        self._baud = QtWidgets.QComboBox()
        self._baud.addItems(["1200", "2400", "4800", "9600", "19200"])
        self._baud.setCurrentText("9600")
        self._baud.setAccessibleName("SQUID serial baud rate")
        fl.addRow("Baud rate:", self._baud)

        self._parity = QtWidgets.QComboBox()
        self._parity.addItems(["None", "Even", "Odd"])
        self._parity.setCurrentText("None")
        self._parity.setEnabled(False)
        self._parity.setAccessibleName("SQUID serial parity fixed to none")
        self._parity.setToolTip("The current SQUID transport uses fixed no-parity framing.")
        fl.addRow("Parity:", self._parity)
        settings_layout.addWidget(grp_port)

        grp_meas = QtWidgets.QGroupBox("Measurement")
        fl2 = QtWidgets.QFormLayout(grp_meas)
        fl2.setSpacing(8)
        fl2.setLabelAlignment(QtCore.Qt.AlignRight)
        fl2.setRowWrapPolicy(QtWidgets.QFormLayout.RowWrapPolicy.WrapLongRows)

        self._range = QtWidgets.QComboBox()
        self._range.addItems(["1×", "10×", "100×", "1000×"])
        self._range.setAccessibleName("SQUID sensitivity range")
        fl2.addRow("Sensitivity range:", self._range)

        self._samples = QtWidgets.QSpinBox()
        self._samples.setRange(1, 64)
        self._samples.setValue(8)
        self._samples.setAccessibleName("SQUID samples per position")
        fl2.addRow("Samples per position:", self._samples)

        self._settle = QtWidgets.QDoubleSpinBox()
        self._settle.setRange(0.1, 10.0)
        self._settle.setValue(1.0)
        self._settle.setSuffix(" s")
        self._settle.setSingleStep(0.1)
        self._settle.setAccessibleName("SQUID settling time in seconds")
        fl2.addRow("Settle time:", self._settle)

        settings_layout.addWidget(grp_meas)
        settings_layout.addStretch()
        self._settings_scroll.setWidget(settings_host)
        vl.addWidget(self._settings_scroll, 1)

        test_row = QtWidgets.QHBoxLayout()
        self._test_btn = QtWidgets.QPushButton("Test Connection")
        self._test_btn.setAccessibleName("Test SQUID connection")
        self._test_btn.clicked.connect(self._test_connection)
        self._status_lbl = QtWidgets.QLabel()
        self._status_lbl.setWordWrap(True)
        self._load_backend_settings()
        connected = False
        try:
            connected = bool(self._backend.is_connected())
        except Exception:
            pass
        try:
            backend_status = self._backend.status()
        except Exception as exc:
            backend_status = f"SQUID status read failed: {exc}"
        self._set_status(
            backend_status,
            "ready"
            if connected
            else ("error" if "failed" in backend_status else "unavailable"),
        )
        test_row.addWidget(self._test_btn)
        test_row.addWidget(self._status_lbl, 1)
        vl.addLayout(test_row)

        self._button_box = QtWidgets.QDialogButtonBox(
            QtWidgets.QDialogButtonBox.Ok | QtWidgets.QDialogButtonBox.Cancel
        )
        self._button_box.setAccessibleName("SQUID settings actions")
        self._save_btn = self._button_box.button(QtWidgets.QDialogButtonBox.StandardButton.Ok)
        self._cancel_btn = self._button_box.button(
            QtWidgets.QDialogButtonBox.StandardButton.Cancel
        )
        if self._save_btn is not None:
            self._save_btn.setText("Save Settings")
            self._save_btn.setAccessibleName("Save SQUID settings")
        if self._cancel_btn is not None:
            self._cancel_btn.setAccessibleName("Cancel SQUID settings changes")
        self._button_box.accepted.connect(self._accept_settings)
        self._button_box.rejected.connect(self.reject)
        vl.addWidget(self._button_box)

    def _settings_config(self) -> object | None:
        parent = self.parentWidget()
        app_config = getattr(parent, "config", None)
        parent_squid = getattr(app_config, "squid", None)
        return parent_squid if parent_squid is not None else getattr(self._backend, "_cfg", None)

    def _load_backend_settings(self) -> None:
        cfg = self._settings_config()
        if cfg is None:
            return
        self._port.setCurrentText(str(getattr(cfg, "port", self._port.currentText())))
        self._baud.setCurrentText(str(getattr(cfg, "baud", self._baud.currentText())))
        self._range.setCurrentText(str(getattr(cfg, "range_label", self._range.currentText())))
        self._samples.setValue(int(getattr(cfg, "samples_per_pos", self._samples.value())))
        self._settle.setValue(float(getattr(cfg, "settle_time", self._settle.value())))

    def _accept_settings(self) -> None:
        cfg = self._settings_config()
        if cfg is not None:
            cfg.port = self._port.currentText().strip()
            cfg.baud = int(self._baud.currentText())
            cfg.range_label = self._range.currentText()
            cfg.samples_per_pos = int(self._samples.value())
            cfg.settle_time = float(self._settle.value())

        parent = self.parentWidget()
        app_config = getattr(parent, "config", None)
        if app_config is not None and hasattr(app_config, "save"):
            try:
                app_config.save()
            except Exception as exc:
                self._set_status(f"Unable to save SQUID settings: {exc}", "error")
                return
        self.accept()

    def _annotate_status(self, text: str) -> str:
        status = text.strip()
        if self._backend.simulated and "sim" not in status.lower():
            status = f"{status} (simulated)"
        return status

    def _set_status(self, text: str, level: str) -> None:
        if self._backend.simulated and level in {"neutral", "ready", "active", "unavailable"}:
            level = "simulated"
        set_semantic_status(
            self._status_lbl,
            self._annotate_status(text),
            level,
            accessible_name="SQUID connection status",
        )

    def _test_connection(self) -> None:
        try:
            connected = bool(self._backend.test_connection())
        except Exception as exc:
            self._set_status(f"Connection failed: {exc}", "error")
            return

        if connected:
            self._set_status(f"Connected: {self._backend.status()}", "ready")
        else:
            self._set_status("SQUID backend is not connected", "unavailable")
