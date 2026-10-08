from __future__ import annotations

from types import SimpleNamespace

from PySide6 import QtCore, QtWidgets

from rapid_main.config import normalize_squid_range_label
from rapid_main.diagnostic_services import SquidBackend, SquidNoCommBackend
from rapid_main.glass_theme import set_semantic_status

#: Read-rate / motion-capture fields edited here and applied on Save Settings.
STREAM_FIELDS = (
    ("stream_interval_ms", 500),
    ("stream_axes", "XYZ"),
    ("stream_counts_every", 1),
    ("stream_latch_count_hold_ms", 100),
    ("stream_latch_data_hold_ms", 120),
)


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
        self._baud.setCurrentText("1200")  # VB6 frmSQUID: 1200,N,8,1
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
        self._samples.setValue(1)
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

        grp_stream = QtWidgets.QGroupBox("Continuous read")
        stream_layout = QtWidgets.QVBoxLayout(grp_stream)
        stream_layout.setSpacing(8)
        self._stream_settings = SimpleNamespace(baud=9600, **dict(STREAM_FIELDS))
        self._stream_summary = QtWidgets.QLabel()
        self._stream_summary.setObjectName("guidanceText")
        self._stream_summary.setWordWrap(True)
        stream_layout.addWidget(self._stream_summary)
        self._stream_btn = QtWidgets.QPushButton("Live Stream…")
        self._stream_btn.setAccessibleName("Open live SQUID stream with read-rate control")
        self._stream_btn.setToolTip("Set the SQUID read rate, watch the live signal, filter it and view its spectrum.")
        self._stream_btn.clicked.connect(self._open_stream)
        stream_layout.addWidget(self._stream_btn)
        self._capture_enabled = QtWidgets.QCheckBox("Capture while moving")
        self._capture_enabled.setToolTip(
            "Opt-in. Streams the SQUID at the read rate above during the borehole descent,\n"
            "the 90° turns and the ascent of every measurement block. Saved as auxiliary\n"
            "CSV + analysis files; never used for the published measurement."
        )
        self._capture_enabled.setAccessibleName("Enable SQUID motion capture")
        stream_layout.addWidget(self._capture_enabled)
        segment_row = QtWidgets.QHBoxLayout()
        self._capture_descent = QtWidgets.QCheckBox("Descent")
        self._capture_turns = QtWidgets.QCheckBox("Turns")
        self._capture_ascent = QtWidgets.QCheckBox("Ascent")
        for box in (self._capture_descent, self._capture_turns, self._capture_ascent):
            box.setChecked(True)
            segment_row.addWidget(box)
        segment_row.addStretch(1)
        stream_layout.addLayout(segment_row)
        self._capture_dir = QtWidgets.QLineEdit()
        self._capture_dir.setPlaceholderText("Default: <data folder>/squid_motion_capture")
        self._capture_dir.setAccessibleName("Motion capture output folder")
        stream_layout.addWidget(self._capture_dir)
        self._capture_enabled.toggled.connect(self._update_capture_controls)
        self._update_stream_summary()
        self._update_capture_controls()
        settings_layout.addWidget(grp_stream)
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
        self._range.setCurrentText(normalize_squid_range_label(getattr(cfg, "range_label", self._range.currentText())))
        self._samples.setValue(int(getattr(cfg, "samples_per_pos", self._samples.value())))
        self._settle.setValue(float(getattr(cfg, "settle_time", self._settle.value())))
        for name, default in STREAM_FIELDS:
            setattr(self._stream_settings, name, getattr(cfg, name, default))
        self._capture_enabled.setChecked(bool(getattr(cfg, "motion_capture_enabled", False)))
        self._capture_descent.setChecked(bool(getattr(cfg, "motion_capture_descent", True)))
        self._capture_turns.setChecked(bool(getattr(cfg, "motion_capture_turns", True)))
        self._capture_ascent.setChecked(bool(getattr(cfg, "motion_capture_ascent", True)))
        self._capture_dir.setText(str(getattr(cfg, "motion_capture_dir", "") or ""))
        self._update_stream_summary()
        self._update_capture_controls()

    def _accept_settings(self) -> None:
        cfg = self._settings_config()
        if cfg is not None:
            cfg.port = self._port.currentText().strip()
            cfg.baud = int(self._baud.currentText())
            cfg.range_label = self._range.currentText()
            cfg.samples_per_pos = int(self._samples.value())
            cfg.settle_time = float(self._settle.value())
            for name, _default in STREAM_FIELDS:
                setattr(cfg, name, getattr(self._stream_settings, name))
            cfg.motion_capture_enabled = bool(self._capture_enabled.isChecked())
            cfg.motion_capture_descent = bool(self._capture_descent.isChecked())
            cfg.motion_capture_turns = bool(self._capture_turns.isChecked())
            cfg.motion_capture_ascent = bool(self._capture_ascent.isChecked())
            cfg.motion_capture_dir = self._capture_dir.text().strip()

        parent = self.parentWidget()
        app_config = getattr(parent, "config", None)
        if app_config is not None and hasattr(app_config, "save"):
            try:
                app_config.save()
            except Exception as exc:
                self._set_status(f"Unable to save SQUID settings: {exc}", "error")
                return
        self.accept()

    def _update_capture_controls(self) -> None:
        enabled = self._capture_enabled.isChecked()
        for widget in (self._capture_descent, self._capture_turns, self._capture_ascent, self._capture_dir):
            widget.setEnabled(enabled)

    def _update_stream_summary(self) -> None:
        settings = self._stream_settings
        interval = int(settings.stream_interval_ms)
        rate = "as fast as the link allows" if interval <= 0 else f"every {interval} ms"
        self._stream_summary.setText(
            f"Read rate: {rate}, axes {settings.stream_axes}, flux counter every "
            f"{settings.stream_counts_every} sample(s). Static bracketed reads keep the VB6 timing."
        )

    def _raw_client(self):
        connect = getattr(self._backend, "connect_for_acquisition", None)
        if callable(connect):
            connect()
        return getattr(self._backend, "raw_client", None)

    def _open_stream(self) -> None:
        from rapid_main.dialogs.squid_stream import SquidStreamDialog

        try:
            self._stream_settings.baud = int(self._baud.currentText())
        except ValueError:
            self._stream_settings.baud = 1200
        dialog = SquidStreamDialog(
            self,
            client_provider=self._raw_client,
            squid_config=self._stream_settings,
            simulated=bool(self._backend.simulated),
        )
        dialog.exec()
        self._update_stream_summary()

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
