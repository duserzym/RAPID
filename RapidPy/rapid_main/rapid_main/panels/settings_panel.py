from __future__ import annotations

from PySide6 import QtCore, QtWidgets

from rapid_main.config import AppConfig, SequenceTimesConfig


def _sec_hdr(text: str) -> QtWidgets.QLabel:
    lbl = QtWidgets.QLabel(text)
    lbl.setStyleSheet(
        "font-size: 10px; font-weight: 700; color: #9a8885; letter-spacing: 1.2px;"
        " margin-top: 6px;"
    )
    return lbl


class SettingsPanel(QtWidgets.QWidget):
    """System settings panel — tabbed interface.

    Maps to: frmSettings_new in VB6.
    Tabs mirror the VB6 frameOptions index array (8 pages).
    Phase 2: all 8 tabs fully populated.
    """

    def __init__(self, parent: QtWidgets.QWidget | None = None) -> None:
        super().__init__(parent)
        root = QtWidgets.QVBoxLayout(self)
        root.setContentsMargins(16, 12, 16, 16)
        root.setSpacing(10)

        hdr = QtWidgets.QLabel("SYSTEM SETTINGS")
        hdr.setObjectName("sectionHdr")
        root.addWidget(hdr)

        self._tabs = QtWidgets.QTabWidget()
        root.addWidget(self._tabs)

        self._tabs.addTab(self._build_general_tab(),    "General")
        self._tabs.addTab(self._build_squid_tab(),      "SQUID")
        self._tabs.addTab(self._build_irm_arm_tab(),    "IRM / ARM")
        self._tabs.addTab(self._build_af_demag_tab(),   "AF Demag")
        self._tabs.addTab(self._build_vacuum_tab(),     "Vacuum")
        self._tabs.addTab(self._build_data_files_tab(), "Data Files")
        self._tabs.addTab(self._build_changer_tab(),    "Changer")
        self._tabs.addTab(self._build_calibration_tab(), "Calibration")
        self._tabs.addTab(self._build_sequence_tab(),   "Sequence")

        bottom = QtWidgets.QHBoxLayout()
        bottom.addStretch()
        save_btn = QtWidgets.QPushButton("💾  Save Settings")
        save_btn.setObjectName("accent")
        save_btn.clicked.connect(self._save)
        cancel_btn = QtWidgets.QPushButton("Cancel")
        cancel_btn.clicked.connect(self._revert)
        bottom.addWidget(cancel_btn)
        bottom.addWidget(save_btn)
        root.addLayout(bottom)

    # ── General tab ───────────────────────────────────────────────────────────
    def _build_general_tab(self) -> QtWidgets.QWidget:
        w = QtWidgets.QWidget()
        scroll = QtWidgets.QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QtWidgets.QFrame.NoFrame)

        inner = QtWidgets.QWidget()
        fl = QtWidgets.QFormLayout(inner)
        fl.setContentsMargins(18, 16, 18, 16)
        fl.setSpacing(10)
        fl.setLabelAlignment(QtCore.Qt.AlignRight)

        def _path_row(label: str) -> QtWidgets.QHBoxLayout:
            rl = QtWidgets.QHBoxLayout()
            edit = QtWidgets.QLineEdit()
            edit.setPlaceholderText("Browse or type path…")
            browse = QtWidgets.QPushButton("Browse…")
            browse.setFixedWidth(80)
            rl.addWidget(edit)
            rl.addWidget(browse)
            fl.addRow(label, rl)
            return edit

        self._data_dir    = _path_row("Data folder:")
        self._sample_dir  = _path_row("Sample index folder:")
        self._backup_dir  = _path_row("Backup folder:")

        self._operator    = QtWidgets.QLineEdit()
        self._operator.setPlaceholderText("Operator name")
        fl.addRow("Operator:", self._operator)

        self._lab_name    = QtWidgets.QLineEdit()
        self._lab_name.setPlaceholderText("Laboratory name")
        fl.addRow("Laboratory:", self._lab_name)

        self._nocomm_chk  = QtWidgets.QCheckBox("Start in No-Comm mode")
        fl.addRow("", self._nocomm_chk)

        self._auto_save   = QtWidgets.QCheckBox("Auto-save after each measurement")
        self._auto_save.setChecked(True)
        fl.addRow("", self._auto_save)

        scroll.setWidget(inner)
        vl = QtWidgets.QVBoxLayout(w)
        vl.setContentsMargins(0, 0, 0, 0)
        vl.addWidget(scroll)
        return w

    # ── SQUID tab ─────────────────────────────────────────────────────────────
    def _build_squid_tab(self) -> QtWidgets.QWidget:
        w = QtWidgets.QWidget()
        fl = QtWidgets.QFormLayout(w)
        fl.setContentsMargins(18, 16, 18, 16)
        fl.setSpacing(10)
        fl.setLabelAlignment(QtCore.Qt.AlignRight)

        self._squid_port = QtWidgets.QComboBox()
        self._squid_port.setEditable(True)
        self._squid_port.addItems(["COM1", "COM2", "COM3", "COM4"])
        fl.addRow("Serial port:", self._squid_port)

        self._squid_baud = QtWidgets.QComboBox()
        self._squid_baud.addItems(["1200", "2400", "4800", "9600", "19200"])
        self._squid_baud.setCurrentText("9600")
        fl.addRow("Baud rate:", self._squid_baud)

        self._squid_range = QtWidgets.QComboBox()
        self._squid_range.addItems(["1×", "10×", "100×", "1000×"])
        fl.addRow("Default range:", self._squid_range)

        self._squid_samples = QtWidgets.QSpinBox()
        self._squid_samples.setRange(1, 64)
        self._squid_samples.setValue(4)
        fl.addRow("Samples per position:", self._squid_samples)

        self._squid_settle = QtWidgets.QDoubleSpinBox()
        self._squid_settle.setRange(0.0, 30.0)
        self._squid_settle.setValue(1.5)
        self._squid_settle.setSuffix(" s")
        fl.addRow("Settle time:", self._squid_settle)

        return w

    # ── IRM / ARM tab ─────────────────────────────────────────────────────────
    def _build_irm_arm_tab(self) -> QtWidgets.QWidget:
        w = QtWidgets.QWidget()
        fl = QtWidgets.QFormLayout(w)
        fl.setContentsMargins(18, 16, 18, 16)
        fl.setSpacing(10)
        fl.setLabelAlignment(QtCore.Qt.AlignRight)

        fl.addRow(_sec_hdr("IRM (Isothermal Remanence)"))

        self._irm_max_field = QtWidgets.QDoubleSpinBox()
        self._irm_max_field.setRange(0, 2000)
        self._irm_max_field.setValue(1000)
        self._irm_max_field.setSuffix(" mT")
        self._irm_max_field.setSingleStep(50)
        fl.addRow("Max IRM field:", self._irm_max_field)

        self._irm_axis = QtWidgets.QComboBox()
        self._irm_axis.addItems(["Z (up-axis)", "X", "Y"])
        fl.addRow("Default axis:", self._irm_axis)

        self._irm_ramp = QtWidgets.QComboBox()
        self._irm_ramp.addItems(["Slow (60 s)", "Medium (30 s)", "Fast (10 s)"])
        fl.addRow("Ramp speed:", self._irm_ramp)

        self._irm_steps = QtWidgets.QSpinBox()
        self._irm_steps.setRange(1, 50)
        self._irm_steps.setValue(10)
        fl.addRow("Default step count:", self._irm_steps)

        fl.addRow(_sec_hdr("ARM (Anhysteretic Remanence)"))

        self._arm_peak_af = QtWidgets.QDoubleSpinBox()
        self._arm_peak_af.setRange(0, 200)
        self._arm_peak_af.setValue(100)
        self._arm_peak_af.setSuffix(" mT")
        fl.addRow("Default peak AF:", self._arm_peak_af)

        self._arm_bias = QtWidgets.QDoubleSpinBox()
        self._arm_bias.setRange(0, 100)
        self._arm_bias.setValue(0.05)
        self._arm_bias.setDecimals(3)
        self._arm_bias.setSingleStep(0.005)
        self._arm_bias.setSuffix(" mT")
        fl.addRow("Bias field:", self._arm_bias)

        return w

    # ── AF Demag tab ───────────────────────────────────────────────────────────
    def _build_af_demag_tab(self) -> QtWidgets.QWidget:
        w = QtWidgets.QWidget()
        fl = QtWidgets.QFormLayout(w)
        fl.setContentsMargins(18, 16, 18, 16)
        fl.setSpacing(10)
        fl.setLabelAlignment(QtCore.Qt.AlignRight)

        fl.addRow(_sec_hdr("ADwin Connection"))

        self._af_board = QtWidgets.QSpinBox()
        self._af_board.setRange(0, 7)
        self._af_board.setValue(0)
        fl.addRow("ADwin board #:", self._af_board)

        fl.addRow(_sec_hdr("Coil Parameters"))

        self._af_peak = QtWidgets.QDoubleSpinBox()
        self._af_peak.setRange(0, 300)
        self._af_peak.setValue(180)
        self._af_peak.setSuffix(" mT")
        fl.addRow("Peak AF field:", self._af_peak)

        self._af_ramp_speed = QtWidgets.QComboBox()
        self._af_ramp_speed.addItems(["Slow (3 Hz)", "Medium (8 Hz)", "Fast (15 Hz)"])
        fl.addRow("Ramp speed:", self._af_ramp_speed)

        self._af_settle = QtWidgets.QDoubleSpinBox()
        self._af_settle.setRange(0.1, 10.0)
        self._af_settle.setValue(1.5)
        self._af_settle.setSuffix(" s")
        self._af_settle.setSingleStep(0.1)
        fl.addRow("Settle time:", self._af_settle)

        fl.addRow(_sec_hdr("3-Axis Tumbling"))

        self._af_tumble = QtWidgets.QCheckBox("Enable 3-axis tumbling sequence")
        fl.addRow("", self._af_tumble)

        self._af_tumble_pause = QtWidgets.QDoubleSpinBox()
        self._af_tumble_pause.setRange(0.1, 5.0)
        self._af_tumble_pause.setValue(0.5)
        self._af_tumble_pause.setSuffix(" s")
        self._af_tumble_pause.setSingleStep(0.1)
        fl.addRow("Inter-axis pause:", self._af_tumble_pause)

        return w

    # ── Vacuum tab ─────────────────────────────────────────────────────────────
    def _build_vacuum_tab(self) -> QtWidgets.QWidget:
        w = QtWidgets.QWidget()
        fl = QtWidgets.QFormLayout(w)
        fl.setContentsMargins(18, 16, 18, 16)
        fl.setSpacing(10)
        fl.setLabelAlignment(QtCore.Qt.AlignRight)

        fl.addRow(_sec_hdr("Sensor Connection"))

        self._vac_port = QtWidgets.QComboBox()
        self._vac_port.setEditable(True)
        self._vac_port.addItems([f"COM{i}" for i in range(1, 9)])
        fl.addRow("Serial port:", self._vac_port)

        self._vac_baud = QtWidgets.QComboBox()
        self._vac_baud.addItems(["1200", "2400", "4800", "9600"])
        self._vac_baud.setCurrentText("9600")
        fl.addRow("Baud rate:", self._vac_baud)

        fl.addRow(_sec_hdr("Pressure Thresholds"))

        self._vac_target = QtWidgets.QDoubleSpinBox()
        self._vac_target.setRange(0.0, 100.0)
        self._vac_target.setValue(5.0)
        self._vac_target.setSuffix(" mTorr")
        fl.addRow("Target pressure:", self._vac_target)

        self._vac_warn = QtWidgets.QDoubleSpinBox()
        self._vac_warn.setRange(0.0, 200.0)
        self._vac_warn.setValue(20.0)
        self._vac_warn.setSuffix(" mTorr")
        fl.addRow("Warning threshold:", self._vac_warn)

        self._vac_auto_pump = QtWidgets.QCheckBox("Auto-start pump on program launch")
        fl.addRow("", self._vac_auto_pump)

        self._vac_poll = QtWidgets.QDoubleSpinBox()
        self._vac_poll.setRange(0.5, 30.0)
        self._vac_poll.setValue(2.0)
        self._vac_poll.setSuffix(" s")
        fl.addRow("Poll interval:", self._vac_poll)

        return w

    # ── Data Files tab ─────────────────────────────────────────────────────────
    def _build_data_files_tab(self) -> QtWidgets.QWidget:
        w = QtWidgets.QWidget()
        fl = QtWidgets.QFormLayout(w)
        fl.setContentsMargins(18, 16, 18, 16)
        fl.setSpacing(10)
        fl.setLabelAlignment(QtCore.Qt.AlignRight)

        fl.addRow(_sec_hdr("File Format"))

        self._df_format = QtWidgets.QComboBox()
        self._df_format.addItems(["CSV (comma-separated)", "Tab-Delimited", "SRM Format"])
        fl.addRow("Output format:", self._df_format)

        self._df_naming = QtWidgets.QComboBox()
        self._df_naming.addItems([
            "SampleName_Date",
            "Date_SampleName",
            "SampleName only",
        ])
        fl.addRow("File naming:", self._df_naming)

        self._df_decimals = QtWidgets.QSpinBox()
        self._df_decimals.setRange(2, 10)
        self._df_decimals.setValue(6)
        fl.addRow("Decimal places:", self._df_decimals)

        fl.addRow(_sec_hdr("Auto-Save"))

        self._df_auto_save = QtWidgets.QCheckBox("Auto-save after each position")
        self._df_auto_save.setChecked(True)
        fl.addRow("", self._df_auto_save)

        self._df_auto_append = QtWidgets.QCheckBox("Append to existing file (do not overwrite)")
        fl.addRow("", self._df_auto_append)

        self._df_backup = QtWidgets.QCheckBox("Create .bak backup before overwrite")
        self._df_backup.setChecked(True)
        fl.addRow("", self._df_backup)

        fl.addRow(_sec_hdr("Headers"))

        self._df_write_header = QtWidgets.QCheckBox("Write column headers on first row")
        self._df_write_header.setChecked(True)
        fl.addRow("", self._df_write_header)

        self._df_write_meta = QtWidgets.QCheckBox("Include sample metadata block at top")
        self._df_write_meta.setChecked(True)
        fl.addRow("", self._df_write_meta)

        return w

    # ── Changer tab ────────────────────────────────────────────────────────────
    def _build_changer_tab(self) -> QtWidgets.QWidget:
        w = QtWidgets.QWidget()
        scroll = QtWidgets.QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QtWidgets.QFrame.NoFrame)

        inner = QtWidgets.QWidget()
        fl = QtWidgets.QFormLayout(inner)
        fl.setContentsMargins(18, 16, 18, 16)
        fl.setSpacing(10)
        fl.setLabelAlignment(QtCore.Qt.AlignRight)

        fl.addRow(_sec_hdr("Motor Connection"))

        self._ch_port = QtWidgets.QComboBox()
        self._ch_port.setEditable(True)
        self._ch_port.addItems([f"COM{i}" for i in range(1, 9)])
        fl.addRow("Serial port:", self._ch_port)

        self._ch_baud = QtWidgets.QComboBox()
        self._ch_baud.addItems(["9600", "19200", "38400"])
        fl.addRow("Baud rate:", self._ch_baud)

        fl.addRow(_sec_hdr("Stage Speeds"))

        self._ch_speed_xy = QtWidgets.QDoubleSpinBox()
        self._ch_speed_xy.setRange(1, 100)
        self._ch_speed_xy.setValue(20)
        self._ch_speed_xy.setSuffix(" %")
        fl.addRow("XY traverse speed:", self._ch_speed_xy)

        self._ch_speed_z = QtWidgets.QDoubleSpinBox()
        self._ch_speed_z.setRange(1, 100)
        self._ch_speed_z.setValue(15)
        self._ch_speed_z.setSuffix(" %")
        fl.addRow("Z up/down speed:", self._ch_speed_z)

        fl.addRow(_sec_hdr("Home Position"))

        self._ch_home_x = QtWidgets.QDoubleSpinBox()
        self._ch_home_x.setRange(-999, 999)
        self._ch_home_x.setValue(0.0)
        self._ch_home_x.setSuffix(" mm")
        fl.addRow("Home X:", self._ch_home_x)

        self._ch_home_y = QtWidgets.QDoubleSpinBox()
        self._ch_home_y.setRange(-999, 999)
        self._ch_home_y.setValue(0.0)
        self._ch_home_y.setSuffix(" mm")
        fl.addRow("Home Y:", self._ch_home_y)

        self._ch_home_z = QtWidgets.QDoubleSpinBox()
        self._ch_home_z.setRange(-999, 999)
        self._ch_home_z.setValue(0.0)
        self._ch_home_z.setSuffix(" mm")
        fl.addRow("Home Z (inserted):", self._ch_home_z)

        fl.addRow(_sec_hdr("Safety"))

        self._ch_soft_limits = QtWidgets.QCheckBox("Enable software travel limits")
        self._ch_soft_limits.setChecked(True)
        fl.addRow("", self._ch_soft_limits)

        self._ch_limit_stop = QtWidgets.QCheckBox("Stop on limit-switch contact (halt all motion)")
        self._ch_limit_stop.setChecked(True)
        fl.addRow("", self._ch_limit_stop)

        scroll.setWidget(inner)
        vl = QtWidgets.QVBoxLayout(w)
        vl.setContentsMargins(0, 0, 0, 0)
        vl.addWidget(scroll)
        return w

    # ── Calibration tab ────────────────────────────────────────────────────────
    def _build_calibration_tab(self) -> QtWidgets.QWidget:
        w = QtWidgets.QWidget()
        scroll = QtWidgets.QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QtWidgets.QFrame.NoFrame)

        inner = QtWidgets.QWidget()
        fl = QtWidgets.QFormLayout(inner)
        fl.setContentsMargins(18, 16, 18, 16)
        fl.setSpacing(10)
        fl.setLabelAlignment(QtCore.Qt.AlignRight)

        fl.addRow(_sec_hdr("Cal Rod Reference"))

        self._cal_rod_moment = QtWidgets.QDoubleSpinBox()
        self._cal_rod_moment.setRange(1e-12, 1.0)
        self._cal_rod_moment.setDecimals(8)
        self._cal_rod_moment.setValue(1.234e-5)
        self._cal_rod_moment.setSuffix(" A·m²")
        fl.addRow("Cal rod moment:", self._cal_rod_moment)

        self._cal_date = QtWidgets.QDateEdit()
        self._cal_date.setCalendarPopup(True)
        self._cal_date.setDate(QtCore.QDate.currentDate())
        fl.addRow("Calibration date:", self._cal_date)

        fl.addRow(_sec_hdr("SQUID Sensitivity Factors"))

        self._cal_x = QtWidgets.QDoubleSpinBox()
        self._cal_x.setRange(-1_000_000, 1_000_000)
        self._cal_x.setDecimals(6)
        self._cal_x.setValue(1.0)
        fl.addRow("X-axis cal factor:", self._cal_x)

        self._cal_y = QtWidgets.QDoubleSpinBox()
        self._cal_y.setRange(-1_000_000, 1_000_000)
        self._cal_y.setDecimals(6)
        self._cal_y.setValue(1.0)
        fl.addRow("Y-axis cal factor:", self._cal_y)

        self._cal_z = QtWidgets.QDoubleSpinBox()
        self._cal_z.setRange(-1_000_000, 1_000_000)
        self._cal_z.setDecimals(6)
        self._cal_z.setValue(1.0)
        fl.addRow("Z-axis cal factor:", self._cal_z)

        self._range_factor = QtWidgets.QDoubleSpinBox()
        self._range_factor.setRange(1e-12, 1.0)
        self._range_factor.setDecimals(8)
        self._range_factor.setValue(1e-5)
        fl.addRow("Range factor:", self._range_factor)

        fl.addRow(_sec_hdr("Background Subtraction"))

        self._bg_x = QtWidgets.QDoubleSpinBox()
        self._bg_x.setRange(-10, 10)
        self._bg_x.setDecimals(6)
        self._bg_x.setValue(0.0)
        fl.addRow("Background X (V):", self._bg_x)

        self._bg_y = QtWidgets.QDoubleSpinBox()
        self._bg_y.setRange(-10, 10)
        self._bg_y.setDecimals(6)
        self._bg_y.setValue(0.0)
        fl.addRow("Background Y (V):", self._bg_y)

        self._bg_z = QtWidgets.QDoubleSpinBox()
        self._bg_z.setRange(-10, 10)
        self._bg_z.setDecimals(6)
        self._bg_z.setValue(0.0)
        fl.addRow("Background Z (V):", self._bg_z)

        self._bg_subtract = QtWidgets.QCheckBox("Apply background subtraction to all readings")
        fl.addRow("", self._bg_subtract)

        scroll.setWidget(inner)
        vl = QtWidgets.QVBoxLayout(w)
        vl.setContentsMargins(0, 0, 0, 0)
        vl.addWidget(scroll)
        return w

    # ── Sequence tab ───────────────────────────────────────────────────────────
    def _build_sequence_tab(self) -> QtWidgets.QWidget:
        """Per-step-type time estimate inputs for the runtime estimator."""
        w = QtWidgets.QWidget()
        scroll = QtWidgets.QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QtWidgets.QFrame.NoFrame)

        inner = QtWidgets.QWidget()
        fl = QtWidgets.QFormLayout(inner)
        fl.setContentsMargins(18, 16, 18, 16)
        fl.setSpacing(10)
        fl.setLabelAlignment(QtCore.Qt.AlignRight)

        fl.addRow(_sec_hdr("Step Time Estimates (seconds)"))
        note = QtWidgets.QLabel(
            "These estimates are used to compute estimated sequence run time "
            "and display a live countdown in the status bar."
        )
        note.setWordWrap(True)
        note.setStyleSheet("color: #7a6f6e; font-size: 12px;")
        fl.addRow(note)

        def _spin(default: int) -> QtWidgets.QSpinBox:
            s = QtWidgets.QSpinBox()
            s.setRange(5, 7200)
            s.setSingleStep(10)
            s.setSuffix(" s")
            s.setValue(default)
            return s

        self._seq_nrm  = _spin(120);  fl.addRow("NRM measurement:",           self._seq_nrm)
        self._seq_af   = _spin(180);  fl.addRow("AF demagnetisation step:",   self._seq_af)
        self._seq_tt   = _spin(600);  fl.addRow("Thermal (TT/TH) step:",      self._seq_tt)
        self._seq_irm  = _spin(300);  fl.addRow("IRM acquisition step:",       self._seq_irm)
        self._seq_arm  = _spin(240);  fl.addRow("ARM acquisition step:",       self._seq_arm)
        self._seq_rrm  = _spin(200);  fl.addRow("RRM step:",                   self._seq_rrm)
        self._seq_ptrm = _spin(600);  fl.addRow("pTRM check step:",            self._seq_ptrm)
        self._seq_def  = _spin(120);  fl.addRow("Default (unknown type):",     self._seq_def)

        scroll.setWidget(inner)
        vl = QtWidgets.QVBoxLayout(w)
        vl.setContentsMargins(0, 0, 0, 0)
        vl.addWidget(scroll)
        return w

    # ── Config I/O ────────────────────────────────────────────────────────────

    def load_from_config(self, cfg: AppConfig) -> None:
        """Populate all widgets from *cfg*."""
        g = cfg.general
        self._data_dir.setText(g.data_dir)
        self._sample_dir.setText(g.sample_dir)
        self._backup_dir.setText(g.backup_dir)
        self._operator.setText(g.operator)
        self._lab_name.setText(g.lab_name)
        self._nocomm_chk.setChecked(g.nocomm)
        self._auto_save.setChecked(g.auto_save)

        sq = cfg.squid
        self._squid_port.setCurrentText(sq.port)
        self._squid_baud.setCurrentText(str(sq.baud))
        self._squid_range.setCurrentText(sq.range_label)
        self._squid_samples.setValue(sq.samples_per_pos)
        self._squid_settle.setValue(sq.settle_time)

        ia = cfg.irm_arm
        self._irm_max_field.setValue(ia.irm_max_field)
        self._irm_axis.setCurrentText(ia.irm_axis)
        self._irm_ramp.setCurrentText(ia.irm_ramp)
        self._irm_steps.setValue(ia.irm_steps)
        self._arm_peak_af.setValue(ia.arm_peak_af)
        self._arm_bias.setValue(ia.arm_bias)

        af = cfg.af_demag
        self._af_board.setValue(af.board)
        self._af_peak.setValue(af.peak)
        self._af_ramp_speed.setCurrentText(af.ramp_speed)
        self._af_settle.setValue(af.settle)
        self._af_tumble.setChecked(af.tumble)
        self._af_tumble_pause.setValue(af.tumble_pause)

        v = cfg.vacuum
        self._vac_port.setCurrentText(v.port)
        self._vac_baud.setCurrentText(str(v.baud))
        self._vac_target.setValue(v.target_pressure)
        self._vac_warn.setValue(v.warn_threshold)
        self._vac_auto_pump.setChecked(v.auto_pump)
        self._vac_poll.setValue(v.poll_interval)

        df = cfg.data_files
        self._df_format.setCurrentText(df.format)
        self._df_naming.setCurrentText(df.naming)
        self._df_decimals.setValue(df.decimals)
        self._df_auto_save.setChecked(df.auto_save)
        self._df_auto_append.setChecked(df.auto_append)
        self._df_backup.setChecked(df.backup)
        self._df_write_header.setChecked(df.write_header)
        self._df_write_meta.setChecked(df.write_meta)

        ch = cfg.changer
        self._ch_port.setCurrentText(ch.port)
        self._ch_baud.setCurrentText(str(ch.baud))
        self._ch_speed_xy.setValue(ch.speed_xy)
        self._ch_speed_z.setValue(ch.speed_z)
        self._ch_home_x.setValue(ch.home_x)
        self._ch_home_y.setValue(ch.home_y)
        self._ch_home_z.setValue(ch.home_z)
        self._ch_soft_limits.setChecked(ch.soft_limits)
        self._ch_limit_stop.setChecked(ch.limit_stop)

        cal = cfg.calibration
        self._cal_rod_moment.setValue(cal.cal_rod_moment)
        if cal.cal_date_iso:
            try:
                from datetime import date as _date
                d = _date.fromisoformat(cal.cal_date_iso)
                self._cal_date.setDate(QtCore.QDate(d.year, d.month, d.day))
            except ValueError:
                pass
        self._cal_x.setValue(cal.cal_x)
        self._cal_y.setValue(cal.cal_y)
        self._cal_z.setValue(cal.cal_z)
        self._range_factor.setValue(cal.range_factor)
        self._bg_x.setValue(cal.bg_x)
        self._bg_y.setValue(cal.bg_y)
        self._bg_z.setValue(cal.bg_z)
        self._bg_subtract.setChecked(cal.bg_subtract)

        seq = cfg.sequence
        self._seq_nrm.setValue(seq.NRM)
        self._seq_af.setValue(seq.AF)
        self._seq_tt.setValue(seq.TT)
        self._seq_irm.setValue(seq.IRM)
        self._seq_arm.setValue(seq.ARM)
        self._seq_rrm.setValue(seq.RRM)
        self._seq_ptrm.setValue(seq.PTRM)
        self._seq_def.setValue(seq.default)

    def save_to_config(self, cfg: AppConfig) -> None:
        """Read all widget values into *cfg* (mutates in-place)."""
        g = cfg.general
        g.data_dir    = self._data_dir.text()
        g.sample_dir  = self._sample_dir.text()
        g.backup_dir  = self._backup_dir.text()
        g.operator    = self._operator.text()
        g.lab_name    = self._lab_name.text()
        g.nocomm      = self._nocomm_chk.isChecked()
        g.auto_save   = self._auto_save.isChecked()

        sq = cfg.squid
        sq.port            = self._squid_port.currentText()
        sq.baud            = int(self._squid_baud.currentText())
        sq.range_label     = self._squid_range.currentText()
        sq.samples_per_pos = self._squid_samples.value()
        sq.settle_time     = self._squid_settle.value()

        ia = cfg.irm_arm
        ia.irm_max_field = self._irm_max_field.value()
        ia.irm_axis      = self._irm_axis.currentText()
        ia.irm_ramp      = self._irm_ramp.currentText()
        ia.irm_steps     = self._irm_steps.value()
        ia.arm_peak_af   = self._arm_peak_af.value()
        ia.arm_bias      = self._arm_bias.value()

        af = cfg.af_demag
        af.board        = self._af_board.value()
        af.peak         = self._af_peak.value()
        af.ramp_speed   = self._af_ramp_speed.currentText()
        af.settle       = self._af_settle.value()
        af.tumble       = self._af_tumble.isChecked()
        af.tumble_pause = self._af_tumble_pause.value()

        v = cfg.vacuum
        v.port            = self._vac_port.currentText()
        v.baud            = int(self._vac_baud.currentText())
        v.target_pressure = self._vac_target.value()
        v.warn_threshold  = self._vac_warn.value()
        v.auto_pump       = self._vac_auto_pump.isChecked()
        v.poll_interval   = self._vac_poll.value()

        df = cfg.data_files
        df.format       = self._df_format.currentText()
        df.naming       = self._df_naming.currentText()
        df.decimals     = self._df_decimals.value()
        df.auto_save    = self._df_auto_save.isChecked()
        df.auto_append  = self._df_auto_append.isChecked()
        df.backup       = self._df_backup.isChecked()
        df.write_header = self._df_write_header.isChecked()
        df.write_meta   = self._df_write_meta.isChecked()

        ch = cfg.changer
        ch.port        = self._ch_port.currentText()
        ch.baud        = int(self._ch_baud.currentText())
        ch.speed_xy    = self._ch_speed_xy.value()
        ch.speed_z     = self._ch_speed_z.value()
        ch.home_x      = self._ch_home_x.value()
        ch.home_y      = self._ch_home_y.value()
        ch.home_z      = self._ch_home_z.value()
        ch.soft_limits = self._ch_soft_limits.isChecked()
        ch.limit_stop  = self._ch_limit_stop.isChecked()

        cal = cfg.calibration
        cal.cal_rod_moment = self._cal_rod_moment.value()
        qd = self._cal_date.date()
        cal.cal_date_iso = f"{qd.year():04d}-{qd.month():02d}-{qd.day():02d}"
        cal.cal_x        = self._cal_x.value()
        cal.cal_y        = self._cal_y.value()
        cal.cal_z        = self._cal_z.value()
        cal.range_factor = self._range_factor.value()
        cal.bg_x         = self._bg_x.value()
        cal.bg_y         = self._bg_y.value()
        cal.bg_z         = self._bg_z.value()
        cal.bg_subtract  = self._bg_subtract.isChecked()

        seq = cfg.sequence
        seq.NRM     = self._seq_nrm.value()
        seq.AF      = self._seq_af.value()
        seq.TH      = self._seq_tt.value()
        seq.TT      = self._seq_tt.value()
        seq.IRM     = self._seq_irm.value()
        seq.ARM     = self._seq_arm.value()
        seq.RRM     = self._seq_rrm.value()
        seq.PTRM    = self._seq_ptrm.value()
        seq.default = self._seq_def.value()

    # ── Actions ───────────────────────────────────────────────────────────────
    def _save(self) -> None:
        mw = self.window()
        if hasattr(mw, "config"):
            self.save_to_config(mw.config)
            mw.config.save()
            # Propagate updated step times to the runtime estimator
            if hasattr(mw, "_estimator"):
                mw._estimator.step_times = mw.config.sequence.as_estimator_dict()
        QtWidgets.QMessageBox.information(self, "Settings", "Settings saved.")

    def _revert(self) -> None:
        mw = self.window()
        if hasattr(mw, "navigate_to"):
            mw.navigate_to("dashboard")
