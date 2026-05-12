from __future__ import annotations

import sys
from datetime import datetime
from pathlib import Path

from PySide6 import QtCore, QtWidgets

from rapidpy_common.ui import apply_liquid_glass_theme, set_app_icon

from .config import AppConfig
from .dialogs import (
    AboutDialog,
    DebugConsoleDialog,
    IrmArmDialog,
    LoginDialog,
    PlotsDialog,
    SampleSelectDialog,
    SquidCommDialog,
    StepMonitorDialog,
    VacuumDialog,
    WebcamDialog,
)
from .panels import (
    DashboardPanel,
    MeasurementPanel,
    SampleQueuePanel,
    SequencePanel,
    SettingsPanel,
)
from .runtime_estimator import RuntimeEstimator


# ── Extra stylesheet (appended to shared theme) ───────────────────────────────
_EXTRA_CSS = """
    QFrame#sidebar {
        background: rgba(122, 2, 25, 0.04);
        border-right: 1px solid rgba(122, 2, 25, 0.14);
        border-radius: 0px;
    }
    QPushButton#navBtn {
        background: transparent;
        border: none;
        border-radius: 10px;
        padding: 10px 16px;
        text-align: left;
        color: #4d3a39;
        font-size: 13px;
    }
    QPushButton#navBtn:hover  { background: rgba(122, 2, 25, 0.08); }
    QPushButton#navBtn:checked {
        background: rgba(122, 2, 25, 0.14);
        color: #7A0219;
        font-weight: 680;
    }
    QToolBar {
        background: transparent;
        border: none;
        padding: 0;
        margin: 0;
        spacing: 0;
    }
    QFrame#header {
        background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
            stop:0 #fffdf9, stop:1 #f5ede2);
        border-top: 3px solid #7A0219;
        border-bottom: 2px solid rgba(122, 2, 25, 0.20);
        border-radius: 0px;
    }
    QPushButton#headerBtn {
        background: rgba(122, 2, 25, 0.07);
        border: 1px solid rgba(122, 2, 25, 0.18);
        border-radius: 8px;
        padding: 4px 11px;
        font-size: 12px;
        color: #4d3a39;
        min-height: 26px;
    }
    QPushButton#headerBtn:hover {
        background: rgba(122, 2, 25, 0.13);
        border-color: rgba(122, 2, 25, 0.30);
    }
    QPushButton#headerBtn:checked {
        background: rgba(107, 114, 128, 0.18);
        border-color: rgba(107, 114, 128, 0.35);
        color: #374151;
        font-weight: 600;
    }
    QPushButton#headerBtnHalt {
        background: rgba(220, 38, 38, 0.07);
        border: 1px solid rgba(220, 38, 38, 0.22);
        border-radius: 8px;
        padding: 4px 11px;
        font-size: 12px;
        color: #b91c1c;
        min-height: 26px;
    }
    QPushButton#headerBtnHalt:hover { background: rgba(220, 38, 38, 0.14); }
    QPushButton#headerBtnExit {
        background: #7A0219;
        border: none;
        border-radius: 8px;
        padding: 4px 13px;
        font-size: 12px;
        color: #ffffff;
        font-weight: 600;
        min-height: 26px;
    }
    QPushButton#headerBtnExit:hover { background: #9c0220; }
    QLabel#headerTitle {
        font-size: 15px;
        font-weight: 700;
        color: #7A0219;
        letter-spacing: 0.5px;
    }
    QLabel#flowRunning {
        background: rgba(34, 197, 94, 0.15); border: 1px solid rgba(34, 197, 94, 0.45);
        border-radius: 8px; padding: 3px 10px; color: #15803d; font-weight: 600;
    }
    QLabel#flowPaused {
        background: rgba(251, 191, 36, 0.18); border: 1px solid rgba(251, 191, 36, 0.55);
        border-radius: 8px; padding: 3px 10px; color: #92400e; font-weight: 600;
    }
    QLabel#flowHalted {
        background: rgba(220, 38, 38, 0.12); border: 1px solid rgba(220, 38, 38, 0.4);
        border-radius: 8px; padding: 3px 10px; color: #b91c1c; font-weight: 600;
    }
    QLabel#flowNocomm {
        background: rgba(107, 114, 128, 0.12); border: 1px solid rgba(107, 114, 128, 0.4);
        border-radius: 8px; padding: 3px 10px; color: #374151; font-weight: 600;
    }
    QLabel#instOk {
        background: rgba(34,197,94,0.12); border: 1px solid rgba(34,197,94,0.35);
        border-radius: 7px; padding: 2px 9px; color: #15803d; font-size: 12px;
    }
    QLabel#instErr {
        background: rgba(220,38,38,0.10); border: 1px solid rgba(220,38,38,0.35);
        border-radius: 7px; padding: 2px 9px; color: #b91c1c; font-size: 12px;
    }
    QLabel#instUnk {
        background: rgba(107,114,128,0.10); border: 1px solid rgba(107,114,128,0.3);
        border-radius: 7px; padding: 2px 9px; color: #6b7280; font-size: 12px;
    }
    QLabel#readBig  { font-size: 22px; font-weight: 700; color: #2f2827; }
    QLabel#readMed  { font-size: 16px; font-weight: 600; color: #2f2827; }
    QLabel#readLbl  { font-size: 11px; color: #9a8885; }
    QLabel#sectionHdr {
        font-size: 10px; font-weight: 700; color: #9a8885; letter-spacing: 1.2px;
    }
    QLabel#warnOrange {
        background: rgba(251,146,60,0.15); border: 1px solid rgba(251,146,60,0.5);
        border-radius: 8px; padding: 5px 12px; color: #c2410c;
    }
    QLabel#warnRed {
        background: rgba(220,38,38,0.12); border: 1px solid rgba(220,38,38,0.45);
        border-radius: 8px; padding: 5px 12px; color: #b91c1c;
    }
    QLabel#valuePill {
        background: rgba(243,238,226,0.85); border: 1px solid rgba(122,2,25,0.18);
        border-radius: 7px; padding: 3px 10px; color: #2f2827;
    }
    QLabel#valueMonospace {
        background: rgba(28, 20, 19, 0.06); border: 1px solid rgba(122,2,25,0.18);
        border-radius: 7px; padding: 4px 10px; color: #2f2827;
        font-family: "Courier New", monospace;
    }
"""

_NAV_ITEMS: list[tuple[str, str, int]] = [
    ("🏠", "Dashboard",    0),
    ("📋", "Sample Queue", 1),
    ("🔬", "Sequence",     2),
    ("📊", "Live Measure", 3),
    ("⚙️", "Settings",    4),
]


class MainWindow(QtWidgets.QMainWindow):

    def __init__(self) -> None:
        super().__init__()
        self.setWindowTitle("RAPID v4 — Paleomagnetics Control System")
        self.resize(1440, 880)
        self._nav_btns: list[QtWidgets.QPushButton] = []

        # Persistent configuration (load or create defaults)
        self.config: AppConfig = AppConfig.load()

        # Runtime estimator — initialised from config step times
        self._estimator = RuntimeEstimator(self.config.sequence.as_estimator_dict())
        self._sequence_labels: list[str] = []  # current loaded sequence step labels

        # Live countdown timer (1 Hz, used when a run is active)
        self._run_start_time: "datetime | None" = None
        self._run_current_idx: int = 0
        self._run_timer = QtCore.QTimer(self)
        self._run_timer.setInterval(1000)
        self._run_timer.timeout.connect(self._tick_run_countdown)

        self._build_header()
        self._build_central()
        self._build_statusbar()
        self._build_menu()
        self._nav_select(0)

        # Populate settings panel from config
        self._settings_panel.load_from_config(self.config)

        self._clock = QtCore.QTimer(self)
        self._clock.timeout.connect(self._tick_clock)
        self._clock.start(1000)


    # ── Header toolbar ────────────────────────────────────────────────────────
    def _build_header(self) -> None:
        header = QtWidgets.QFrame()
        header.setObjectName("header")
        header.setFixedHeight(54)

        hl = QtWidgets.QHBoxLayout(header)
        hl.setContentsMargins(18, 0, 14, 0)
        hl.setSpacing(10)

        title = QtWidgets.QLabel("⚗  RAPID v4")
        title.setObjectName("headerTitle")
        hl.addWidget(title)
        hl.addWidget(_vline())

        self._flow_lbl = QtWidgets.QLabel("◉  Running")
        self._flow_lbl.setObjectName("flowRunning")
        hl.addWidget(self._flow_lbl)
        hl.addWidget(_vline())

        self._sample_hdr = QtWidgets.QLabel("Sample: —")
        self._sample_hdr.setStyleSheet("color: #4d3a39; font-size: 13px;")
        hl.addWidget(self._sample_hdr)

        self._step_hdr = QtWidgets.QLabel("Step: —")
        self._step_hdr.setStyleSheet("color: #7a6f6e; font-size: 12px;")
        hl.addWidget(self._step_hdr)

        hl.addStretch()

        self._pause_btn = QtWidgets.QPushButton("⏸  Pause")
        self._pause_btn.setObjectName("headerBtn")
        self._halt_btn  = QtWidgets.QPushButton("■  Halt")
        self._halt_btn.setObjectName("headerBtnHalt")
        self._nocomm_btn = QtWidgets.QPushButton("⊘  No-Comm")
        self._nocomm_btn.setObjectName("headerBtn")
        self._nocomm_btn.setCheckable(True)
        self._nocomm_btn.toggled.connect(self._on_nocomm_toggled)

        quit_btn = QtWidgets.QPushButton("✕  Exit")
        quit_btn.setObjectName("headerBtnExit")
        quit_btn.clicked.connect(QtWidgets.QApplication.quit)

        for btn in (self._pause_btn, self._halt_btn, self._nocomm_btn, quit_btn):
            hl.addWidget(btn)

        tb = QtWidgets.QToolBar()
        tb.setMovable(False)
        tb.setFloatable(False)
        tb.setStyleSheet("QToolBar { border: none; padding: 0; margin: 0; }")
        tb.addWidget(header)
        self.addToolBar(QtCore.Qt.TopToolBarArea, tb)

    # ── Central widget: sidebar + stacked panels ──────────────────────────────
    def _build_central(self) -> None:
        root = QtWidgets.QWidget()
        layout = QtWidgets.QHBoxLayout(root)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(0)
        self.setCentralWidget(root)

        # ── Sidebar ──
        sidebar = QtWidgets.QFrame()
        sidebar.setObjectName("sidebar")
        sidebar.setFixedWidth(210)
        sl = QtWidgets.QVBoxLayout(sidebar)
        sl.setContentsMargins(8, 14, 8, 14)
        sl.setSpacing(2)

        def _sec_hdr(text: str) -> QtWidgets.QLabel:
            lbl = QtWidgets.QLabel(text)
            lbl.setObjectName("sectionHdr")
            lbl.setContentsMargins(8, 0, 0, 4)
            return lbl

        sl.addWidget(_sec_hdr("MAIN"))
        self._btn_group = QtWidgets.QButtonGroup(self)
        self._btn_group.setExclusive(True)
        for icon, label, idx in _NAV_ITEMS:
            btn = QtWidgets.QPushButton(f"  {icon}  {label}")
            btn.setObjectName("navBtn")
            btn.setCheckable(True)
            btn.setMinimumHeight(42)
            btn.clicked.connect(lambda _checked, i=idx: self._nav_select(i))
            self._btn_group.addButton(btn)
            self._nav_btns.append(btn)
            sl.addWidget(btn)

        sl.addSpacing(16)
        sl.addWidget(_sec_hdr("DIAGNOSTICS"))

        for icon, label, slot in [
            ("🔌", "DC Motors",  self._launch_dc_motors),
            ("🌊", "AF Demag",   self._launch_af),
            ("🧲", "IRM / ARM",  self._launch_irm),
            ("💧", "Vacuum",     self._launch_vacuum),
            ("🔭", "SQUID Comm", self._launch_squid),
        ]:
            btn = QtWidgets.QPushButton(f"  {icon}  {label}")
            btn.setObjectName("navBtn")
            btn.setMinimumHeight(38)
            btn.clicked.connect(slot)
            sl.addWidget(btn)

        sl.addStretch()
        ver = QtWidgets.QLabel("RAPID v4.0 · Phase 2")
        ver.setStyleSheet("color: #c4b7b3; font-size: 10px; padding: 0 8px;")
        sl.addWidget(ver)

        # ── Stacked panels ──
        self._stack = QtWidgets.QStackedWidget()
        self._dashboard   = DashboardPanel()
        self._sample_queue = SampleQueuePanel()
        self._sequence    = SequencePanel()
        self._measurement = MeasurementPanel()
        self._settings_panel    = SettingsPanel()
        for panel in (self._dashboard, self._sample_queue, self._sequence,
                      self._measurement, self._settings_panel):
            self._stack.addWidget(panel)

        layout.addWidget(sidebar)
        layout.addWidget(self._stack)

    # ── Status bar ────────────────────────────────────────────────────────────
    def _build_statusbar(self) -> None:
        sb = self.statusBar()
        sb.setSizeGripEnabled(False)
        sb.setStyleSheet("QStatusBar { border-top: 1px solid rgba(122,2,25,0.12); }")

        self._sb_status  = QtWidgets.QLabel("Initializing…")
        self._sb_pos     = QtWidgets.QLabel("Pos: —")
        self._sb_sample  = QtWidgets.QLabel("Sample: —")
        self._sb_runtime = QtWidgets.QLabel("No sequence loaded")
        self._sb_time    = QtWidgets.QLabel("00:00:00")

        for lbl in (self._sb_status, self._sb_pos, self._sb_sample,
                    self._sb_runtime, self._sb_time):
            lbl.setStyleSheet("padding: 1px 10px; color: #4d3a39; font-size: 12px;")

        sb.addWidget(self._sb_status, 3)
        sb.addWidget(_vline(), 0)
        sb.addWidget(self._sb_pos, 1)
        sb.addWidget(_vline(), 0)
        sb.addWidget(self._sb_sample, 1)
        sb.addWidget(_vline(), 0)
        sb.addWidget(self._sb_runtime, 2)
        sb.addPermanentWidget(self._sb_time)

    def _update_runtime_display(
        self,
        current_idx: int = 0,
        *,
        running: bool = False,
        start_time: "datetime | None" = None,
    ) -> None:
        """Update the runtime status bar label.

        Parameters
        ----------
        current_idx:
            0-based index of the step currently executing (or next to run).
        running:
            When True shows remaining time; when False shows total estimate.
        start_time:
            Wall-clock time the run started (for ETA calculation).
        """
        labels = self._sequence_labels
        if not labels:
            self._sb_runtime.setText("No sequence loaded")
            return
        if running:
            text = self._estimator.status_bar_text(labels, current_idx, start_time=start_time)
        else:
            text = self._estimator.total_bar_text(labels)
        self._sb_runtime.setText(text)


    def load_sequence_labels(self, labels: list[str]) -> None:
        """Set the sequence step labels and update the status bar estimate."""
        self._sequence_labels = list(labels)
        self._update_runtime_display(running=False)

    def start_run(self, labels: list[str] | None = None) -> None:
        """Begin a measurement run — start the live countdown timer."""
        if labels is not None:
            self._sequence_labels = list(labels)
        self._run_current_idx = 0
        self._run_start_time = datetime.now()
        self._run_timer.start()
        self._update_runtime_display(current_idx=0, running=True)
        self._flow_lbl.setText("◉  Running")
        self._flow_lbl.setObjectName("flowRunning")
        self._flow_lbl.style().polish(self._flow_lbl)

    def stop_run(self) -> None:
        """End the measurement run — stop the countdown timer."""
        self._run_timer.stop()
        self._run_start_time = None
        self._run_current_idx = 0
        self._update_runtime_display(running=False)
        self._flow_lbl.setText("◎  Idle")
        self._flow_lbl.setObjectName("flowIdle")
        self._flow_lbl.style().polish(self._flow_lbl)

    def advance_step(self) -> None:
        """Call after each step completes to increment the countdown index."""
        self._run_current_idx += 1
        self._update_runtime_display(
            current_idx=self._run_current_idx,
            running=self._run_timer.isActive(),
            start_time=self._run_start_time,
        )

    @QtCore.Slot()
    def _tick_run_countdown(self) -> None:
        """Called every second during an active run to refresh the countdown."""
        self._update_runtime_display(
            current_idx=self._run_current_idx,
            running=True,
            start_time=self._run_start_time,
        )

    # ── Menu bar ──────────────────────────────────────────────────────────────
    def _build_menu(self) -> None:
        mb = self.menuBar()

        fm = mb.addMenu("&File")
        fm.addAction("&New Session")
        fm.addAction("&Log Out", self._launch_login)
        fm.addSeparator()
        fm.addAction("E&xit", QtWidgets.QApplication.quit)

        vm = mb.addMenu("&View")
        for label, idx in [("&Dashboard", 0), ("Sample &Queue", 1),
                            ("Se&quence", 2), ("Live &Measurement", 3), ("&Settings", 4)]:
            vm.addAction(label, lambda _c=False, i=idx: self._nav_select(i))
        vm.addSeparator()
        vm.addAction("Step &Monitor", self._launch_step_monitor)
        vm.addAction("&Debug Console", self._launch_debug_console)
        vm.addAction("&Sample Queue Monitor")
        vm.addSeparator()
        vm.addAction("&Webcam Monitor", self._launch_webcam)

        flm = mb.addMenu("F&low")
        flm.addAction("&Running")
        flm.addAction("&Paused")
        flm.addAction("&Halted")
        flm.addSeparator()
        flm.addAction("Code &Override")

        dm = mb.addMenu("&Diagnostics")
        dm.addAction("DC &Motors",    self._launch_dc_motors)
        dm.addAction("&SQUID Comm",   self._launch_squid)
        dm.addAction("&Vacuum",       self._launch_vacuum)
        dm.addSeparator()
        af_sub = dm.addMenu("AF &Demagnetizer")
        af_sub.addAction("AF Demag Window",     self._launch_af)
        af_sub.addAction("AF Tuner / ClipTest")
        af_sub.addAction("AF Field Calibration")
        irm_sub = dm.addMenu("&IRM / ARM")
        irm_sub.addAction("IRM / ARM Window",       self._launch_irm)
        irm_sub.addAction("IRM Field Calibration")
        irm_sub.addAction("IRM Voltage Calibration")
        dm.addAction("908A &Gaussmeter")
        dm.addAction("Susceptibility &Bridge")
        dm.addSeparator()
        dm.addAction("&VRM Data Collection")
        dm.addAction("Calibrate &Rod")

        hm = mb.addMenu("&Help")
        hm.addAction("&About RAPID", self._launch_about)

    # ── Navigation ────────────────────────────────────────────────────────────
    def _nav_select(self, index: int) -> None:
        self._stack.setCurrentIndex(index)
        if 0 <= index < len(self._nav_btns):
            self._nav_btns[index].setChecked(True)

    # ── Clock ─────────────────────────────────────────────────────────────────
    def _tick_clock(self) -> None:
        self._sb_time.setText(
            QtCore.QDateTime.currentDateTime().toString("hh:mm:ss AP")
        )

    # ── No-Comm toggle ────────────────────────────────────────────────────────
    def _on_nocomm_toggled(self, on: bool) -> None:
        if on:
            self._flow_lbl.setObjectName("flowNocomm")
            self._flow_lbl.setText("⊘  No-Comm")
        else:
            self._flow_lbl.setObjectName("flowRunning")
            self._flow_lbl.setText("◉  Running")
        self._flow_lbl.style().unpolish(self._flow_lbl)
        self._flow_lbl.style().polish(self._flow_lbl)

    # ── Diagnostic launchers (stubs — wired in Phase 3) ───────────────────────
    def _launch_dc_motors(self) -> None:
        _stub_dialog(self, "DC Motors", "Launches dc_motor_control app (Phase 3)")

    def _launch_af(self) -> None:
        _stub_dialog(self, "AF Demag", "Launches af_tuner app (Phase 3)")

    def _launch_irm(self) -> None:
        IrmArmDialog(self).exec()

    def _launch_vacuum(self) -> None:
        VacuumDialog(self).exec()

    def _launch_squid(self) -> None:
        SquidCommDialog(self).exec()

    def _launch_login(self) -> None:
        LoginDialog(self).exec()

    def _launch_about(self) -> None:
        AboutDialog(self).exec()

    def _launch_debug_console(self) -> None:
        if not hasattr(self, "_debug_dlg") or not self._debug_dlg.isVisible():
            self._debug_dlg = DebugConsoleDialog(self)
            self._debug_dlg.setModal(False)
        self._debug_dlg.show()
        self._debug_dlg.raise_()

    def _launch_step_monitor(self) -> None:
        if not hasattr(self, "_step_dlg") or not self._step_dlg.isVisible():
            self._step_dlg = StepMonitorDialog(self)
            self._step_dlg.setModal(False)
        self._step_dlg.show()
        self._step_dlg.raise_()

    def _launch_webcam(self) -> None:
        if not hasattr(self, "_webcam_dlg"):
            self._webcam_dlg = WebcamDialog(self)
        self._webcam_dlg.show()
        self._webcam_dlg.raise_()

    # ── Public API for panels ─────────────────────────────────────────────────
    def navigate_to(self, key: str) -> None:
        mapping = {"dashboard": 0, "queue": 1, "sequence": 2, "measure": 3, "settings": 4}
        if key in mapping:
            self._nav_select(mapping[key])

    def set_status(self, text: str) -> None:
        self._sb_status.setText(text)

    def set_sample(self, name: str) -> None:
        self._sb_sample.setText(f"Sample: {name}")
        self._sample_hdr.setText(f"Sample: {name}")

    def set_step(self, step: str) -> None:
        self._step_hdr.setText(f"Step: {step}")

    def set_position(self, pos: str) -> None:
        self._sb_pos.setText(f"Pos: {pos}")

    def set_flow_state(self, state: str) -> None:
        """state: 'running' | 'paused' | 'halted'"""
        icons = {"running": "◉  Running", "paused": "⏸  Paused", "halted": "■  Halted"}
        names = {"running": "flowRunning", "paused": "flowPaused", "halted": "flowHalted"}
        self._flow_lbl.setText(icons.get(state, state))
        self._flow_lbl.setObjectName(names.get(state, "flowRunning"))
        self._flow_lbl.style().unpolish(self._flow_lbl)
        self._flow_lbl.style().polish(self._flow_lbl)

    def log_event(self, text: str) -> None:
        self._dashboard.append_log(text)


# ── Helpers ───────────────────────────────────────────────────────────────────
def _vline() -> QtWidgets.QFrame:
    f = QtWidgets.QFrame()
    f.setFrameShape(QtWidgets.QFrame.VLine)
    f.setStyleSheet("color: rgba(122,2,25,0.18); margin: 6px 2px;")
    return f


def _stub_dialog(parent: QtWidgets.QWidget, title: str, msg: str) -> None:
    QtWidgets.QMessageBox.information(parent, title, msg)


def main() -> int:
    app = QtWidgets.QApplication(sys.argv)
    apply_liquid_glass_theme(app)
    app.setStyleSheet(app.styleSheet() + _EXTRA_CSS)
    assets_dir = Path(__file__).resolve().parent.parent / "assets"
    set_app_icon(app, "rapid_icon.png", assets_dir)
    window = MainWindow()
    window.show()
    return app.exec()
