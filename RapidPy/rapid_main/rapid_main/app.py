from __future__ import annotations

import sys
import json
import subprocess
from datetime import datetime
from pathlib import Path

from PySide6 import QtCore, QtWidgets

from rapidpy_common.ui import apply_liquid_glass_theme, set_app_icon

from .config import AppConfig
from .hardware_contracts import MeasurementBackend, build_measurement_backend
from .device_ownership import DeviceOwnershipError, DeviceOwnershipManager
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
    CalibrationCenterPanel,
    MeasurementPanel,
    SampleQueuePanel,
    SequencePanel,
    SettingsPanel,
)
from .queue_compiler import QueueCommand, QueueOptions, QueueSample, compile_queue
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
    QLabel#flowIdle {
        background: rgba(107, 114, 128, 0.08);
        border: 1px solid rgba(107, 114, 128, 0.35);
        border-radius: 8px; padding: 3px 10px; color: #4b5563; font-weight: 600;
    }
    QLabel#flowPreflight {
        background: rgba(37, 99, 235, 0.11);
        border: 1px solid rgba(37, 99, 235, 0.35);
        border-radius: 8px; padding: 3px 10px; color: #1d4ed8; font-weight: 600;
    }
    QLabel#flowLoading {
        background: rgba(14, 116, 144, 0.11);
        border: 1px solid rgba(14, 116, 144, 0.32);
        border-radius: 8px; padding: 3px 10px; color: #0f766e; font-weight: 600;
    }
    QLabel#flowTreating {
        background: rgba(120, 53, 15, 0.10);
        border: 1px solid rgba(120, 53, 15, 0.30);
        border-radius: 8px; padding: 3px 10px; color: #7c2d12; font-weight: 600;
    }
    QLabel#flowPositioning {
        background: rgba(109, 40, 217, 0.08);
        border: 1px solid rgba(109, 40, 217, 0.30);
        border-radius: 8px; padding: 3px 10px; color: #6b21a8; font-weight: 600;
    }
    QLabel#flowMeasuring {
        background: rgba(190, 24, 93, 0.08);
        border: 1px solid rgba(190, 24, 93, 0.28);
        border-radius: 8px; padding: 3px 10px; color: #9f1239; font-weight: 600;
    }
    QLabel#flowValidating {
        background: rgba(8, 145, 178, 0.10);
        border: 1px solid rgba(8, 145, 178, 0.35);
        border-radius: 8px; padding: 3px 10px; color: #0e7490; font-weight: 600;
    }
    QLabel#flowSaving {
        background: rgba(20, 83, 45, 0.09);
        border: 1px solid rgba(22, 101, 52, 0.30);
        border-radius: 8px; padding: 3px 10px; color: #166534; font-weight: 600;
    }
    QLabel#flowReturning {
        background: rgba(71, 85, 105, 0.10);
        border: 1px solid rgba(71, 85, 105, 0.28);
        border-radius: 8px; padding: 3px 10px; color: #334155; font-weight: 600;
    }
    QLabel#flowError {
        background: rgba(220, 38, 38, 0.10);
        border: 1px solid rgba(220, 38, 38, 0.30);
        border-radius: 8px; padding: 3px 10px; color: #b91c1c; font-weight: 600;
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
    ("🧪", "Calibration", 5),
]


_DEFAULT_WINDOW_SIZE = (1440, 880)
_DEFAULT_SIDEBAR_WIDTH = 260
_QSETTINGS_ORG = "RAPID"
_QSETTINGS_APP = "RapidPy-rapid_main"
_QSETTINGS_QUEUE_ROWS = "ui/queue_rows"
_QSETTINGS_QUEUE_PROGRESS = "ui/queue_progress"
_QSETTINGS_QUEUE_CURRENT_SAMPLE = "ui/queue_current_sample"
_QSETTINGS_QUEUE_ACTIVE = "ui/queue_active"


class MainWindow(QtWidgets.QMainWindow):

    def __init__(self) -> None:
        super().__init__()
        self.setWindowTitle("RAPID v4 — Paleomagnetics Control System")
        self.resize(*_DEFAULT_WINDOW_SIZE)
        self._settings = QtCore.QSettings(_QSETTINGS_ORG, _QSETTINGS_APP)
        self._nav_btns: list[QtWidgets.QPushButton] = []

        # Persistent configuration (load or create defaults)
        self.config: AppConfig = AppConfig.load()
        self._current_sample = "UNKNOWN"
        self._measurement_backend: MeasurementBackend = build_measurement_backend(self.config)
        self._ownership = DeviceOwnershipManager()

        # Runtime estimator — initialised from config step times
        self._estimator = RuntimeEstimator(self.config.sequence.as_estimator_dict())
        self._sequence_labels: list[str] = []  # current loaded sequence step labels
        self._queue_plan: list[QueueCommand] = []
        self._queue_targets: list[QueueCommand] = []
        self._queue_pos: int = 0
        self._queue_current_sample: str | None = None
        self._queue_active: bool = False
        self._queue_last_warnings: list[str] = []

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
        self._restore_layout_state()

        # Populate settings panel from config
        self._settings_panel.load_from_config(self.config)
        self._nocomm_btn.blockSignals(True)
        self._nocomm_btn.setChecked(self.config.general.nocomm)
        self._nocomm_btn.blockSignals(False)
        self._on_nocomm_toggled(self.config.general.nocomm)

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
        quit_btn.clicked.connect(self._request_shutdown)

        self._pause_btn.clicked.connect(self._on_header_pause)
        self._halt_btn.clicked.connect(self._on_header_halt)
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
        sidebar.setMinimumWidth(240)
        self._sidebar = sidebar
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
            btn.setSizePolicy(
                QtWidgets.QSizePolicy.Policy.Expanding,
                QtWidgets.QSizePolicy.Policy.Fixed,
            )
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
            btn.setSizePolicy(
                QtWidgets.QSizePolicy.Policy.Expanding,
                QtWidgets.QSizePolicy.Policy.Fixed,
            )
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
        self._measurement.sample_run_finished.connect(self._on_queue_sample_finished)
        self._settings_panel    = SettingsPanel()
        self._calibration_panel = CalibrationCenterPanel(
            self, backend_provider=lambda: self.measurement_backend()
        )
        for panel in (
            self._dashboard,
            self._sample_queue,
            self._sequence,
            self._measurement,
            self._settings_panel,
            self._calibration_panel,
        ):
            self._stack.addWidget(panel)

        splitter = QtWidgets.QSplitter(QtCore.Qt.Horizontal)
        splitter.setObjectName("mainSplit")
        splitter.setChildrenCollapsible(False)
        splitter.setHandleWidth(2)
        splitter.addWidget(sidebar)
        splitter.addWidget(self._stack)
        splitter.setStretchFactor(0, 0)
        splitter.setStretchFactor(1, 1)
        self._main_splitter = splitter
        layout.addWidget(splitter)

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

    def start_queue_run(self, samples: list[QueueSample], options: QueueOptions) -> bool:
        """Start a queue-driven measurement run.

        Validation is intentionally strict for safety and queue automation.
        """
        self._queue_active = False
        self._queue_targets = []
        self._queue_plan = []
        self._queue_pos = 0
        self._queue_current_sample = None

        if not samples:
            self.set_status("Queue is empty.")
            return False

        if not self._sequence_labels:
            self.set_status("Load a measurement sequence first.")
            return False

        try:
            self._queue_plan = compile_queue(samples, options, strict=True)
        except Exception as exc:
            self.set_status(f"Queue failed validation: {exc}")
            QtWidgets.QMessageBox.critical(
                self,
                "Queue Error",
                f"Queue failed validation:\n\n{exc}",
            )
            return False

        self._queue_targets = [cmd for cmd in self._queue_plan if cmd.command_type == "Meas"]
        if not self._queue_targets:
            self.set_status("Queue has no measurement commands.")
            QtWidgets.QMessageBox.information(self, "Queue", "Queue has no measurement steps.")
            return False

        self._queue_active = True
        self._queue_pos = 0
        self._queue_last_warnings = []
        self._queue_current_sample = None
        self.set_status(f"Queue run started with {len(self._queue_targets)} measurements.")
        self._save_queue_state()
        self._run_next_queue_sample()
        return True

    def cancel_queue_run(self, reason: str = "Queue cancelled.") -> None:
        """Cancel active queue automation while leaving controls in a safe state."""
        if not self._queue_active:
            return
        if hasattr(self._measurement, "halt_run"):
            self._measurement.halt_run()
        self._queue_active = False
        self._queue_plan = []
        self._queue_targets = []
        self._queue_pos = 0
        self._queue_current_sample = None
        self._save_queue_state()
        self.set_status(reason)
        self.set_flow_state("idle")

    def _run_next_queue_sample(self) -> None:
        if not self._queue_active:
            return
        if self._queue_pos >= len(self._queue_targets):
            self._queue_active = False
            self._queue_current_sample = None
            self._save_queue_state()
            self.set_status("Queue run complete.")
            self.set_flow_state("complete")
            return
        target = self._queue_targets[self._queue_pos]
        self._queue_pos += 1
        self._queue_current_sample = target.sample_name
        self._sample_queue.start_queue_sample(target.sample_name)
        if not self._measurement.start_measurement_for_sample(target.sample_name):
            self._queue_active = False
            self._sample_queue.set_queue_sample_failed(target.sample_name)
            self._queue_current_sample = None
            self._save_queue_state()
            self.set_status("Queue run failed to start.")
            self.set_flow_state("error")
            return
        self._save_queue_state()

    def _on_queue_sample_finished(self, aborted: bool, sample: str) -> None:
        if not self._queue_active:
            return
        had_error = (
            self._measurement.take_last_run_error()
            if hasattr(self._measurement, "take_last_run_error")
            else False
        )
        if aborted:
            self._queue_active = False
            self._sample_queue.set_queue_sample_failed(sample)
            self._queue_current_sample = None
            self._save_queue_state()
            self.set_status(f"Queue stopped after sample {sample}.")
            return
        if had_error:
            self._queue_active = False
            self._sample_queue.set_queue_sample_failed(sample)
            self._queue_current_sample = None
            self._save_queue_state()
            self.set_status(f"Queue sample {sample} failed with an error.")
            return
        self._sample_queue.mark_queue_sample_done(sample)
        self._save_queue_state()
        self._run_next_queue_sample()

    def measurement_backend(self) -> MeasurementBackend:
        """Return the active workflow backend."""
        return self._measurement_backend

    def acquire_measurement_device(self, owner: str, *, allow_reentrant: bool = True) -> object:
        """Acquire the shared measurement device lease for the provided owner.

        Callers hold the returned lease and must call ``release()`` after completion.
        """
        return self.acquire_device("measurement", owner, allow_reentrant=allow_reentrant)

    def acquire_device(
        self,
        resource: str,
        owner: str,
        *,
        allow_reentrant: bool = True,
    ) -> object:
        """Acquire a shared device lease used to prevent concurrent hardware ownership."""
        return self._ownership.acquire(resource, owner, allow_reentrant=allow_reentrant)

    def release_measurement_device(self, owner: str) -> None:
        """Release measurement ownership for this owner when no longer running."""
        self._ownership.release("measurement", owner)

    def release_device(self, resource: str, owner: str) -> None:
        """Release a shared device lease."""
        self._ownership.release(resource, owner)

    def set_current_sample(self, name: str) -> None:
        """Update the app-level current sample context."""
        self._current_sample = name or "UNKNOWN"
        self.set_sample(self._current_sample)

    def start_run(self, labels: list[str] | None = None) -> None:
        """Begin a measurement run — start the live countdown timer."""
        if labels is not None:
            self._sequence_labels = list(labels)
        self._run_current_idx = 0
        self._run_start_time = datetime.now()
        self._run_timer.start()
        self._update_runtime_display(current_idx=0, running=True)
        self.set_flow_state("running")

    def stop_run(self) -> None:
        """End the measurement run — stop the countdown timer."""
        self._run_timer.stop()
        self._run_start_time = None
        self._run_current_idx = 0
        self._update_runtime_display(running=False)
        self.set_flow_state("idle")

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
        fm.addAction("E&xit", self._request_shutdown)

        vm = mb.addMenu("&View")
        for label, idx in [("&Dashboard", 0), ("Sample &Queue", 1),
                            ("Se&quence", 2), ("Live &Measurement", 3),
                            ("&Settings", 4), ("&Calibration", 5)]:
            vm.addAction(label, lambda _c=False, i=idx: self._nav_select(i))
        vm.addSeparator()
        vm.addAction("Step &Monitor", self._launch_step_monitor)
        vm.addAction("&Data Review", self._launch_data_viewer)
        vm.addAction("&Debug Console", self._launch_debug_console)
        vm.addAction("&Sample Queue Monitor")
        vm.addAction("&Reset Layout", self._reset_layout)
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
        af_sub.addAction("Run AF Demo Sequence", self._launch_af_demo)
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
        self.config.general.nocomm = bool(on)
        self.config.save()
        self._measurement_backend = build_measurement_backend(self.config)
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
        self._run_owned_stub_dialog(
            "dc_motors",
            "dc_motors_panel",
            "DC Motors",
            "Launches dc_motor_control app (Phase 3)",
        )

    @staticmethod
    def _af_demo_labels() -> list[str]:
        """Canonical AF demo sequence labels."""
        return ["NRM", "AF25", "AF50", "AF100", "AF200", "AF400", "AF800", "SUSC"]

    def _launch_af(self) -> None:
        self._prepare_af_workflow(auto_start=False)

    def _launch_af_demo(self) -> None:
        """Run the AF demo sequence directly from the diagnostics launcher."""
        self._prepare_af_workflow(auto_start=True)

    def _prepare_af_workflow(self, *, auto_start: bool) -> bool:
        """Load AF sequence defaults and optionally auto-start the run."""
        self.load_sequence_labels(self._af_demo_labels())
        self.set_current_sample("AF_DEMO")
        if hasattr(self._measurement, "set_specimen_context"):
            self._measurement.set_specimen_context(
                sample="AF_DEMO",
                depth="—",
                treatment="AF workflow",
            )
        self.set_status("AF workflow loaded in Live Measurement.")
        self._nav_select(3)
        if not auto_start:
            return True

        started = bool(
            hasattr(self._measurement, "start_measurement_for_sample")
            and self._measurement.start_measurement_for_sample("AF_DEMO")
        )
        if not started:
            self.set_status("Unable to auto-start AF demo. Manual start is available in Live Measurement.")
            return False

        self.set_status(
            "AF demo sequence started. Use Pause/Resume and Halt in Live Measurement to control."
        )
        return True

    def _launch_irm(self) -> None:
        self._run_owned_dialog("irm", "irm_panel", IrmArmDialog)

    def _launch_vacuum(self) -> None:
        self._run_owned_dialog("vacuum", "vacuum_panel", VacuumDialog)

    def _launch_squid(self) -> None:
        self._run_owned_dialog("squid", "squid_panel", SquidCommDialog)

    def _run_owned_dialog(
        self,
        resource: str,
        owner: str,
        dialog_factory,
        *,
        modal: bool = True,
    ) -> None:
        """Open a dialog while owning one hardware resource.

        This prevents multiple widgets from attempting to control the same
        subsystem at the same time.
        """
        lease = None
        try:
            lease = self.acquire_device(resource, owner)
        except DeviceOwnershipError as exc:
            QtWidgets.QMessageBox.warning(self, "Device Busy", str(exc))
            return

        try:
            dlg = dialog_factory(self)
            if modal:
                dlg.exec()
            else:
                dlg.show()
        finally:
            if lease is not None:
                lease.release()

    def _run_owned_stub_dialog(
        self,
        resource: str,
        owner: str,
        title: str,
        message: str,
        *,
        modal: bool = True,
    ) -> None:
        """Open a message-style action while owning a shared hardware resource."""

        class _StubDialog:
            def exec(self_inner) -> None:
                _stub_dialog(self, title, message)

            def show(self_inner) -> None:
                _stub_dialog(self, title, message)

        self._run_owned_dialog(
            resource,
            owner,
            lambda _parent: _StubDialog(),
            modal=modal,
        )

    def closeEvent(self, event) -> None:
        """Confirm safe shutdown and persist layout settings before closing."""
        if not self._confirm_shutdown(prompt=True, on_close=True):
            event.ignore()
            return

        self._save_layout_state()
        super().closeEvent(event)

    def _has_active_automation(self) -> bool:
        """Return true when a live measurement or queue run is active."""
        measurement_active = bool(
            hasattr(self._measurement, "is_active") and self._measurement.is_active()
        )
        return measurement_active or bool(self._queue_active)

    def _confirm_shutdown(self, *, prompt: bool = True, on_close: bool = False) -> bool:
        """Prompt the operator to confirm shutdown and safely halt active workflow."""
        if on_close and self._has_active_automation():
            if (
                QtWidgets.QMessageBox.question(
                    self,
                    "Active run in progress",
                    (
                        "A measurement or queue run is currently active.\n"
                        "Shutdown will halt the run and stop queue execution.\n\n"
                        "Continue?"
                    ),
                    QtWidgets.QMessageBox.StandardButton.Yes
                    | QtWidgets.QMessageBox.StandardButton.No,
                    QtWidgets.QMessageBox.StandardButton.No,
                )
                != QtWidgets.QMessageBox.StandardButton.Yes
            ):
                return False
        elif on_close and prompt:
            if (
                QtWidgets.QMessageBox.question(
                    self,
                    "Exit RAPID",
                    "Exit RAPID now?",
                    QtWidgets.QMessageBox.StandardButton.Yes
                    | QtWidgets.QMessageBox.StandardButton.No,
                    QtWidgets.QMessageBox.StandardButton.No,
                )
                != QtWidgets.QMessageBox.StandardButton.Yes
            ):
                return False

        if self._has_active_automation():
            self.halt_measurement()

        if prompt and not on_close:
            if (
                QtWidgets.QMessageBox.question(
                    self,
                    "Shutdown requested",
                    "Stop and close RAPID?",
                    QtWidgets.QMessageBox.StandardButton.Yes
                    | QtWidgets.QMessageBox.StandardButton.No,
                    QtWidgets.QMessageBox.StandardButton.Yes,
                )
                != QtWidgets.QMessageBox.StandardButton.Yes
            ):
                return False

        return True

    def _request_shutdown(self) -> None:
        """Run shutdown flow from menu and toolbar actions."""
        if self._confirm_shutdown():
            self.close()

    def _restore_layout_state(self) -> None:
        geometry = self._settings.value("ui/window_geometry")
        if geometry:
            self.restoreGeometry(geometry)
        else:
            self.resize(*_DEFAULT_WINDOW_SIZE)

        panel_index = self._settings.value("ui/active_panel", 0, type=int)
        self._nav_select(panel_index or 0)

        sidebar_width = self._settings.value("ui/sidebar_width", _DEFAULT_SIDEBAR_WIDTH, type=int)
        splitter_state = self._settings.value("ui/main_splitter_state")

        if splitter_state:
            self._main_splitter.restoreState(splitter_state)
            if self._sidebar.width() < _DEFAULT_SIDEBAR_WIDTH:
                self._main_splitter.setSizes(
                    [_DEFAULT_SIDEBAR_WIDTH, max(1, self.width() - _DEFAULT_SIDEBAR_WIDTH)]
                )
        elif isinstance(sidebar_width, int) and sidebar_width >= 120:
            self._main_splitter.setSizes([sidebar_width, max(1, self.width() - sidebar_width)])
        else:
            self._main_splitter.setSizes(
                [_DEFAULT_SIDEBAR_WIDTH, max(1, self.width() - _DEFAULT_SIDEBAR_WIDTH)]
            )
        self._restore_queue_state()

    def _save_layout_state(self) -> None:
        self._settings.setValue("ui/window_geometry", self.saveGeometry())
        self._settings.setValue("ui/active_panel", self._stack.currentIndex())
        self._settings.setValue("ui/main_splitter_state", self._main_splitter.saveState())
        self._settings.setValue("ui/sidebar_width", self._sidebar.width())
        self._save_queue_state()

    def _save_queue_state(self) -> None:
        try:
            queue_rows = self._sample_queue.row_snapshot()
            self._settings.setValue(_QSETTINGS_QUEUE_ROWS, json.dumps(queue_rows))
            self._settings.setValue(_QSETTINGS_QUEUE_PROGRESS, int(self._queue_pos))
            self._settings.setValue(
                _QSETTINGS_QUEUE_CURRENT_SAMPLE, self._queue_current_sample or ""
            )
            self._settings.setValue(_QSETTINGS_QUEUE_ACTIVE, bool(self._queue_active))
        except Exception:
            self._settings.remove(_QSETTINGS_QUEUE_ROWS)
            self._settings.remove(_QSETTINGS_QUEUE_PROGRESS)
            self._settings.remove(_QSETTINGS_QUEUE_CURRENT_SAMPLE)
            self._settings.remove(_QSETTINGS_QUEUE_ACTIVE)

    def _restore_queue_state(self) -> None:
        raw = self._settings.value(_QSETTINGS_QUEUE_ROWS)
        if not raw:
            self._queue_current_sample = None
            return
        try:
            rows = json.loads(raw) if isinstance(raw, str) else raw
            if isinstance(rows, list):
                self._sample_queue.load_rows(rows)
                was_active = bool(
                    self._settings.value(_QSETTINGS_QUEUE_ACTIVE, False, type=bool)
                )
                recovered = self._sample_queue.recover_interrupted_samples()
                if recovered:
                    self.set_status(
                        f"Recovered {recovered} interrupted queue row(s) from last session."
                    )
                elif was_active:
                    self.set_status(
                        "Previous queue run was active; review pending rows and press Run Queue to continue."
                    )
        except Exception:
            return

    def _reset_layout(self) -> None:
        self._settings.remove("ui/window_geometry")
        self._settings.remove("ui/active_panel")
        self._settings.remove("ui/sidebar_width")
        self._settings.remove("ui/main_splitter_state")
        self._settings.remove(_QSETTINGS_QUEUE_ROWS)
        self._settings.remove(_QSETTINGS_QUEUE_PROGRESS)
        self._settings.remove(_QSETTINGS_QUEUE_CURRENT_SAMPLE)
        self._settings.remove(_QSETTINGS_QUEUE_ACTIVE)
        self.resize(*_DEFAULT_WINDOW_SIZE)
        self._main_splitter.setSizes(
            [_DEFAULT_SIDEBAR_WIDTH, max(1, self.width() - _DEFAULT_SIDEBAR_WIDTH)]
        )
        self._nav_select(0)
        QtWidgets.QMessageBox.information(self, "Reset Layout", "Layout reset to defaults.")

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

    def _launch_data_viewer(self) -> None:
        data_viewer_script = Path(__file__).resolve().parents[2] / "data_viewer" / "main.py"
        if not data_viewer_script.exists():
            QtWidgets.QMessageBox.warning(
                self,
                "Data Review",
                "Data viewer launcher script is missing.",
            )
            return

        if (
            hasattr(self, "_data_viewer_proc")
            and getattr(self, "_data_viewer_proc")
            and self._data_viewer_proc.poll() is None
        ):
            QtWidgets.QMessageBox.information(
                self,
                "Data Review",
                "Data Review window is already running.",
            )
            return

        try:
            self._data_viewer_proc = subprocess.Popen(
                [sys.executable, str(data_viewer_script)],
                cwd=str(data_viewer_script.parent),
            )
        except OSError as exc:
            QtWidgets.QMessageBox.warning(
                self,
                "Data Review",
                f"Unable to launch data review utility:\n{exc}",
            )
            self._data_viewer_proc = None
            return

        self.set_status("Data Review launched in a separate window.")

    def _launch_webcam(self) -> None:
        if not hasattr(self, "_webcam_dlg"):
            self._webcam_dlg = WebcamDialog(self)
        self._webcam_dlg.show()
        self._webcam_dlg.raise_()

    # ── Public API for panels ─────────────────────────────────────────────────
    def navigate_to(self, key: str) -> None:
        mapping = {
            "dashboard": 0,
            "queue": 1,
            "sequence": 2,
            "measure": 3,
            "settings": 4,
            "calibration": 5,
        }
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
        """Update the top-of-screen workflow label from a phase name."""
        icons = {
            "running": "◉  Running",
            "paused": "⏸  Paused",
            "halted": "■  Halted",
            "idle": "◎  Idle",
            "preflight": "⏳  Preflight",
            "loading": "⤴  Loading",
            "treating": "⚙  Treating",
            "positioning": "🎯  Positioning",
            "measuring": "📈  Measuring",
            "validating": "✓  Validating",
            "saving": "💾  Saving",
            "returning": "↩  Returning",
            "complete": "✅  Complete",
            "error": "⚠  Error",
        }
        names = {
            "running": "flowRunning",
            "paused": "flowPaused",
            "halted": "flowHalted",
            "idle": "flowIdle",
            "preflight": "flowPreflight",
            "loading": "flowLoading",
            "treating": "flowTreating",
            "positioning": "flowPositioning",
            "measuring": "flowMeasuring",
            "validating": "flowValidating",
            "saving": "flowSaving",
            "returning": "flowReturning",
            "complete": "flowRunning",
            "error": "flowError",
        }
        self._flow_lbl.setText(icons.get(state, state))
        self._flow_lbl.setObjectName(names.get(state, "flowRunning"))
        self._flow_lbl.style().unpolish(self._flow_lbl)
        self._flow_lbl.style().polish(self._flow_lbl)

    def toggle_queue_pause(self) -> None:
        """Pause or resume active measurement automation."""
        if hasattr(self._measurement, "toggle_pause"):
            self._measurement.toggle_pause()
        elif hasattr(self._measurement, "pause_run"):
            self._measurement.pause_run()

    def halt_measurement(self) -> None:
        """Stop active measurement and queue automation in a single action."""
        if hasattr(self._measurement, "halt_run"):
            self._measurement.halt_run()
        self.cancel_queue_run("Queue halted by user.")

    def _on_header_pause(self) -> None:
        if hasattr(self._measurement, "is_active") and self._measurement.is_active():
            self.toggle_queue_pause()
        else:
            self.set_status("No active measurement to pause.")

    def _on_header_halt(self) -> None:
        if hasattr(self._measurement, "is_active") and self._measurement.is_active():
            self.halt_measurement()
            self.set_flow_state("halted")
        elif self._queue_active:
            self.halt_measurement()
        else:
            self.set_status("No active run to halt.")

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
    set_app_icon(window, "rapid_icon.png", assets_dir)
    window.show()
    return app.exec()
