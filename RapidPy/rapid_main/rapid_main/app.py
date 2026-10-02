from __future__ import annotations

import sys
import json
import os
import subprocess
from datetime import datetime
from pathlib import Path

from PySide6 import QtCore, QtGui, QtWidgets

from rapidpy_common.ui import (
    MIN_WINDOW_WIDTH,
    apply_liquid_glass_theme,
    apply_window_bounds_guard,
    clamp_window_geometry,
    _screen_area_for_widget,
    set_app_icon,
)

from .config import AppConfig
from .hardware_contracts import (
    MeasurementAutomationBackend,
    build_measurement_backend,
)
from .diagnostic_services import (
    build_af_demag_backend,
    build_backend_or_unavailable,
    build_dcmotor_backend,
    build_irm_arm_backend,
    build_squid_backend,
    build_susceptibility_backend,
    build_vacuum_backend,
    collect_diagnostic_status,
    require_squid_ready,
    require_vacuum_ready,
)
from .device_ownership import DeviceOwnershipError, DeviceOwnershipManager
from .dialogs import (
    AboutDialog,
    DebugConsoleDialog,
    IrmArmDialog,
    LoginDialog,
    PlotsDialog,
    SampleSelectDialog,
    SquidCommDialog,
    SusceptibilityDialog,
    StartupGuideDialog,
    TransitionHelpDialog,
    DCMotorDialog,
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
from .package_launch import ToolUnavailableError, resolve_tool_launch
from .runtime_estimator import RuntimeEstimator
from .vrm import VRM_CONTEXT_ENV, build_vrm_launch_context, write_vrm_launch_context
from .glass_theme import (
    GlassBackdrop,
    apply_main_glass_theme,
    install_glass_elevation,
)
from .startup import main_assets_dir, select_main_icon


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
        padding: 2px 2px;
        text-align: left;
        color: #4d3a39;
        font-size: 11px;
        font-weight: 500;
        min-width: 0;
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
    QLabel#flowComplete {
        background: rgba(34, 197, 94, 0.14);
        border: 1px solid rgba(22, 101, 52, 0.40);
        border-radius: 8px; padding: 3px 10px; color: #15803d; font-weight: 600;
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
    QLabel#flowOverride {
        background: rgba(71, 85, 105, 0.15); border: 1px dashed rgba(71, 85, 105, 0.55);
        border-radius: 8px; padding: 3px 10px; color: #334155; font-weight: 700;
    }
    QLabel#instOk {
        background: rgba(34,197,94,0.12); border: 1px solid rgba(34,197,94,0.35);
        border-radius: 7px; padding: 2px 9px; color: #15803d; font-size: 12px;
    }
    QLabel#instErr {
        background: rgba(220,38,38,0.10); border: 1px solid rgba(220,38,38,0.35);
        border-radius: 7px; padding: 2px 9px; color: #b91c1c; font-size: 12px;
    }
    QLabel#instSim {
        background: rgba(245,158,11,0.13); border: 1px solid rgba(217,119,6,0.38);
        border-radius: 7px; padding: 2px 9px; color: #92400e; font-size: 12px;
        font-weight: 650;
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


_DEFAULT_WINDOW_SIZE = (1320, 820)
_DEFAULT_SIDEBAR_WIDTH = 252
_MIN_SIDEBAR_WIDTH = 240
_MAX_SIDEBAR_WIDTH = 288
_SIDEBAR_RESTORE_RATIO = 0.25
_MAIN_MAX_WIDTH_RATIO = 0.94
_MAIN_MAX_WIDTH = 1680
_MAIN_MAX_HEIGHT_RATIO = 0.92
_MAIN_MAX_HEIGHT = 1100
_MAIN_MIN_WIDTH = 900
_MAIN_MIN_HEIGHT = 640
_MAIN_FIXED_MIN_WIDTH_THRESHOLD = 1024
_QSETTINGS_ORG = "RAPID"
_QSETTINGS_APP = "RapidPy-rapid_main"
_QSETTINGS_QUEUE_ROWS = "ui/queue_rows"
_QSETTINGS_QUEUE_PROGRESS = "ui/queue_progress"
_QSETTINGS_QUEUE_RESUME_POS = "ui/queue_resume_pos"
_QSETTINGS_QUEUE_CURRENT_SAMPLE = "ui/queue_current_sample"
_QSETTINGS_QUEUE_ACTIVE = "ui/queue_active"
_QSETTINGS_SHOW_STARTUP_GUIDE = "ui/show_startup_guide"


def clamp_window_size_for_screen(available: QtCore.QRect, requested: tuple[int, int]) -> tuple[int, int]:
    """Backward-compatible alias retained for legacy tests/callers.

    Rapidly moved into :func:`rapidpy_common.ui.clamp_window_geometry`.
    """
    return clamp_window_geometry(available, requested)


def _clamp_main_window_size(available: QtCore.QRect, requested: tuple[int, int]) -> tuple[int, int]:
    requested_width, requested_height = requested
    available_width = max(1, int(available.width()))
    available_height = max(1, int(available.height()))
    width_cap = min(
        available_width,
        _MAIN_MAX_WIDTH,
        max(1, int(available_width * _MAIN_MAX_WIDTH_RATIO)),
    )
    height_cap = min(
        available_height,
        _MAIN_MAX_HEIGHT,
        max(1, int(available_height * _MAIN_MAX_HEIGHT_RATIO)),
    )
    tiny_profile = available_width < _MAIN_FIXED_MIN_WIDTH_THRESHOLD
    if tiny_profile:
        minimum_width = min(width_cap, max(MIN_WINDOW_WIDTH, int(available_width * 0.82)))
    else:
        minimum_width = min(width_cap, _MAIN_MIN_WIDTH)
    minimum_height = min(height_cap, _MAIN_MIN_HEIGHT)
    requested_width = max(1, int(requested_width))
    requested_height = max(1, int(requested_height))
    width = max(minimum_width, min(requested_width, width_cap))
    height = max(minimum_height, min(requested_height, height_cap))
    return min(width, available_width), min(height, available_height)


def clamp_sidebar_target(
    requested: int,
    *,
    window_width: int,
    minimum: int,
    maximum: int,
    ratio: float,
) -> int:
    """Return a sidebar width bounded by explicit limits and fractional window width."""
    cap = max(minimum, int(max(1, window_width) * ratio))
    return max(minimum, min(requested, maximum, cap))


def _sidebar_button_label(icon: str, label: str) -> str:
    return f"  {icon}  {label}"


def _allow_horizontal_compression(widget: QtWidgets.QWidget) -> None:
    widget.setMinimumWidth(0)
    widget.setSizePolicy(
        QtWidgets.QSizePolicy.Policy.Ignored,
        widget.sizePolicy().verticalPolicy(),
    )


class _AdaptivePanelStack(QtWidgets.QStackedWidget):
    def sizeHint(self) -> QtCore.QSize:
        return super().sizeHint()

    def minimumSizeHint(self) -> QtCore.QSize:
        hint = super().minimumSizeHint()
        # Hidden pages must not force an oversized startup width, while their
        # vertical minimum remains available to the window layout.
        screen = self.screen()
        app = QtWidgets.QApplication.instance()
        if screen is None and app is not None:
            screen = app.primaryScreen()
        max_height = hint.height()
        if screen is not None:
            max_height = min(max_height, screen.availableGeometry().height())
        return QtCore.QSize(0, max_height)


class MainWindow(QtWidgets.QMainWindow):

    def __init__(self) -> None:
        super().__init__()
        self.setObjectName("rapidMainWindow")
        self.setWindowTitle("RAPID v4 — Paleomagnetics Control System")
        self.resize(*_DEFAULT_WINDOW_SIZE)
        self._sidebar_min_width = _MIN_SIDEBAR_WIDTH
        self._sidebar_default_width = _DEFAULT_SIDEBAR_WIDTH
        self._settings = QtCore.QSettings(_QSETTINGS_ORG, _QSETTINGS_APP)
        self._startup_guide_dialog: StartupGuideDialog | None = None
        self._startup_guide_scheduled = False
        self._nav_btns: list[QtWidgets.QPushButton] = []

        # Persistent configuration (load or create defaults)
        self.config: AppConfig = AppConfig.load()
        self._current_sample = "UNKNOWN"
        self.sample_registrations = None
        self._measurement_backend: MeasurementAutomationBackend = build_measurement_backend(self.config)
        self._rebuild_diagnostic_backends(nocomm=bool(self.config.general.nocomm))
        self._ownership = DeviceOwnershipManager()
        self._owned_dialog_leases: dict[str, object] = {}

        # Runtime estimator — initialised from config step times
        self._estimator = RuntimeEstimator(self.config.sequence.as_estimator_dict())
        self._sequence_labels: list[str] = []  # current loaded sequence step labels
        self._queue_plan: list[QueueCommand] = []
        self._queue_pos: int = 0
        self._queue_resume_pos: int = 0
        self._queue_current_sample: str | None = None
        self._queue_current_command: QueueCommand | None = None
        self._queue_active: bool = False
        self._queue_last_warnings: list[str] = []
        self._queue_paused: bool = False
        self._queue_lease: object | None = None
        self._workflow_state = "idle"
        self._current_step = "—"
        self._current_treatment = "—"
        self._status_override_active = False

        # Live countdown timer (1 Hz, used when a run is active)
        self._run_start_time: "datetime | None" = None
        self._run_current_idx: int = 0
        self._run_timer = QtCore.QTimer(self)
        self._run_timer.setInterval(1000)
        self._run_timer.timeout.connect(self._tick_run_countdown)

        self._build_header()
        self._build_central()
        QtCore.QTimer.singleShot(0, lambda: install_glass_elevation(self))
        self._build_statusbar()
        self._build_menu()
        self._restore_layout_state()

        # Populate settings panel from config
        self._settings_panel.load_from_config(self.config)
        self._nocomm_btn.blockSignals(True)
        self._nocomm_btn.setChecked(self.config.general.nocomm)
        self._nocomm_btn.blockSignals(False)
        self._on_nocomm_toggled(self.config.general.nocomm)

        self._diagnostic_timer = QtCore.QTimer(self)
        self._diagnostic_timer.setInterval(10_000)
        self._diagnostic_timer.timeout.connect(self._refresh_dashboard_diagnostics)
        self._diagnostic_timer.start()
        QtCore.QTimer.singleShot(0, self._refresh_dashboard_diagnostics)

        self._clock = QtCore.QTimer(self)
        self._clock.timeout.connect(self._tick_clock)
        self._clock.start(1000)
        self._fit_window_to_current_screen()
        QtCore.QTimer.singleShot(0, self._fit_window_to_current_screen)
        QtCore.QTimer.singleShot(0, self._wire_screen_guard)


    # ── Header toolbar ────────────────────────────────────────────────────────
    def _build_header(self) -> None:
        header = QtWidgets.QFrame()
        header.setObjectName("header")
        header.setFixedHeight(54)
        _allow_horizontal_compression(header)

        hl = QtWidgets.QHBoxLayout(header)
        hl.setContentsMargins(18, 0, 14, 0)
        hl.setSpacing(10)

        title = QtWidgets.QLabel("⚗  RAPID v4")
        title.setObjectName("headerTitle")
        title.setToolTip("RAPID v4")
        _allow_horizontal_compression(title)
        hl.addWidget(title)
        hl.addWidget(_vline())

        self._flow_lbl = QtWidgets.QLabel("◎  Idle")
        self._flow_lbl.setObjectName("flowIdle")
        self._flow_lbl.setToolTip("Run state")
        _allow_horizontal_compression(self._flow_lbl)
        hl.addWidget(self._flow_lbl)
        hl.addWidget(_vline())

        self._sample_hdr = QtWidgets.QLabel("Sample: —")
        self._sample_hdr.setStyleSheet("color: #4d3a39; font-size: 13px;")
        self._sample_hdr.setToolTip("Current sample")
        _allow_horizontal_compression(self._sample_hdr)
        hl.addWidget(self._sample_hdr)

        self._step_hdr = QtWidgets.QLabel("Step: —")
        self._step_hdr.setStyleSheet("color: #7a6f6e; font-size: 12px;")
        self._step_hdr.setToolTip("Current step")
        _allow_horizontal_compression(self._step_hdr)
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
        for btn, tip, width in (
            (self._pause_btn, "Pause the active queue or measurement", 68),
            (self._halt_btn, "Halt the active queue or measurement", 58),
            (self._nocomm_btn, "Toggle no-communication simulation mode", 76),
            (quit_btn, "Exit RAPID", 54),
        ):
            btn.setToolTip(tip)
            btn.setMinimumWidth(0)
            btn.setMaximumWidth(width)
            hl.addWidget(btn)

        tb = QtWidgets.QToolBar()
        tb.setMovable(False)
        tb.setFloatable(False)
        tb.setStyleSheet("QToolBar { border: none; padding: 0; margin: 0; }")
        tb.setMinimumWidth(0)
        tb.addWidget(header)
        self.addToolBar(QtCore.Qt.TopToolBarArea, tb)

    # ── Central widget: sidebar + stacked panels ──────────────────────────────
    def _build_central(self) -> None:
        root = GlassBackdrop()
        layout = QtWidgets.QHBoxLayout(root)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(0)
        self.setCentralWidget(root)
        sidebar_fit_width = _DEFAULT_SIDEBAR_WIDTH

        # ── Sidebar ──
        sidebar = QtWidgets.QFrame()
        sidebar.setObjectName("sidebar")
        sidebar.setMinimumWidth(_MIN_SIDEBAR_WIDTH)
        sidebar.setMaximumWidth(_MAX_SIDEBAR_WIDTH)
        sidebar.setContentsMargins(0, 0, 0, 0)
        sidebar.setSizePolicy(
            QtWidgets.QSizePolicy.Policy.Fixed,
            QtWidgets.QSizePolicy.Policy.Expanding,
        )
        self._sidebar = sidebar
        sl = QtWidgets.QVBoxLayout(sidebar)
        sl.setContentsMargins(2, 6, 2, 6)
        sl.setSpacing(3)

        def _sec_hdr(text: str) -> QtWidgets.QLabel:
            lbl = QtWidgets.QLabel(text)
            lbl.setObjectName("sectionHdr")
            lbl.setContentsMargins(8, 0, 0, 4)
            return lbl

        sl.addWidget(_sec_hdr("MAIN"))
        self._btn_group = QtWidgets.QButtonGroup(self)
        self._btn_group.setExclusive(True)
        sidebar_buttons: list[QtWidgets.QPushButton] = []
        for icon, label, idx in _NAV_ITEMS:
            btn = QtWidgets.QPushButton(_sidebar_button_label(icon, label))
            btn.setObjectName("navBtn")
            btn.setCheckable(True)
            btn.setSizePolicy(
                QtWidgets.QSizePolicy.Policy.Expanding,
                QtWidgets.QSizePolicy.Policy.Fixed,
            )
            btn.setMinimumHeight(40)
            btn.setToolTip(label)
            sidebar_fit_width = max(sidebar_fit_width, btn.sizeHint().width() + 20)
            btn.clicked.connect(lambda _checked, i=idx: self._nav_select(i))
            self._btn_group.addButton(btn)
            self._nav_btns.append(btn)
            sl.addWidget(btn)
            sidebar_buttons.append(btn)

        sl.addSpacing(16)
        sl.addWidget(_sec_hdr("DIAGNOSTICS"))

        for icon, label, slot in [
            ("🔌", "DC Motors",  self._launch_dc_motors),
            ("🌊", "AF Demag",   self._launch_af),
            ("🧲", "IRM / ARM",  self._launch_irm),
            ("💧", "Vacuum",     self._launch_vacuum),
            ("🔭", "SQUID Comm", self._launch_squid),
        ]:
            btn = QtWidgets.QPushButton(_sidebar_button_label(icon, label))
            btn.setObjectName("navBtn")
            btn.setSizePolicy(
                QtWidgets.QSizePolicy.Policy.Expanding,
                QtWidgets.QSizePolicy.Policy.Fixed,
            )
            btn.setMinimumHeight(40)
            btn.setToolTip(label)
            sidebar_fit_width = max(sidebar_fit_width, btn.sizeHint().width() + 16)
            btn.clicked.connect(slot)
            sl.addWidget(btn)
            sidebar_buttons.append(btn)
        sidebar_fit_width = max(_MIN_SIDEBAR_WIDTH, sidebar_fit_width)
        sidebar_fit_width = min(_MAX_SIDEBAR_WIDTH, sidebar_fit_width)
        sidebar.setMinimumWidth(_MIN_SIDEBAR_WIDTH)
        self._sidebar_min_width = _MIN_SIDEBAR_WIDTH
        self._sidebar_default_width = max(
            self._sidebar_min_width,
            min(
                _MAX_SIDEBAR_WIDTH,
                clamp_sidebar_target(
                    sidebar_fit_width,
                    window_width=max(1, self.width()),
                    minimum=self._sidebar_min_width,
                    maximum=_MAX_SIDEBAR_WIDTH,
                    ratio=_SIDEBAR_RESTORE_RATIO,
                ),
            ),
        )
        for btn in sidebar_buttons:
            btn.setSizePolicy(
                QtWidgets.QSizePolicy.Policy.Expanding,
                QtWidgets.QSizePolicy.Policy.Fixed,
            )
            btn.setContentsMargins(0, 0, 0, 0)

        sl.addStretch()
        ver = QtWidgets.QLabel("RAPID v4.0 · Phase 2")
        ver.setStyleSheet("color: #c4b7b3; font-size: 10px; padding: 0 8px;")
        sl.addWidget(ver)

        # ── Stacked panels ──
        self._stack = _AdaptivePanelStack()
        self._stack.currentChanged.connect(self._stack.updateGeometry)
        self._stack.setSizePolicy(
            QtWidgets.QSizePolicy.Policy.Expanding,
            QtWidgets.QSizePolicy.Policy.Expanding,
        )
        self._stack.setMinimumSize(0, 0)
        self._dashboard   = DashboardPanel()
        self._dashboard.refresh_diagnostics_requested.connect(
            self._refresh_dashboard_diagnostics
        )
        self._dashboard.load_sample_requested.connect(self._load_sample_from_index)
        self._sample_queue = SampleQueuePanel()
        self._sample_queue.sample_index_requested.connect(self._load_sample_from_index)
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
            panel.setMinimumWidth(0)
            panel.setMinimumHeight(0)
            panel.setSizePolicy(
                QtWidgets.QSizePolicy.Policy.Expanding,
                QtWidgets.QSizePolicy.Policy.Expanding,
            )
            panel.setMinimumSize(0, 0)
            self._stack.addWidget(panel)
        self._sync_dashboard_run_state()

        splitter = QtWidgets.QSplitter(QtCore.Qt.Horizontal)
        splitter.setObjectName("mainSplit")
        splitter.setChildrenCollapsible(False)
        splitter.setHandleWidth(2)
        splitter.addWidget(sidebar)
        splitter.addWidget(self._stack)
        splitter.setStretchFactor(0, 0)
        splitter.setStretchFactor(1, 1)
        self._main_splitter = splitter
        self._main_splitter.splitterMoved.connect(self._on_splitter_moved)
        layout.addWidget(splitter)

    def _current_screen(self) -> QtCore.QRect | None:
        return _screen_area_for_widget(self)

    def _resolved_work_area(self) -> QtCore.QRect:
        available = self._current_screen()
        if available is not None and available.width() > 0 and available.height() > 0:
            return available

        screen = QtWidgets.QApplication.primaryScreen()
        if screen is not None:
            primary = screen.availableGeometry()
            if primary.width() > 0 and primary.height() > 0:
                return primary

        return QtCore.QRect(0, 0, 1366, 768)

    def _clamp_window_to_work_area(self, available: QtCore.QRect | None = None) -> tuple[int, int]:
        work_area = available or self._resolved_work_area()
        fitted_w, fitted_h = _clamp_main_window_size(work_area, (self.width(), self.height()))
        maximum_w, maximum_h = _clamp_main_window_size(
            work_area,
            (_MAIN_MAX_WIDTH, _MAIN_MAX_HEIGHT),
        )
        # Bound the resizable workspace to the current monitor without locking
        # it to the restored/current size.
        self.setMaximumWidth(maximum_w)
        self.setMaximumHeight(maximum_h)
        if self.isMaximized():
            self.showNormal()

        min_size = self.minimumSize()
        if min_size.isValid() and not min_size.isNull():
            self.setMinimumSize(min(min_size.width(), fitted_w), min(min_size.height(), fitted_h))

        self.resize(min(self.width(), fitted_w), min(self.height(), fitted_h))
        return fitted_w, fitted_h

    def _fit_window_to_current_screen(self, *_args: object) -> None:
        available = self._current_screen()
        if available is None:
            available = self._resolved_work_area()
        if available.width() <= 0 or available.height() <= 0:
            return

        clamp_to_work_area = getattr(self, "_clamp_window_to_work_area", None)
        if callable(clamp_to_work_area):
            fitted_w, fitted_h = clamp_to_work_area(available)
        else:
            fitted_w, fitted_h = _clamp_main_window_size(
                available,
                (self.width(), self.height()),
            )
            if self.isMaximized():
                self.showNormal()
            minimum_size = self.minimumSize()
            if minimum_size.isValid() and not minimum_size.isNull():
                self.setMinimumSize(
                    min(minimum_size.width(), fitted_w),
                    min(minimum_size.height(), fitted_h),
                )
            self.setMaximumSize(fitted_w, fitted_h)
            self.resize(min(self.width(), fitted_w), min(self.height(), fitted_h))
        frame = self.frameGeometry()
        frame.setSize(QtCore.QSize(
            min(frame.width(), fitted_w),
            min(frame.height(), fitted_h),
        ))
        x = max(
            available.left(),
            min(frame.left(), available.right() - frame.width() + 1),
        )
        y = max(
            available.top(),
            min(frame.top(), available.bottom() - frame.height() + 1),
        )
        frame.moveTopLeft(QtCore.QPoint(x, y))
        self.setGeometry(frame)
        if hasattr(self, "_main_splitter") and hasattr(self, "_sidebar") and self.width() > 0:
            sidebar_target = clamp_sidebar_target(
                self._sidebar_default_width,
                window_width=self.width(),
                minimum=self._sidebar_min_width,
                maximum=_MAX_SIDEBAR_WIDTH,
                ratio=_SIDEBAR_RESTORE_RATIO,
            )
            self._main_splitter.setSizes(
                [sidebar_target, max(1, self.width() - sidebar_target)]
            )

    def _wire_screen_guard(self) -> None:
        if getattr(self, "_screen_guarded", False):
            return
        handle = self.windowHandle()
        if handle is None or handle.screen() is None:
            QtCore.QTimer.singleShot(75, self._wire_screen_guard)
            return
        handle.screen().availableGeometryChanged.connect(
            lambda *_sig_args: self._fit_window_to_current_screen(*_sig_args)
        )
        handle.screenChanged.connect(lambda _screen: self._fit_window_to_current_screen())
        self._screen_guarded = True

    def showEvent(self, event: QtGui.QShowEvent) -> None:  # type: ignore[override]
        super().showEvent(event)
        self._fit_window_to_current_screen()
        QtCore.QTimer.singleShot(0, self._fit_window_to_current_screen)
        self._wire_screen_guard()
        self._schedule_startup_guide()

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
            _allow_horizontal_compression(lbl)

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
        if self._queue_active:
            self.set_status("Queue run is already active.")
            return False

        if hasattr(self._measurement, "is_active") and self._measurement.is_active():
            self.set_status("Cannot start queue while a live measurement is running.")
            return False

        if self._queue_lease is not None:
            self._release_queue_lease()
        self._queue_paused = False
        self._queue_active = False
        self._queue_plan = []
        self._queue_pos = 0
        self._queue_resume_pos = 0
        self._queue_current_sample = None
        self._queue_current_command = None
        self.set_flow_state("running")

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

        if not self._queue_plan:
            self.set_status("Queue has no executable commands.")
            QtWidgets.QMessageBox.information(self, "Queue", "Queue has no measurement steps.")
            return False

        vacuum_fault = self._queue_vacuum_fault_reason()
        if vacuum_fault:
            self.set_flow_state("error")
            self.set_status(f"Queue cannot start: {vacuum_fault}")
            QtWidgets.QMessageBox.critical(
                self,
                "Vacuum Fault",
                f"Queue cannot start because vacuum is not ready:\n\n{vacuum_fault}",
            )
            return False
        squid_fault = self._queue_squid_fault_reason()
        if squid_fault:
            self.set_flow_state("error")
            self.set_status(f"Queue cannot start: {squid_fault}")
            QtWidgets.QMessageBox.critical(
                self,
                "SQUID Communication Fault",
                f"Queue cannot start because SQUID communication is not ready:\n\n{squid_fault}",
            )
            return False

        try:
            self._queue_lease = self.acquire_device("changer", "queue_workflow")
        except DeviceOwnershipError as exc:
            self.set_status(f"Queue cannot start: {exc}")
            QtWidgets.QMessageBox.critical(self, "Queue Error", str(exc))
            return False

        missing = self._missing_required_queue_methods(self._queue_plan)
        if missing:
            message = "Queue requires unsupported backend commands:\n\n" + "\n".join(
                f"• {item}" for item in missing
            )
            self.set_status("Queue cannot start: missing required hardware commands.")
            QtWidgets.QMessageBox.critical(
                self,
                "Queue Error",
                message,
            )
            self._release_queue_lease()
            return False

        self._queue_active = True
        self._queue_pos = 0
        self._queue_last_warnings = []
        self._queue_current_sample = None
        self._queue_current_command = None
        self.set_status(f"Queue run started with {len(self._queue_plan)} commands.")
        self._save_queue_state()
        self._run_next_queue_command()
        return True

    def _release_queue_lease(self) -> None:
        """Release queue workflow lease if present."""
        if self._queue_lease is None:
            return
        try:
            self._queue_lease.release()
        except Exception:
            pass
        finally:
            self._queue_lease = None

    def _finalize_queue_run(
        self,
        state: str,
        *,
        reason: str | None = None,
        return_to_safe: bool = True,
        clear_plan: bool = True,
        clear_current: bool = True,
        clear_command: bool = True,
    ) -> None:
        """Set a stable terminal queue state and persist state/ownership."""
        if return_to_safe:
            self._return_queue_to_safe_state()

        self._queue_active = False
        if clear_plan:
            self._queue_plan = []
            self._queue_pos = 0
            self._queue_resume_pos = 0
        else:
            self._queue_resume_pos = self._queue_pos

        if clear_current:
            self._queue_current_sample = None
        if clear_command:
            self._queue_current_command = None

        self._release_queue_lease()
        self._save_queue_state()
        if state:
            self.set_flow_state(state)
        if reason is not None:
            self.set_status(reason)

    def cancel_queue_run(self, reason: str = "Queue cancelled.") -> None:
        """Cancel active queue automation while leaving controls in a safe state."""
        self._queue_paused = False
        if not self._queue_active:
            self._release_queue_lease()
            return

        if hasattr(self._measurement, "halt_run"):
            self._measurement.halt_run()
        if self._queue_current_sample:
            self._sample_queue.set_queue_sample_failed(self._queue_current_sample)
        self._return_queue_to_safe_state()
        self._finalize_queue_run("idle", reason=reason, return_to_safe=False)

    def _run_next_queue_command(self) -> None:
        if not self._queue_active:
            return
        if self._queue_paused:
            self.set_flow_state("paused")
            return
        vacuum_fault = self._queue_vacuum_fault_reason()
        if vacuum_fault:
            if self._queue_current_sample:
                self._sample_queue.set_queue_sample_failed(self._queue_current_sample)
            self._finalize_queue_run(
                "error",
                reason=f"Queue halted by vacuum fault: {vacuum_fault}",
                return_to_safe=True,
            )
            return
        squid_fault = self._queue_squid_fault_reason()
        if squid_fault:
            if self._queue_current_sample:
                self._sample_queue.set_queue_sample_failed(self._queue_current_sample)
            self._finalize_queue_run(
                "error",
                reason=f"Queue halted by SQUID communication fault: {squid_fault}",
                return_to_safe=True,
            )
            return
        if self._queue_pos >= len(self._queue_plan):
            self._finalize_queue_run("complete", reason="Queue run complete.")
            return

        self._queue_current_command = self._queue_plan[self._queue_pos]
        self._queue_pos += 1
        self._queue_resume_pos = self._queue_pos

        if self._queue_current_command is None:
            return

        if self._queue_current_command.command_type == "Meas":
            self._run_measure_command()
        else:
            self._run_automation_command(self._queue_current_command)

    def _run_measure_command(self) -> None:
        if self._queue_current_command is None:
            return
        sample_name = self._queue_current_command.sample_name or ""
        if not sample_name:
            self.set_status("Queue measurement command missing sample name; continuing.")
            self._run_next_queue_command()
            return

        self._queue_current_sample = sample_name
        self._sample_queue.start_queue_sample(sample_name)
        try:
            started = self._measurement.start_measurement_for_sample(
                sample_name,
                queue_run=True,
                owner="queue_workflow",
            )
        except TypeError:
            started = self._measurement.start_measurement_for_sample(
                sample_name,
                queue_run=True,
            )

        if not started:
            self._sample_queue.set_queue_sample_failed(sample_name)
            self._finalize_queue_run(
                "error",
                reason="Queue run failed to start.",
                return_to_safe=True,
                clear_plan=True,
            )
            return
        self._save_queue_state()

    def _run_automation_command(self, command: QueueCommand) -> None:
        """Execute queue commands that do not require a full measurement sequence."""
        if not self._queue_active:
            return
        if self._queue_paused:
            self.set_flow_state("paused")
            return
        self._queue_current_sample = None
        lease: object | None = None
        command_type = command.command_type
        try:
            lease = self.acquire_device("changer", "queue_workflow")

            method_name: str
            arg: str | int | None
            if command_type == "InitUp":
                method_name = "init_up"
                arg = command.file_id
            elif command_type == "Holder":
                method_name = "holder"
                arg = command.hole
            elif command_type == "Goto":
                method_name = "goto_hole"
                arg = command.hole
            elif command_type == "Flip":
                method_name = "flip"
                arg = None
            else:
                raise ValueError(f"Unsupported queue command '{command_type}'.")
            self._run_device_command(
                self._measurement_backend,
                method_name,
                arg,
                required=self._is_queue_command_required(command_type),
            )
            self._set_queue_position_status(command_type, command)
            self._save_queue_state()
            if self._queue_paused:
                self.set_flow_state("paused")
                return
            self._run_next_queue_command()
            return
        except Exception as exc:
            if self._queue_current_sample:
                self._sample_queue.set_queue_sample_failed(self._queue_current_sample)
            self._finalize_queue_run(
                "error",
                reason=f"Queue {command.command_type} failed: {exc}",
                return_to_safe=True,
            )
        finally:
            if lease is not None:
                try:
                    lease.release()
                except Exception:
                    pass

    def _run_device_command(
        self,
        backend: object,
        method_name: str,
        arg: str | int | None = None,
        *,
        required: bool,
    ) -> None:
        method = getattr(backend, method_name, None)
        if method is None or not callable(method):
            if required:
                raise AttributeError(
                    f"Backend does not support required command '{method_name}'."
                )
            self._queue_last_warnings.append(
                f"Backend missing optional command '{method_name}'"
            )
            self.set_status(f"Queue warning: {method_name} command not supported by backend.")
            return

        if arg is None:
            method()
        else:
            method(arg)

        self.set_flow_state("positioning")

    def _is_queue_command_required(self, command_type: str) -> bool:
        """Return whether queue movement must be implemented before run proceeds."""
        # In no-comm mode we allow graceful fallback so queue compilation can still
        # exercise the flow without a queue-capable hardware backend.
        return not bool(self.config.general.nocomm)

    def _missing_required_queue_methods(
        self,
        plan: list[QueueCommand],
    ) -> list[str]:
        if self.config.general.nocomm:
            return []

        missing: list[str] = []
        for cmd in plan:
            command_type = cmd.command_type
            if command_type not in {"InitUp", "Holder", "Goto", "Flip", "Meas"}:
                continue
            if command_type == "Meas":
                continue
            if command_type == "InitUp":
                method_name = "init_up"
            elif command_type == "Holder":
                method_name = "holder"
            elif command_type == "Goto":
                method_name = "goto_hole"
            elif command_type == "Flip":
                method_name = "flip"
            else:
                continue

            method = getattr(self._measurement_backend, method_name, None)
            if not callable(method):
                if method_name not in missing:
                    missing.append(method_name)
        return missing

    def _set_queue_position_status(self, command_type: str, command: QueueCommand) -> None:
        details = f"{command_type} command"
        if command.hole not in (0, -1):
            details = f"{details} @ hole {command.hole}"
        if command_type == "InitUp" and command.file_id:
            details = f"{details} for {command.file_id}"
        if command.hole in (0, -1):
            self.set_position("Queue position: return / home")
        elif command.hole > 0:
            self.set_position(f"Hole {command.hole}")
        self.set_status(f"{details} completed.")

    def _return_queue_to_safe_state(self) -> None:
        try:
            return_to_safe_state = getattr(self._measurement_backend, "return_to_safe_state")
        except AttributeError:
            return

        if not callable(return_to_safe_state):
            return

        try:
            return_to_safe_state()
        except Exception as exc:
            self.set_status(f"Queue safe-state returned with warning: {exc}")
            self.log_event(f"Queue safe-state warning: {exc}")

    def _queue_vacuum_fault_reason(self) -> str | None:
        try:
            require_vacuum_ready(
                self._vacuum_backend,
                warn_threshold=float(self.config.vacuum.warn_threshold),
            )
        except Exception as exc:
            return str(exc)
        return None

    def _queue_squid_fault_reason(self) -> str | None:
        if self.config.general.nocomm or getattr(self._squid_backend, "simulated", False):
            return None
        if not any(cmd.command_type == "Meas" for cmd in self._queue_plan):
            return None
        try:
            require_squid_ready(self._squid_backend)
        except Exception as exc:
            return str(exc)
        return None

    def _on_queue_sample_finished(self, aborted: bool, sample: str) -> None:
        if not self._queue_active:
            return
        had_error = (
            self._measurement.take_last_run_error()
            if hasattr(self._measurement, "take_last_run_error")
            else False
        )
        if aborted:
            self._return_queue_to_safe_state()
            self._sample_queue.set_queue_sample_failed(sample)
            self._finalize_queue_run("halted", reason=f"Queue stopped after sample {sample}.")
            return
        if had_error:
            self._sample_queue.set_queue_sample_failed(sample)
            self._finalize_queue_run("error", reason=f"Queue sample {sample} failed with an error.")
            return
        self._sample_queue.mark_queue_sample_done(sample)
        if self._queue_current_command is not None:
            self._set_queue_position_status("Meas", self._queue_current_command)
        self._save_queue_state()
        self._run_next_queue_command()

    def measurement_backend(self) -> MeasurementAutomationBackend:
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
        self._sync_dashboard_run_state()

    # ── Menu bar ──────────────────────────────────────────────────────────────
    def _build_menu(self) -> None:
        mb = self.menuBar()

        fm = mb.addMenu("&File")
        fm.addAction("&New Session", self._new_session)
        fm.addAction("&Load Sample Index…", self._load_sample_from_index)
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
        vm.addAction("&Sample Queue Monitor", self._launch_step_monitor)
        vm.addAction("&Reset Layout", self._reset_layout)
        vm.addSeparator()
        vm.addAction("&Webcam Monitor", self._launch_webcam)

        flm = mb.addMenu("F&low")
        self._flow_running_action = flm.addAction("&Running / Resume", self._flow_resume_requested)
        self._flow_paused_action = flm.addAction("&Paused", self._flow_pause_requested)
        self._flow_halted_action = flm.addAction("&Halted", self._flow_halt_requested)
        for action in (
            self._flow_running_action,
            self._flow_paused_action,
            self._flow_halted_action,
        ):
            action.setCheckable(True)
        self._flow_action_group = QtGui.QActionGroup(self)
        self._flow_action_group.setExclusive(True)
        self._flow_action_group.addAction(self._flow_running_action)
        self._flow_action_group.addAction(self._flow_paused_action)
        self._flow_action_group.addAction(self._flow_halted_action)
        flm.addSeparator()
        self._status_override_action = flm.addAction(
            "Status Color &Override", self._toggle_status_override
        )
        self._status_override_action.setCheckable(True)
        self._status_override_action.setToolTip(
            "Visual maintenance indicator only; never bypasses hardware interlocks or preflight."
        )

        dm = mb.addMenu("&Diagnostics")
        dm.addAction("DC &Motors",    self._launch_dc_motors)
        dm.addAction("&SQUID Comm",   self._launch_squid)
        dm.addAction("&Vacuum",       self._launch_vacuum)
        dm.addSeparator()
        af_sub = dm.addMenu("AF &Demagnetizer")
        af_sub.addAction("Set Up AF Sequence", self._launch_af)
        self._af_demo_action = af_sub.addAction(
            "Run SIMULATED AF Example", self._launch_af_demo
        )
        self._af_queue_demo_action = af_sub.addAction(
            "Run SIMULATED AF Queue Example", self._launch_af_queue_demo
        )
        for action in (self._af_demo_action, self._af_queue_demo_action):
            action.setEnabled(bool(self.config.general.nocomm))
            action.setToolTip(
                "Available only in No-Communication mode; never runs live hardware."
            )
        af_sub.addAction("AF Tuner / ClipTest", self._launch_af_tuner)
        af_sub.addAction("AF Field Calibration", self._launch_af_tuner)
        irm_sub = dm.addMenu("&IRM / ARM")
        irm_sub.addAction("IRM / ARM Window",       self._launch_irm)
        irm_sub.addAction("IRM Field Calibration", self._launch_irm)
        irm_sub.addAction("IRM Voltage Calibration", self._launch_irm_voltage_calibration)
        dm.addAction("Thermal Routine Planning", self._launch_thermal_routine_planning)
        dm.addAction("908A &Gaussmeter", self._launch_gaussmeter)
        dm.addAction("Susceptibility &Bridge", self._launch_susceptibility_bridge)
        dm.addSeparator()
        dm.addAction("&VRM Data Collection", self._launch_vrm)
        dm.addAction("Calibrate &Rod", self._launch_calibrate_rod)

        hm = mb.addMenu("&Help")
        hm.addAction("&Quick Start", self._launch_startup_guide)
        hm.addAction("Where did this &VB6 control go?", self._launch_transition_help)
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

    def _rebuild_diagnostic_backends(self, *, nocomm: bool) -> None:
        """Build the diagnostic backends, failing closed in hardware mode.

        A device that cannot be constructed becomes an explicit
        ``UnavailableBackend`` that reports the real error. It is never a
        simulator, so the diagnostics windows stay usable while queue and
        measurement preflight still block.
        """
        self._vacuum_backend = build_backend_or_unavailable(
            "Vacuum", build_vacuum_backend, self.config.vacuum, nocomm=nocomm
        )
        self._irm_arm_backend = build_backend_or_unavailable(
            "IRM/ARM", build_irm_arm_backend, self.config.irm_arm, nocomm=nocomm
        )
        self._af_demag_backend = build_backend_or_unavailable(
            "AF demagnetizer", build_af_demag_backend, self.config.af_demag, nocomm=nocomm
        )
        self._squid_backend = build_backend_or_unavailable(
            "SQUID", build_squid_backend, self.config.squid, nocomm=nocomm
        )
        self._susceptibility_backend = build_backend_or_unavailable(
            "Susceptibility bridge",
            build_susceptibility_backend,
            self.config.susceptibility,
            nocomm=nocomm,
        )
        self._dc_motor_backend = build_backend_or_unavailable(
            "DC motors",
            build_dcmotor_backend,
            port=(self.config.changer.port or "COM3").strip(),
            baud=int(self.config.changer.baud or 9600),
            nocomm=nocomm,
        )

    # ── No-Comm toggle ────────────────────────────────────────────────────────
    def _on_nocomm_toggled(self, on: bool) -> None:
        self.config.general.nocomm = bool(on)
        self.config.save()
        self._measurement_backend = build_measurement_backend(self.config)
        self._rebuild_diagnostic_backends(nocomm=bool(on))
        self.set_flow_state(self._workflow_state)
        mode = "No-Comm simulation" if on else "hardware"
        self.set_status(f"Operating mode changed to {mode}.")
        for name in ("_af_demo_action", "_af_queue_demo_action"):
            action = getattr(self, name, None)
            if action is not None:
                action.setEnabled(bool(on))
        if hasattr(self, "_dashboard"):
            self._refresh_dashboard_diagnostics()

    # ── Diagnostic launchers ───────────────────────────────────────────────────
    def _launch_dc_motors(self) -> None:
        self._run_owned_dialog(
            "changer",
            "dc_motors_panel",
            lambda owner: DCMotorDialog(
                owner,
                backend=self._dc_motor_backend,
                port=(self.config.changer.port or "COM3").strip(),
                baud=int(self.config.changer.baud or 9600),
            ),
            modal=False,
        )

        # Ensure all configured port settings are visible if user asks to open again.
        if self.config.changer.port:
            self.set_status(f"DC Motor dialog opened for {self.config.changer.port}.")
        else:
            self.set_status("DC Motor dialog opened.")

    @staticmethod
    def _af_demo_labels() -> list[str]:
        """Canonical AF demo sequence labels."""
        return ["NRM", "AF25", "AF50", "AF100", "AF200", "AF400", "AF800", "SUSC"]

    def _launch_af(self) -> None:
        self._prepare_af_workflow(auto_start=False)

    def _launch_af_demo(self) -> None:
        """Run the AF demo sequence directly from the diagnostics launcher."""
        self._prepare_af_workflow(auto_start=True)

    def _launch_af_queue_demo(self) -> None:
        """Run the AF demo sequence as a queue sample for automated handling."""
        self._prepare_af_workflow(auto_start=True, queue_mode=True)

    def _build_af_demo_queue_samples(
        self, *, sample_name: str = "SIMULATED_AF_EXAMPLE"
    ) -> list[QueueSample]:
        """Build a single synthetic queue sample for AF demo automation."""
        return [
            QueueSample(
                sample_name=sample_name,
                file_id="SIMULATED_AF_EXAMPLE",
                hole=1,
                do_up=True,
                do_both=False,
                measurement_step_count=max(1, len(self._af_demo_labels())),
            )
        ]

    def _prepare_af_workflow(self, *, auto_start: bool, queue_mode: bool = False) -> bool:
        """Load AF sequence defaults and optionally auto-start the run."""
        if auto_start and not bool(getattr(self.config.general, "nocomm", False)):
            message = (
                "The simulated AF example is disabled in hardware mode. "
                "Use Set Up AF Sequence with an operator-selected specimen, or "
                "explicitly enable No-Communication mode for a simulation."
            )
            self.set_status(message)
            QtWidgets.QMessageBox.warning(self, "Simulated AF Example", message)
            return False

        self.load_sequence_labels(self._af_demo_labels())
        sample_name = "SIMULATED_AF_EXAMPLE" if auto_start else self._current_sample
        if auto_start:
            self.set_current_sample(sample_name)
        if sample_name and sample_name != "UNKNOWN" and hasattr(
            self._measurement, "set_specimen_context"
        ):
            self._measurement.set_specimen_context(
                sample=sample_name,
                depth="—",
                treatment=(
                    "SIMULATED AF example — not hardware evidence"
                    if auto_start
                    else "AF workflow"
                ),
            )
        if queue_mode:
            self.set_status("SIMULATED AF example loaded into queue workflow.")
            self._nav_select(1)
            if hasattr(self._measurement, "start_measurement_for_sample"):
                # Measurement panel context is prepared so the queued run uses AF labels
                # without requiring user intervention.
                self.set_status("SIMULATED AF example queue started.")
            samples = self._build_af_demo_queue_samples()
            options = QueueOptions(
                ascending=True,
                load_return=True,
                do_return=True,
                repeat_holder=True,
                samples_between_holder=8,
                use_xy_table=not bool(getattr(self.config.general, "nocomm", False)),
            )
            if self.start_queue_run(samples, options):
                return True
            self.set_status("Unable to start simulated AF example queue. Check readiness.")
            return False

        if auto_start:
            self.set_status("SIMULATED AF example loaded in Live Measurement.")
        elif sample_name and sample_name != "UNKNOWN":
            self.set_status(f"AF sequence loaded for {sample_name}; review before starting.")
        else:
            self.set_status("AF sequence loaded. Select a real sample before starting.")
        self._nav_select(3)
        if not auto_start:
            return True

        started = bool(
            hasattr(self._measurement, "start_measurement_for_sample")
            and self._measurement.start_measurement_for_sample(sample_name)
        )
        if not started:
            self.set_status(
                "Unable to auto-start simulated AF example. Manual controls remain available."
            )
            return False

        self.set_status(
            "SIMULATED AF example started — not hardware evidence. "
            "Use Pause/Resume and Halt in Live Measurement to control."
        )
        return True

    def _launch_irm(self) -> None:
        self._run_owned_dialog(
            "irm",
            "irm_panel",
            lambda owner: IrmArmDialog(owner, backend=self._irm_arm_backend),
        )

    def _launch_af_tuner(self) -> None:
        self._launch_external_tool(
            target_path="af_tuner/main.py",
            module="af_tuner",
            app_name="AF Tuner / ClipTest",
        )

    def _launch_vacuum(self) -> None:
        self._run_owned_dialog(
            "vacuum",
            "vacuum_panel",
            lambda owner: VacuumDialog(owner, backend=self._vacuum_backend),
        )

    def _launch_squid(self) -> None:
        self._run_owned_dialog(
            "squid",
            "squid_panel",
            lambda owner: SquidCommDialog(owner, backend=self._squid_backend),
        )

    def _launch_gaussmeter(self) -> None:
        self._launch_external_tool(
            target_path="gaussmeter_control/main.py",
            module="gaussmeter_control",
            app_name="908A Gaussmeter",
        )

    def _launch_vrm(self) -> None:
        context = build_vrm_launch_context(
            active_automation=self._has_active_automation(),
            nocomm=bool(getattr(self.config.general, "nocomm", False)),
        )
        try:
            context_path = write_vrm_launch_context(context=context)
        except OSError as exc:
            QtWidgets.QMessageBox.warning(
                self,
                "VRM launcher",
                f"Unable to prepare VRM run context: {exc}",
            )
            return
        self._launch_external_tool(
            target_path="vrm_logger/main.py",
            module="vrm_logger",
            app_name="VRM Logger",
            env={VRM_CONTEXT_ENV: str(context_path)},
        )

    def _launch_susceptibility_bridge(self) -> None:
        self._run_owned_dialog(
            "susceptibility",
            "susceptibility_panel",
            lambda owner: SusceptibilityDialog(
                owner, backend=self._susceptibility_backend
            ),
        )

    def _launch_calibrate_rod(self) -> None:
        self._nav_select(5)
        self._calibration_panel.set_procedure("gaussmeter_baseline")
        self._calibration_panel.set_mode("automated")
        self.set_status("Opened Calibration Center for rod-related SQUID workflow setup.")

    def _launch_irm_voltage_calibration(self) -> None:
        self._nav_select(5)
        self._calibration_panel.set_procedure("irm_voltage")
        self.set_status("Opened Calibration Center for IRM voltage calibration.")

    def _launch_thermal_routine_planning(self) -> None:
        self._nav_select(5)
        self._calibration_panel.set_procedure("thermal_routine")
        self.set_status("Opened Calibration Center for thermal routine planning.")

    def _launch_external_tool(
        self,
        *,
        target_path: str,
        module: str,
        app_name: str,
        env: dict[str, str] | None = None,
    ) -> None:
        """Launch a sibling package's main entry point as a helper process."""
        if self._has_active_automation():
            if (
                QtWidgets.QMessageBox.question(
                    self,
                    "Active Automation",
                    (
                        "A live measurement or queue run is active.\n\n"
                        "Launching this tool now may interfere with hardware ownership.\n"
                        "Continue?"
                    ),
                    QtWidgets.QMessageBox.StandardButton.Yes
                    | QtWidgets.QMessageBox.StandardButton.No,
                    QtWidgets.QMessageBox.StandardButton.No,
                )
                != QtWidgets.QMessageBox.StandardButton.Yes
            ):
                return

        base_root = Path(__file__).resolve().parents[2]
        try:
            launch = resolve_tool_launch(
                module=module,
                source_root=base_root,
                source_relative=target_path,
            )
        except ToolUnavailableError as exc:
            QtWidgets.QMessageBox.warning(
                self,
                "Tool unavailable",
                f"{app_name} is not installed correctly.\n\n{exc}",
            )
            return

        try:
            process_env = os.environ.copy()
            if env:
                process_env.update(env)
            subprocess.Popen(
                list(launch.command),
                cwd=str(launch.cwd) if launch.cwd is not None else None,
                env=process_env,
            )
            self.set_status(f"Launched {app_name} from {launch.source}.")
        except Exception as exc:  # pragma: no cover - platform/environment dependent
            QtWidgets.QMessageBox.warning(
                self,
                f"{app_name} launcher",
                f"Unable to launch {app_name}: {exc}",
            )

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

        if not callable(dialog_factory):
            if lease is not None:
                lease.release()
            QtWidgets.QMessageBox.warning(self, "Dialog", "Invalid dialog factory.")
            return

        try:
            dlg = dialog_factory(self)
            if not isinstance(dlg, QtWidgets.QWidget):
                raise TypeError("Dialog factory must return a QWidget/QDialog instance.")
            dlg.setWindowIcon(self.windowIcon())

            if modal:
                dlg.exec()
            else:
                dlg.setAttribute(
                    QtCore.Qt.WidgetAttribute.WA_DeleteOnClose,
                    True,
                )
                self._owned_dialog_leases[resource] = lease

                @QtCore.Slot(object)
                def _release_owner(_obj: object) -> None:
                    held_lease = self._owned_dialog_leases.pop(resource, None)
                    if held_lease is not None:
                        held_lease.release()

                dlg.destroyed.connect(_release_owner)
                dlg.show()
                dlg.raise_()
        except Exception as exc:
            if lease is not None:
                lease.release()
                lease = None
            QtWidgets.QMessageBox.warning(
                self,
                "Dialog Error",
                f"Unable to open diagnostic dialog: {exc}",
            )
        else:
            if lease is not None and modal:
                lease.release()

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
        del on_close  # retained for compatibility with existing callers/tests
        if prompt and hasattr(self, "_sequence"):
            if not self._sequence.confirm_discard_changes():
                return False

        active = self._has_active_automation()
        if prompt and active:
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
        elif prompt:
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

        if active:
            self.halt_measurement()
        return True

    def _request_shutdown(self) -> None:
        """Run shutdown flow from menu and toolbar actions."""
        self.close()

    def _restore_layout_state(self) -> None:
        geometry = self._settings.value("ui/window_geometry")
        if geometry:
            if self.restoreGeometry(geometry):
                self.showNormal()
                self.setWindowState(QtCore.Qt.WindowState.WindowNoState)
            # Avoid resurrecting legacy full-width restores after package
            # upgrades: enforce compact startup bounds while preserving
            # splitter/queue state persistence.
            available = self._resolved_work_area()
            self._clamp_window_to_work_area(available)
            self._fit_window_to_current_screen()
        else:
            screen = QtWidgets.QApplication.primaryScreen()
            if screen is not None:
                self.resize(*_clamp_main_window_size(screen.availableGeometry(), _DEFAULT_WINDOW_SIZE))
            else:
                self.resize(*_DEFAULT_WINDOW_SIZE)

        panel_index = self._settings.value("ui/active_panel", 0, type=int)
        self._nav_select(panel_index or 0)

        sidebar_width = self._settings.value("ui/sidebar_width", self._sidebar_default_width, type=int)
        splitter_state = self._settings.value("ui/main_splitter_state")
        sidebar_target = self._sidebar_default_width
        sidebar_cap = clamp_sidebar_target(
            self._sidebar_default_width,
            window_width=self.width(),
            minimum=self._sidebar_min_width,
            maximum=_MAX_SIDEBAR_WIDTH,
            ratio=_SIDEBAR_RESTORE_RATIO,
        )
        if isinstance(sidebar_width, int):
            sidebar_target = clamp_sidebar_target(
                int(sidebar_width),
                window_width=self.width(),
                minimum=self._sidebar_min_width,
                maximum=_MAX_SIDEBAR_WIDTH,
                ratio=_SIDEBAR_RESTORE_RATIO,
            )

        if splitter_state:
            self._main_splitter.restoreState(splitter_state)
            sidebar_target = clamp_sidebar_target(
                self._sidebar.width(),
                window_width=self.width(),
                minimum=self._sidebar_min_width,
                maximum=_MAX_SIDEBAR_WIDTH,
                ratio=_SIDEBAR_RESTORE_RATIO,
            )
        self._main_splitter.setSizes([sidebar_target, max(1, self.width() - sidebar_target)])
        self._restore_queue_state()

    def _on_splitter_moved(self, *_args: int) -> None:
        """Persist manual sidebar width changes for the next launch."""
        if not hasattr(self, "_main_splitter"):
            return
        width = int(self._sidebar.width())
        width_cap = clamp_sidebar_target(
            self._sidebar.width(),
            window_width=self.width(),
            minimum=self._sidebar_min_width,
            maximum=_MAX_SIDEBAR_WIDTH,
            ratio=_SIDEBAR_RESTORE_RATIO,
        )
        if width < self._sidebar_min_width:
            width = self._sidebar_min_width
            self._main_splitter.setSizes(
                [width, max(1, self.width() - width)]
            )
            return
        if width > width_cap:
            width = width_cap
            self._main_splitter.setSizes(
                [width, max(1, self.width() - width)]
            )
        self._settings.setValue("ui/sidebar_width", width)

    def _save_layout_state(self) -> None:
        self._clamp_window_to_work_area()
        self._settings.setValue("ui/window_geometry", self.saveGeometry())
        self._settings.setValue("ui/active_panel", self._stack.currentIndex())
        self._settings.setValue("ui/main_splitter_state", self._main_splitter.saveState())
        self._settings.setValue(
            "ui/sidebar_width",
            clamp_sidebar_target(
                self._sidebar.width(),
                window_width=self.width(),
                minimum=self._sidebar_min_width,
                maximum=_MAX_SIDEBAR_WIDTH,
                ratio=_SIDEBAR_RESTORE_RATIO,
            ),
        )
        self._save_queue_state()

    def _sync_settings(self) -> None:
        """Flush real QSettings while supporting lightweight test bridges."""
        sync = getattr(self._settings, "sync", None)
        if callable(sync):
            sync()

    def _save_queue_state(self) -> None:
        try:
            queue_rows = self._sample_queue.row_snapshot()
            self._settings.setValue(_QSETTINGS_QUEUE_ROWS, json.dumps(queue_rows))
            self._settings.setValue(_QSETTINGS_QUEUE_PROGRESS, int(self._queue_pos))
            self._settings.setValue(_QSETTINGS_QUEUE_RESUME_POS, int(self._queue_resume_pos))
            self._settings.setValue(
                _QSETTINGS_QUEUE_CURRENT_SAMPLE, self._queue_current_sample or ""
            )
            self._settings.setValue(_QSETTINGS_QUEUE_ACTIVE, bool(self._queue_active))
            self._sync_settings()
        except Exception:
            self._settings.remove(_QSETTINGS_QUEUE_ROWS)
            self._settings.remove(_QSETTINGS_QUEUE_PROGRESS)
            self._settings.remove(_QSETTINGS_QUEUE_RESUME_POS)
            self._settings.remove(_QSETTINGS_QUEUE_CURRENT_SAMPLE)
            self._settings.remove(_QSETTINGS_QUEUE_ACTIVE)

    def _restore_queue_state(self) -> None:
        # Pull settings written by another window/process before rebuilding the
        # queue.  QSettings otherwise may retain a stale per-instance cache.
        self._sync_settings()
        raw = self._settings.value(_QSETTINGS_QUEUE_ROWS)
        if not raw:
            self._queue_current_sample = None
            return
        try:
            rows = json.loads(raw) if isinstance(raw, str) else raw
            if isinstance(rows, list):
                self._sample_queue.load_rows(rows)
                self._queue_resume_pos = int(
                    self._settings.value(_QSETTINGS_QUEUE_RESUME_POS, 0, type=int)
                )
                self._queue_pos = int(
                    self._settings.value(_QSETTINGS_QUEUE_PROGRESS, 0, type=int)
                )
                was_active = bool(
                    self._settings.value(_QSETTINGS_QUEUE_ACTIVE, False, type=bool)
                )
                recovered = self._sample_queue.recover_interrupted_samples()
                if recovered:
                    self.set_status(
                        f"Recovered {recovered} interrupted queue row(s) from last session; choose Resume, Re-run, Skip, or Abort."
                    )
                elif was_active:
                    self.set_status(
                        "Previous queue run was active; review pending rows and press Run Queue to continue."
                    )
                if self._queue_resume_pos > 0 and not was_active:
                    self.set_status(f"Recovered queue progress index {self._queue_resume_pos}.")
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
        screen = QtWidgets.QApplication.primaryScreen()
        if screen is None:
            self.resize(*_DEFAULT_WINDOW_SIZE)
        else:
            self.resize(*_clamp_main_window_size(screen.availableGeometry(), _DEFAULT_WINDOW_SIZE))
        self._main_splitter.setSizes(
            [self._sidebar_default_width, max(1, self.width() - self._sidebar_default_width)]
        )
        self._nav_select(0)
        QtWidgets.QMessageBox.information(self, "Reset Layout", "Layout reset to defaults.")

    def _new_session(self) -> None:
        """Reset transient operator work without touching configuration or holder state."""
        if self._has_active_automation():
            QtWidgets.QMessageBox.warning(
                self,
                "New Session",
                "Halt or finish the active workflow before starting a new session.",
            )
            return
        if not self._sequence.confirm_discard_changes():
            return
        if (
            QtWidgets.QMessageBox.question(
                self,
                "New Session",
                "Start a new session? This clears the current queue, sequence, and "
                "measurement view. Saved files, settings, calibration, and holder state remain unchanged.",
                QtWidgets.QMessageBox.StandardButton.Yes
                | QtWidgets.QMessageBox.StandardButton.No,
                QtWidgets.QMessageBox.StandardButton.No,
            )
            != QtWidgets.QMessageBox.StandardButton.Yes
        ):
            return
        self.cancel_queue_run("New session started.")
        self._sample_queue._clear_table()
        self._sequence.clear_for_new_session()
        if hasattr(self._measurement, "_clear_measurement_plot"):
            self._measurement._clear_measurement_plot()
        if hasattr(self._measurement, "set_specimen_context"):
            self._measurement.set_specimen_context("UNKNOWN")
        self._queue_plan = []
        self._queue_pos = 0
        self._queue_resume_pos = 0
        self._queue_current_sample = None
        self._queue_current_command = None
        self._current_step = "—"
        self._current_treatment = "—"
        self.set_step("—")
        self.set_flow_state("idle")
        self._save_queue_state()
        self.set_status("New session ready. Saved data and holder correction were preserved.")
        self.log_event("New operator session started; transient queue and sequence state cleared.")
        self._nav_select(0)

    def _launch_login(self) -> None:
        if self._has_active_automation():
            QtWidgets.QMessageBox.warning(
                self,
                "Change Operator",
                "Halt or finish the active run before changing operator identity.",
            )
            return
        dialog = LoginDialog(self)
        if dialog.exec() != QtWidgets.QDialog.DialogCode.Accepted:
            return
        self.config.general.operator = dialog.operator_name
        self.config.general.nocomm = bool(dialog.nocomm)
        self.config.save()
        self._settings_panel.load_from_config(self.config)
        self._nocomm_btn.blockSignals(True)
        self._nocomm_btn.setChecked(bool(dialog.nocomm))
        self._nocomm_btn.blockSignals(False)
        self._on_nocomm_toggled(bool(dialog.nocomm))
        self.set_status(f"Operator changed to {dialog.operator_name}.")
        self.log_event(f"Operator session changed to {dialog.operator_name}.")

    def _launch_about(self) -> None:
        AboutDialog(self).exec()

    def _show_startup_guide_enabled(self) -> bool:
        return bool(
            self._settings.value(
                _QSETTINGS_SHOW_STARTUP_GUIDE,
                True,
                type=bool,
            )
        )

    def _set_show_startup_guide(self, enabled: bool) -> None:
        self._settings.setValue(_QSETTINGS_SHOW_STARTUP_GUIDE, bool(enabled))
        self._sync_settings()

    def _schedule_startup_guide(self) -> None:
        if self._startup_guide_scheduled:
            return
        self._startup_guide_scheduled = True
        if self._show_startup_guide_enabled():
            QtCore.QTimer.singleShot(350, self._launch_startup_guide_if_enabled)

    def _launch_startup_guide_if_enabled(self) -> None:
        if self._show_startup_guide_enabled():
            self._launch_startup_guide()

    def _launch_startup_guide(self) -> None:
        dialog = self._startup_guide_dialog
        if dialog is None:
            dialog = StartupGuideDialog(
                self,
                show_at_startup=self._show_startup_guide_enabled(),
            )
            dialog.show_at_startup_changed.connect(self._set_show_startup_guide)
            dialog.open_settings_requested.connect(lambda: self._open_startup_guide_panel(4))
            dialog.open_queue_requested.connect(lambda: self._open_startup_guide_panel(1))
            self._startup_guide_dialog = dialog
        else:
            dialog.set_show_at_startup(self._show_startup_guide_enabled())
        dialog.show()
        dialog.raise_()
        dialog.activateWindow()

    def _open_startup_guide_panel(self, index: int) -> None:
        self._nav_select(index)
        if self._startup_guide_dialog is not None:
            self._startup_guide_dialog.hide()

    def _launch_debug_console(self) -> None:
        if not hasattr(self, "_debug_dlg") or not self._debug_dlg.isVisible():
            self._debug_dlg = DebugConsoleDialog(
                self,
                snapshot_provider=self._diagnostic_status_lines,
            )
            self._debug_dlg.setModal(False)
        self._debug_dlg.show()
        self._debug_dlg.raise_()

    def _diagnostic_status_lines(self):
        return collect_diagnostic_status(
            {
                "Vacuum": self._vacuum_backend,
                "AF Demag": self._af_demag_backend,
                "IRM/ARM": self._irm_arm_backend,
                "SQUID": self._squid_backend,
                "Susceptibility": self._susceptibility_backend,
                "DC Motors": self._dc_motor_backend,
            }
        )

    @QtCore.Slot()
    def _refresh_dashboard_diagnostics(self) -> None:
        """Refresh the dashboard from the same backend snapshot used by Debug."""
        if not hasattr(self, "_dashboard"):
            return
        lines = self._diagnostic_status_lines()
        self._dashboard.update_diagnostics(lines)

    def _launch_step_monitor(self) -> None:
        if not hasattr(self, "_step_dlg") or not self._step_dlg.isVisible():
            self._step_dlg = StepMonitorDialog(self)
            self._step_dlg.setModal(False)
        self._step_dlg.show()
        self._step_dlg.raise_()

    def _launch_data_viewer(self) -> None:
        try:
            launch = resolve_tool_launch(
                module="data_viewer",
                source_root=Path(__file__).resolve().parents[2],
                source_relative="data_viewer/main.py",
            )
        except ToolUnavailableError as exc:
            QtWidgets.QMessageBox.warning(
                self,
                "Data Review",
                f"Data Review is not installed correctly.\n\n{exc}",
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
                list(launch.command),
                cwd=str(launch.cwd) if launch.cwd is not None else None,
            )
        except OSError as exc:
            QtWidgets.QMessageBox.warning(
                self,
                "Data Review",
                f"Unable to launch data review utility:\n{exc}",
            )
            self._data_viewer_proc = None
            return

        self.set_status(f"Data Review launched from {launch.source} in a separate window.")

    def _launch_webcam(self) -> None:
        if not hasattr(self, "_webcam_dlg"):
            self._webcam_dlg = WebcamDialog(self)
        self._webcam_dlg.show()
        self._webcam_dlg.raise_()

    def _launch_transition_help(self) -> None:
        if not hasattr(self, "_transition_help_dlg"):
            self._transition_help_dlg = TransitionHelpDialog(self)
            self._transition_help_dlg.destination_requested.connect(
                self._open_transition_destination
            )
        self._transition_help_dlg.show()
        self._transition_help_dlg.raise_()
        self._transition_help_dlg.activateWindow()

    def _open_transition_destination(self, destination: str) -> None:
        panel_destinations = {
            "dashboard",
            "queue",
            "sequence",
            "measure",
            "settings",
            "calibration",
        }
        if destination in panel_destinations:
            self.navigate_to(destination)
            return
        launchers = {
            "dc_motors": self._launch_dc_motors,
            "vacuum": self._launch_vacuum,
            "squid": self._launch_squid,
            "irm": self._launch_irm,
            "af": self._launch_af,
            "debug": self._launch_debug_console,
            "step_monitor": self._launch_step_monitor,
            "vrm": self._launch_vrm,
            "webcam": self._launch_webcam,
            "login": self._launch_login,
            "quick_start": self._launch_startup_guide,
        }
        launcher = launchers.get(destination)
        if launcher is not None:
            launcher()

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

    def _load_sample_from_index(self) -> bool:
        """Load a real sample index and add one operator-selected row to the queue."""

        dialog = SampleSelectDialog(self)
        if dialog.exec() != QtWidgets.QDialog.DialogCode.Accepted:
            return False
        record = dialog.selected_record
        if record is None or not record["sample_name"].strip():
            QtWidgets.QMessageBox.warning(
                self,
                "Load Sample Index",
                "Select a specimen before adding it to the queue.",
            )
            return False

        default_position = f"A{self._sample_queue._table.rowCount() + 1}"
        position, accepted = QtWidgets.QInputDialog.getText(
            self,
            "Changer Position",
            f"Changer position for {record['sample_name']}:",
            QtWidgets.QLineEdit.EchoMode.Normal,
            default_position,
        )
        if not accepted:
            return False
        position = position.strip()
        if not position:
            QtWidgets.QMessageBox.warning(
                self,
                "Changer Position",
                "A changer position is required; the sample was not added.",
            )
            return False

        treatment = " → ".join(self._sequence_labels) or "NRM"
        sample_set = record["formation"].strip() or record["location"].strip()
        self._sample_queue.add_sample(
            position=position,
            name=record["sample_name"],
            sample_set=sample_set,
            treatment=treatment,
        )
        if dialog.registrations is not None:
            self.sample_registrations = dialog.registrations
        self._measurement.set_specimen_context(
            record["sample_name"],
            depth=record["depth_cm"] or "—",
            treatment=treatment,
        )
        self._nav_select(1)
        self._save_queue_state()
        source = dialog.source_path.name if dialog.source_path is not None else "sample index"
        self.set_status(
            f"Added {record['sample_name']} at {position} from {source}; review the queue before running."
        )
        return True

    def set_status(self, text: str) -> None:
        self._sb_status.setText(text)

    def set_sample(self, name: str) -> None:
        self._current_sample = name or "UNKNOWN"
        self._sb_sample.setText(f"Sample: {name}")
        self._sample_hdr.setText(f"Sample: {name}")
        self._sync_dashboard_run_state()

    def set_step(self, step: str) -> None:
        self._current_step = step or "—"
        self._current_treatment = step or "—"
        self._step_hdr.setText(f"Step: {step}")
        self._sync_dashboard_run_state()

    def set_position(self, pos: str) -> None:
        self._sb_pos.setText(f"Pos: {pos}")

    def set_flow_state(self, state: str) -> None:
        """Update the top-of-screen workflow label from a phase name."""
        self._workflow_state = state
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
            "complete": "flowComplete",
            "error": "flowError",
        }
        if self._status_override_active:
            self._flow_lbl.setText("◇  Status Override")
            self._flow_lbl.setObjectName("flowOverride")
        else:
            self._flow_lbl.setText(icons.get(state, state))
            self._flow_lbl.setObjectName(names.get(state, "flowRunning"))
        self._flow_lbl.style().unpolish(self._flow_lbl)
        self._flow_lbl.style().polish(self._flow_lbl)
        for action_name, checked_state in (
            ("_flow_running_action", "running"),
            ("_flow_paused_action", "paused"),
            ("_flow_halted_action", "halted"),
        ):
            action = getattr(self, action_name, None)
            if action is not None:
                action.setChecked(state == checked_state)
        self._sync_dashboard_run_state()

    def _elapsed_run_text(self) -> str:
        if self._run_start_time is None:
            return "00:00:00"
        seconds = max(0, int((datetime.now() - self._run_start_time).total_seconds()))
        hours, remainder = divmod(seconds, 3600)
        minutes, seconds = divmod(remainder, 60)
        return f"{hours:02d}:{minutes:02d}:{seconds:02d}"

    def _sync_dashboard_run_state(self) -> None:
        if not hasattr(self, "_dashboard"):
            return
        self._dashboard.update_run_state(
            self._workflow_state.replace("_", " ").title(),
            self._current_sample if self._current_sample != "UNKNOWN" else "—",
            self._current_step,
            self._current_treatment,
            self._elapsed_run_text(),
        )

    def _measurement_is_paused(self) -> bool:
        worker = getattr(self._measurement, "_worker", None)
        return bool(worker is not None and getattr(worker, "is_paused", False))

    def _flow_resume_requested(self, _checked: bool = False) -> None:
        if self._queue_paused or self._measurement_is_paused():
            self.toggle_queue_pause()
            return
        if self._has_active_automation():
            self.set_status("Workflow is already running.")
        else:
            self.set_status("No paused workflow to resume. Start from Queue or Live Measurement.")

    def _flow_pause_requested(self, _checked: bool = False) -> None:
        if not self._has_active_automation():
            self.set_status("No active workflow to pause.")
            self.set_flow_state(self._workflow_state)
            return
        if self._queue_paused or self._measurement_is_paused():
            self.set_status("Workflow is already paused.")
            self.set_flow_state("paused")
            return
        self.toggle_queue_pause()

    def _flow_halt_requested(self, _checked: bool = False) -> None:
        self._on_header_halt()
        self.set_flow_state(self._workflow_state)

    def _toggle_status_override(self, enabled: bool) -> None:
        """Mirror VB6 Code Grey as a visual-only maintenance indicator."""
        self._status_override_active = bool(enabled)
        self.set_flow_state(self._workflow_state)
        if enabled:
            self.set_status(
                "Status color override enabled. Hardware interlocks and preflight remain enforced."
            )
            self.log_event("Visual status override enabled; safety behavior unchanged.")
        else:
            self.set_status("Status color override cleared.")
            self.log_event("Visual status override cleared.")

    def toggle_queue_pause(self) -> None:
        """Pause or resume active queue automation."""
        if not self._queue_active:
            if hasattr(self._measurement, "is_active") and self._measurement.is_active():
                if hasattr(self._measurement, "toggle_pause"):
                    self._measurement.toggle_pause()
                elif hasattr(self._measurement, "pause_run"):
                    self._measurement.pause_run()
            return

        self._queue_paused = not self._queue_paused
        if self._queue_paused:
            if hasattr(self._measurement, "toggle_pause"):
                self._measurement.toggle_pause()
            self.set_flow_state("paused")
            self.set_status("Queue paused.")
            return

        self._queue_paused = False
        if hasattr(self._measurement, "toggle_pause"):
            self._measurement.toggle_pause()
        self.set_status("Queue resumed.")
        self._run_next_queue_command()

    def halt_measurement(self) -> None:
        """Stop active measurement and queue automation in a single action."""
        self._queue_paused = False
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


def main() -> int:
    app = QtWidgets.QApplication(sys.argv)
    apply_window_bounds_guard(app)
    apply_liquid_glass_theme(app)
    app.setStyleSheet(app.styleSheet() + _EXTRA_CSS)
    apply_main_glass_theme(app)
    assets_dir = main_assets_dir()
    icon_name, _icon_path = select_main_icon(assets_dir)
    if not icon_name:
        icon_name = "rapid_main_icon.png"
    set_app_icon(app, icon_name, assets_dir)
    window = MainWindow()
    set_app_icon(window, icon_name, assets_dir)
    screen = app.primaryScreen()
    if screen is not None:
        window.setWindowState(QtCore.Qt.WindowState.WindowNoState)
        compact_w, compact_h = _clamp_main_window_size(
            screen.availableGeometry(),
            (window.width(), window.height()),
        )
        maximum_w, maximum_h = _clamp_main_window_size(
            screen.availableGeometry(),
            (_MAIN_MAX_WIDTH, _MAIN_MAX_HEIGHT),
        )
        window.setMaximumSize(maximum_w, maximum_h)
        window.resize(min(window.width(), compact_w), min(window.height(), compact_h))
        window._fit_window_to_current_screen()
        QtCore.QTimer.singleShot(75, window._fit_window_to_current_screen)
    window.show()

    def _enforce_startup_fit() -> None:
        # If restore/load paths briefly re-expand on the way in, enforce the
        # compact guard immediately after first paint and when monitor metrics
        # have not yet fully settled.
        active = window.screen() or app.primaryScreen()
        if active is None:
            return
        compact_w, compact_h = _clamp_main_window_size(
            active.availableGeometry(),
            (window.width(), window.height()),
        )
        maximum_w, maximum_h = _clamp_main_window_size(
            active.availableGeometry(),
            (_MAIN_MAX_WIDTH, _MAIN_MAX_HEIGHT),
        )
        window.setWindowState(QtCore.Qt.WindowState.WindowNoState)
        window.setMaximumSize(maximum_w, maximum_h)
        if window.isMaximized() or window.isFullScreen():
            window.showNormal()
        window.resize(min(window.width(), compact_w), min(window.height(), compact_h))
        window._fit_window_to_current_screen()

    QtCore.QTimer.singleShot(0, _enforce_startup_fit)
    QtCore.QTimer.singleShot(150, _enforce_startup_fit)
    QtCore.QTimer.singleShot(350, _enforce_startup_fit)
    QtCore.QTimer.singleShot(750, _enforce_startup_fit)
    QtCore.QTimer.singleShot(1400, _enforce_startup_fit)
    return app.exec()
