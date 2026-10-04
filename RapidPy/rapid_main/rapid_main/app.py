from __future__ import annotations

import sys
import json
import hashlib
from copy import deepcopy
import os
import subprocess
import uuid
from dataclasses import asdict
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

from . import software_version
from .config import AppConfig
from .data_model import SampleIndexRegistration, SampleIndexRegistrations
from .io.sample_index import read_sample_index_registrations
from .specimen_metadata import capture_specimen_metadata, restore_specimen_metadata, validate_specimen_provenance
from .specimen_paths import specimen_run_directory
from .io.measurement_bundle import validate_measurement_output
from .hardware_contracts import (
    MeasurementAutomationBackend,
    QueueHardwareBackend,
    build_measurement_backend,
    config_fingerprint,
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
from .queue_compiler import QueueCommand, QueueOptions, QueueSample, compile_queue, resolve_queue_samples
from .queue_command_worker import QueueCommandWorker
from .package_launch import ToolUnavailableError, resolve_tool_launch
from .runtime_estimator import RuntimeEstimator
from .vrm import (
    VRM_CONTEXT_ENV,
    build_vrm_launch_context,
    vrm_launch_block_reason,
    write_vrm_launch_context,
)
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
        self._ownership = DeviceOwnershipManager()
        self._owned_dialog_leases: dict[str, object] = {}
        self._owned_dialogs: dict[str, QtWidgets.QWidget] = {}
        self._shutdown_cleanup_requested = False
        self._shutdown_retry_pending = False
        self._shutdown_requested_dialogs = set()
        self._shutdown_timer = QtCore.QTimer(self)
        self._shutdown_timer.setSingleShot(True)
        self._shutdown_timer.timeout.connect(self._retry_shutdown)
        self._external_process_leases: dict[int, tuple[object, list[object], QtCore.QTimer]] = {}
        self._rebuild_diagnostic_backends(nocomm=bool(self.config.general.nocomm))
        self._measurement_backend: MeasurementAutomationBackend = build_measurement_backend(
            self.config,
            susceptibility_backend=self._susceptibility_backend,
        )

        # Runtime estimator — initialised from config step times
        self._estimator = RuntimeEstimator(self.config.sequence.as_estimator_dict())
        self._sequence_labels: list[str] = []  # current loaded sequence step labels
        self._rockmag_routine_plan: object | None = None
        self._thermal_routine_plan: object | None = None
        self._queue_plan: list[QueueCommand] = []
        self._queue_pos: int = 0
        self._queue_resume_pos: int = 0
        self._queue_current_sample: str | None = None
        self._queue_current_command: QueueCommand | None = None
        self._queue_active: bool = False
        self._queue_last_warnings: list[str] = []
        self._queue_paused: bool = False
        self._queue_lease: object | None = None
        self._queue_native_startup = None
        self._queue_native_session = None
        self._queue_loaded_command = None
        self._queue_meas_pending_start = False
        self._queue_command_thread = self._queue_command_worker = None
        self._queue_command_leases = []
        self._queue_worker_command = None
        self._queue_pending_finalize = None
        self._queue_settle_timer = QtCore.QTimer(self)
        self._queue_settle_timer.setSingleShot(True)
        self._queue_settle_timer.timeout.connect(self._queue_command_settled)
        self._queue_advance_timer = QtCore.QTimer(self)
        self._queue_advance_timer.setSingleShot(True)
        self._queue_advance_timer.timeout.connect(self._run_next_queue_command)
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
        self._glass_elevation_timer = QtCore.QTimer(self)
        self._glass_elevation_timer.setSingleShot(True)
        self._glass_elevation_timer.timeout.connect(lambda: install_glass_elevation(self))
        self._glass_elevation_timer.start(0)
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
        self._defer_window_callback(0, self._refresh_dashboard_diagnostics)

        self._clock = QtCore.QTimer(self)
        self._clock.timeout.connect(self._tick_clock)
        self._clock.start(1000)
        self._fit_window_to_current_screen()
        self._defer_window_callback(0, self._fit_window_to_current_screen)
        self._defer_window_callback(0, self._wire_screen_guard)


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
        self._calibration_panel.thermal_plan_recorded.connect(self._install_thermal_plan)
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
            self._defer_window_callback(75, self._wire_screen_guard)
            return
        handle.screen().availableGeometryChanged.connect(
            lambda *_sig_args: self._fit_window_to_current_screen(*_sig_args)
        )
        handle.screenChanged.connect(lambda _screen: self._fit_window_to_current_screen())
        self._screen_guarded = True

    def showEvent(self, event: QtGui.QShowEvent) -> None:  # type: ignore[override]
        super().showEvent(event)
        self._fit_window_to_current_screen()
        self._defer_window_callback(0, self._fit_window_to_current_screen)
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
        labels = getattr(self, '_run_sequence_labels', None) if running else self._sequence_labels
        labels = labels if labels is not None else self._sequence_labels
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
        if self._rockmag_routine_plan is not None:
            planned = getattr(self._rockmag_routine_plan, "to_queue_labels", lambda: [])()
            if list(planned) != self._sequence_labels:
                self._rockmag_routine_plan = None
        if self._thermal_routine_plan is not None:
            planned = getattr(self._thermal_routine_plan, "to_queue_labels", lambda: [])()
            if list(planned) != self._sequence_labels:
                self._thermal_routine_plan = None
        self._update_runtime_display(running=False)

    def set_rockmag_routine_plan(self, plan: object | None) -> None:
        """Associate the current labels with their compiled rockmag identity."""

        if plan is None:
            self._rockmag_routine_plan = None
            return
        labels = getattr(plan, "to_queue_labels", lambda: [])()
        if list(labels) != self._sequence_labels:
            raise ValueError("rockmag routine plan does not match the active sequence labels")
        self._rockmag_routine_plan = plan
        if plan is not None:
            self._thermal_routine_plan = None

    def set_thermal_routine_plan(self, plan: object | None) -> None:
        """Associate current labels with a reviewed manual-external thermal plan."""

        if plan is None:
            self._thermal_routine_plan = None
            return
        labels = getattr(plan, "to_queue_labels", lambda: [])()
        if list(labels) != self._sequence_labels:
            raise ValueError("thermal routine plan does not match the active sequence labels")
        self._thermal_routine_plan = plan
        self._rockmag_routine_plan = None

    @QtCore.Slot(object)
    def _install_thermal_plan(self, plan: object) -> None:
        """Load a recorded thermal plan without implying automated furnace control."""

        self._sequence.install_thermal_plan(plan)
        self._nav_select(2)
        self.set_status(
            "Loaded manual-external thermal labels. Live furnace automation remains blocked."
        )

    def start_queue_run(self, samples: list[QueueSample], options: QueueOptions) -> bool:
        """Start a queue-driven measurement run.

        Validation is intentionally strict for safety and queue automation.
        """
        if self._queue_active or self._queue_command_thread is not None:
            self.set_status("Queue run is already active.")
            return False

        if hasattr(self._measurement, "is_active") and self._measurement.is_active():
            self.set_status("Cannot start queue while a live measurement is running.")
            return False

        if self._queue_lease is not None and not self._release_queue_lease():
            self.set_status('Recover the original unfinished native queue before starting another run.')
            return False
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

        if not self._sequence_labels and any(not sample.measurement_labels for sample in samples):
            self.set_status("Load a measurement sequence first.")
            return False

        try:
            samples = resolve_queue_samples(samples, self._sequence_labels)
            self._queue_plan = compile_queue(samples, options, strict=True)
            self._queue_source_indexes = {}
            for sample in samples:
                index = Path(sample.source_file).resolve() if sample.source_file else None
                output = self.config.general.data_dir or (Path.home() / 'RAPID_data')
                run_directory = specimen_run_directory(output, sample.sample_name, index)
                validate_measurement_output(run_directory, sample.sample_name,
                    simulated=bool(getattr(self._measurement_backend, 'simulated', False)))
                if not sample.source_file:
                    continue
                path = Path(sample.source_file)
                if not path.is_absolute() or path.suffix.lower() not in {'.sam', '.csv'}:
                    raise ValueError('Source indexes must be absolute SAM or CSV file paths.')
                identity = os.path.normcase(str(path.resolve()))
                if sample.file_id != identity:
                    raise ValueError('Queue source-file identity must match its original index path.')
                if identity not in self._queue_source_indexes:
                    payload = path.read_bytes()
                    registrations = read_sample_index_registrations(path)
                    if payload != path.read_bytes():
                        raise ValueError('Sample index changed while preparing the queue.')
                    self._queue_source_indexes[identity] = dict(source_file=str(path.resolve()),
                        sha256=hashlib.sha256(payload).hexdigest(), entries=[asdict(entry) for entry in registrations.entries])
                matches = [entry for entry in self._queue_source_indexes[identity]['entries'] if entry['specimen_name'] == sample.sample_name]
                if len(matches) != 1:
                    raise ValueError('Each queued specimen must match exactly one entry in its original source index.')
                source = self._queue_source_indexes[identity]
                snapshots = source.setdefault('specimens', {})
                if sample.sample_name not in snapshots:
                    snapshots[sample.sample_name] = capture_specimen_metadata(sample.sample_name,
                        sample_dir=path.parent, registrations=SampleIndexRegistrations([
                            SampleIndexRegistration(**entry) for entry in source['entries']]))
            for plan in (self._rockmag_routine_plan, self._thermal_routine_plan):
                if plan is not None and any(command.command_type == 'Meas'
                        and list(command.measurement_labels) != list(plan.to_queue_labels()) for command in self._queue_plan):
                    raise ValueError('The reviewed routine identity must match every file measurement sequence.')
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

        if self._uses_native_queue():
            return self._start_native_queue(samples, options)

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

    def _uses_native_queue(self):
        return self.config.general.nocomm is False and isinstance(self._measurement_backend, QueueHardwareBackend)

    def _start_native_queue(self, samples, options):
        missing = self._missing_required_queue_methods(self._queue_plan)
        if missing:
            self.set_flow_state('error')
            QtWidgets.QMessageBox.critical(self, 'Native Queue Unavailable',
                'Native queue commands are unavailable: ' + ', '.join(missing))
            return False
        try:
            if options.use_xy_table is not self.config.motor_station.use_xy_table:
                raise ValueError('Queue options must match the accepted physical table mode.')
            from .queue_station import QueueStationGeometry
            station = QueueStationGeometry.from_config(self.config, use_xy_table=True)
            for sample in samples:
                station.specimen_slot(sample.hole)
                station.xy_target(sample.hole)
            self._queue_lease = self.acquire_devices(
                ('measurement', 'changer', 'af_demag', 'vacuum', 'squid', 'susceptibility'), 'queue_workflow',
                allow_reentrant=False)
            if QtWidgets.QMessageBox.question(self, 'Confirm empty control rod',
                    'Remove every specimen from the control rod and verify that the table is mechanically clear.\n\n'
                    'Confirm the rod is empty before native motor and vacuum startup?',
                    QtWidgets.QMessageBox.StandardButton.Yes | QtWidgets.QMessageBox.StandardButton.No,
                    QtWidgets.QMessageBox.StandardButton.No) != QtWidgets.QMessageBox.StandardButton.Yes:
                self._release_queue_lease()
                self.set_flow_state('idle')
                return False
            directions = {}
            for sample in samples:
                if sample.file_id in directions and directions[sample.file_id] is not sample.do_up:
                    raise ValueError('All rows from one specimen file must agree on initial tray orientation.')
                if type(sample.do_up) is not bool:
                    raise ValueError('Queue specimen orientation must be an explicit boolean.')
                directions[sample.file_id] = sample.do_up
            startup = self._measurement_backend.prepare_queue_lifetime(self._vacuum_backend,
                dict(commands=[asdict(command) for command in self._queue_plan],
                     samples=[asdict(sample) for sample in samples], options=asdict(options), file_directions=directions,
                     source_indexes=self._queue_source_indexes),
                run_id='queue-' + uuid.uuid4().hex, operator=self.config.general.operator,
                empty_rod_confirmed=True)
            self._queue_native_startup, self._queue_native_session = startup, startup.session
            self._queue_active = True
            self._queue_loaded_command = None
            self._queue_meas_pending_start = False
            self._save_queue_state()
            self._start_queue_command_worker(QueueCommand('NativeStartup'),
                self._measurement_backend.start_queue_lifetime, [], recover_on_error=False)
            return True
        except Exception as exc:
            self._queue_active = False
            self._release_queue_lease()
            self.set_flow_state('error')
            self.set_status('Native queue startup failed: ' + str(exc))
            QtWidgets.QMessageBox.critical(self, 'Native Queue Error', str(exc))
            return False

    def _release_queue_lease(self) -> bool:
        """Release queue workflow lease if present."""
        session = self._queue_native_session
        if session is not None:
            if self._queue_command_thread is not None or self._measurement.is_active() or session._owner is not None:
                return False
            try:
                state = session.store.read()
                session.store.verify_history(state)
            except Exception as exc:
                self.set_status('Original queue ownership cannot be released: ' + str(exc))
                return False
            if (not state or state['family'] != 'queue' or state['token'] != session.token
                    or state['status'] != 'verified' or self._measurement_backend._client._connections
                    or self._vacuum_backend._queue_binding is not None
                    or self._queue_native_startup is None
                    or not self._queue_native_startup.instruments.is_settled()):
                return False
            try:
                session.release()
            except Exception as exc:
                self.set_status('Original queue ownership cannot be released: ' + str(exc))
                return False
            self._queue_native_session = self._queue_native_startup = None
            self._queue_loaded_command = None
            self._queue_meas_pending_start = False
        if self._queue_lease is None:
            return True
        try:
            self._queue_lease.release()
        except Exception:
            pass
        finally:
            self._queue_lease = None
        return True

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
        if (getattr(self._measurement_backend, 'queue_commands_require_worker', False) is True
                and self._measurement.is_active()):
            self._queue_pending_finalize = dict(state=state, reason=reason, return_to_safe=return_to_safe,
                clear_plan=clear_plan, clear_current=clear_current, clear_command=clear_command)
            self._measurement.halt_run()
            self.set_status('Queue is waiting for acquisition cleanup before terminal recovery.')
            return
        if self._queue_command_thread is not None:
            self._queue_pending_finalize = dict(state=state, reason=reason, return_to_safe=False,
                clear_plan=clear_plan, clear_current=clear_current, clear_command=clear_command)
            self._queue_command_worker.stop()
            self.set_status('Queue is waiting for its active hardware command and terminal recovery.')
            return
        if return_to_safe:
            backend = self._measurement_backend
            cleanup = (getattr(backend, 'finish_queue_lifetime', None) if self._queue_native_session is not None
                       else getattr(backend, 'return_to_safe_state', None))
            if getattr(backend, 'queue_commands_require_worker', False) is True and callable(cleanup):
                self._queue_pending_finalize = dict(state=state, reason=reason, return_to_safe=False,
                    clear_plan=clear_plan, clear_current=clear_current, clear_command=clear_command)
                self._start_queue_command_worker(None, cleanup, [], recover_on_error=False)
                return
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

        released = self._release_queue_lease()
        if not released:
            state = 'error'
            reason = (reason or 'Queue stopped') + '; original native queue ownership remains held for recovery.'
        self._save_queue_state()
        if state:
            self.set_flow_state(state)
        if reason is not None:
            self.set_status(reason)

    def cancel_queue_run(self, reason: str = "Queue cancelled.") -> None:
        """Cancel active queue automation while leaving controls in a safe state."""
        self._queue_paused = False
        if self._queue_command_thread is not None:
            self._finalize_queue_run('idle', reason=reason, return_to_safe=False)
            return
        if not self._queue_active:
            self._release_queue_lease()
            return

        if hasattr(self._measurement, "halt_run"):
            self._measurement.halt_run()
        if self._queue_current_sample:
            self._set_queue_sample_status(self._queue_current_sample, 'Error')
        self._finalize_queue_run("idle", reason=reason, return_to_safe=True)

    @QtCore.Slot()
    def retry_queue_shutdown(self) -> bool:
        """Retry original journaled terminal closes after the actual worker exits."""
        session, backend = self._queue_native_session, self._measurement_backend
        terminal = getattr(backend, '_queue_terminal_cleanup', None)
        coordinator = getattr(backend, '_queue_coordinator', None)
        if (session is None or not self._uses_native_queue() or self._queue_active
                or self._queue_command_thread is not None or self._measurement.is_active()
                or session._owner is not None or session._lease is None
                or coordinator is None or coordinator.session is not session
                or terminal is None or terminal.coordinator is not coordinator or terminal._close_token is None):
            self.set_status('Queue shutdown retry is unavailable; wait for workers or recover the original unfinished stage.')
            return False
        resources = ('measurement', 'changer', 'af_demag', 'vacuum', 'squid', 'susceptibility')
        if any(self._ownership.owner_of(resource) != 'queue_workflow' for resource in resources):
            self.set_status('Original queue device ownership changed; shutdown retry remains blocked.')
            return False
        try:
            state = session.store.read()
            session.store.verify_history(state)
            stage = state['stage']
            if (state['family'] != 'queue' or state['token'] != session.token or stage is None
                    or stage['token'] != terminal._close_token or stage['plan'] != terminal._close_plan
                    or stage['status'] not in {'pending', 'verified'}
                    or (state['status'] == 'verified' and (terminal.last_root_record is None
                        or state['record'] != terminal.last_root_record.to_dict()))):
                raise RuntimeError('The original journaled shutdown does not match this queue.')
        except Exception as exc:
            self.set_status('Original queue shutdown cannot be retried: ' + str(exc))
            return False
        self._queue_pending_finalize = dict(state='idle', reason='Original queue shutdown completed.',
            return_to_safe=False, clear_plan=True, clear_current=True, clear_command=True)
        self._start_queue_command_worker(None, backend.retry_queue_terminal_settlement, [], recover_on_error=False)
        self.set_flow_state('returning')
        self.set_status('Retrying original connection closure; queue ownership remains held until verification.')
        return True

    def _run_next_queue_command(self) -> None:
        if self._queue_command_thread is not None:
            return
        if not self._queue_active:
            return
        if self._queue_paused:
            self.set_flow_state("paused")
            return
        vacuum_fault = self._queue_vacuum_fault_reason()
        if vacuum_fault:
            if self._queue_current_sample:
                self._set_queue_sample_status(self._queue_current_sample, 'Error')
            self._finalize_queue_run(
                "error",
                reason=f"Queue halted by vacuum fault: {vacuum_fault}",
                return_to_safe=True,
            )
            return
        squid_fault = self._queue_squid_fault_reason()
        if squid_fault:
            if self._queue_current_sample:
                self._set_queue_sample_status(self._queue_current_sample, 'Error')
            self._finalize_queue_run(
                "error",
                reason=f"Queue halted by SQUID communication fault: {squid_fault}",
                return_to_safe=True,
            )
            return
        if self._queue_pos >= len(self._queue_plan):
            if self._queue_meas_pending_start:
                self._run_measure_command()
                return
            self._finalize_queue_run("complete", reason="Queue run complete.")
            return

        if self._queue_meas_pending_start:
            self._run_measure_command()
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

    def queue_measurement_labels(self, sample_name: str) -> list[str]:
        """Resolve the original compiled per-file sequence for this handoff."""
        command = self._queue_current_command
        if (not self._queue_active or command is None or command.command_type != 'Meas'
                or command.sample_name != sample_name or not command.measurement_labels
                or not any(item is command for item in self._queue_plan)):
            raise ValueError('The original active queue measurement command is required.')
        session = self._queue_native_session
        if session is not None:
            state = session.store.read()
            session.store.verify_history(state)
            expected = json.loads(json.dumps([asdict(item) for item in self._queue_plan]))
            if (state['family'] != 'queue' or state['token'] != session.token or state['status'] != 'pending'
                    or state['plan']['commands']['commands'] != expected
                    or self._queue_loaded_command is not command):
                raise ValueError('Restore the original journaled queue sequence and loaded specimen before measurement.')
        labels = list(command.measurement_labels)
        for plan in (self._rockmag_routine_plan, self._thermal_routine_plan):
            if plan is not None and list(plan.to_queue_labels()) != labels:
                raise ValueError('The reviewed routine identity does not match this file measurement sequence.')
        return labels

    def queue_measurement_source(self, sample_name: str):
        """Resolve only the original captured source index and registrations."""
        self.queue_measurement_labels(sample_name)
        indexes = getattr(self, '_queue_source_indexes', {})
        session = self._queue_native_session
        if session is not None:
            state = session.store.read()
            session.store.verify_history(state)
            if state['plan']['commands']['source_indexes'] != indexes:
                raise ValueError('Restore the original journaled index metadata before measurement.')
        source = indexes.get(self._queue_current_command.file_id)
        if source is None:
            return None
        path = Path(source['source_file'])
        if hashlib.sha256(path.read_bytes()).hexdigest() != source['sha256']:
            raise ValueError('The original sample index changed after queue preparation.')
        registrations = SampleIndexRegistrations([SampleIndexRegistration(**entry) for entry in source['entries']])
        return path, registrations

    def queue_measurement_metadata(self, sample_name: str):
        source = self.queue_measurement_source(sample_name)
        if source is None:
            return None
        snapshot = self._queue_source_indexes[self._queue_current_command.file_id]['specimens'][sample_name]
        return restore_specimen_metadata(snapshot)

    def queue_measurement_provenance(self, sample_name: str):
        metadata = self.queue_measurement_metadata(sample_name)
        if metadata is None:
            return None
        command = self._queue_current_command
        source = self._queue_source_indexes[command.file_id]
        provenance = deepcopy(dict(schema='rapidpy.queue_specimen_source.v1', file_id=command.file_id,
            row_id=command.row_id, source_file=source['source_file'], index_sha256=source['sha256'],
            specimen=source['specimens'][sample_name]))
        validate_specimen_provenance(provenance, metadata.meta)
        return provenance

    def _set_queue_sample_status(self, sample_name: str, status: str) -> None:
        command = self._queue_current_command
        setter = getattr(self._sample_queue, 'set_queue_command_status', None)
        if callable(setter) and command is not None and command.command_type == 'Meas':
            setter(command, status)
            return
        method = {'Running': 'start_queue_sample', 'Done': 'mark_queue_sample_done', 'Error': 'set_queue_sample_failed'}[status]
        getattr(self._sample_queue, method)(sample_name)

    def _run_measure_command(self) -> None:
        if self._queue_current_command is None:
            return
        sample_name = self._queue_current_command.sample_name or ""
        if not sample_name:
            self.set_status("Queue measurement command missing sample name; continuing.")
            self._run_next_queue_command()
            return

        if self._queue_native_session is not None and self._queue_loaded_command is not self._queue_current_command:
            self._queue_meas_pending_start = True
            command = self._queue_current_command
            def load():
                backend = self._measurement_backend
                backend.set_measurement_context(is_up=backend.queue_file_direction(command.file_id))
                backend.load_queue_specimen(command.hole, command.sample_name, file_id=command.file_id)
            self._start_queue_command_worker(command, load, [], recover_on_error=False)
            return

        if self._queue_native_session is not None and self._queue_paused:
            self.set_flow_state('paused')
            return
        self._queue_meas_pending_start = False

        self._queue_current_sample = sample_name
        self._set_queue_sample_status(sample_name, 'Running')
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
            self._set_queue_sample_status(sample_name, 'Error')
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
        leases: list[object] = []
        command_type = command.command_type
        try:
            if self._queue_native_session is not None:
                backend = self._measurement_backend
                if command_type == 'InitUp':
                    # VB6 uses this as a preprocessing marker, never a lift home.
                    self._set_queue_position_status(command_type, command)
                    self._queue_advance_timer.start(0)
                    return
                if command_type == 'Holder':
                    action = lambda: backend.measure_queue_holder(command.hole)
                elif command_type == 'Goto':
                    action = backend.park_queue_station
                elif command_type == 'Flip':
                    action = backend.park_queue_station
                else:
                    raise ValueError('Unsupported native queue command: ' + command_type)
                self._start_queue_command_worker(command, action, [], recover_on_error=False)
                return
            leases.append(self.acquire_device("changer", "queue_workflow"))
            if command_type == "Holder" and bool(self.config.susceptibility.enabled):
                leases.append(
                    self.acquire_device("susceptibility", "queue_workflow")
                )

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
            if getattr(self._measurement_backend, 'queue_commands_require_worker', False) is True:
                method = getattr(self._measurement_backend, method_name, None)
                if not callable(method):
                    raise AttributeError(f"Backend does not support required command '{method_name}'.")
                self._start_queue_command_worker(command, lambda: method() if arg is None else method(arg), leases)
                leases = []  # Ownership transfers to the live worker's terminal callback.
                return
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
                self._set_queue_sample_status(self._queue_current_sample, 'Error')
            self._finalize_queue_run(
                "error",
                reason=f"Queue {command.command_type} failed: {exc}",
                return_to_safe=True,
            )
        finally:
            for lease in reversed(leases):
                try:
                    lease.release()
                except Exception:
                    pass

    def _start_queue_command_worker(self, command, action, leases, *, recover_on_error=True):
        if self._queue_command_thread is not None:
            raise RuntimeError('Another queue hardware command has not settled.')
        worker = QueueCommandWorker(self._measurement_backend, action, recover_on_error=recover_on_error)
        thread = QtCore.QThread(self)
        worker.moveToThread(thread)
        thread.started.connect(worker.run)
        worker.settled.connect(worker.deleteLater)
        worker.settled.connect(thread.quit)
        thread.finished.connect(self._queue_command_settled)
        self._queue_command_worker, self._queue_command_thread = worker, thread
        self._queue_worker_command = command
        self._queue_command_leases = list(leases)
        self.set_flow_state('positioning')
        label = command.command_type if command is not None else 'terminal recovery'
        self.set_status(f'Queue {label} running; waiting for hardware completion.')
        thread.start()

    @QtCore.Slot()
    def _queue_command_settled(self):
        if self._queue_command_thread is not None and self._queue_command_thread.isRunning():
            self._queue_settle_timer.start(10)
            return
        self._queue_settle_timer.stop()
        worker, command = self._queue_command_worker, self._queue_worker_command
        if worker is None:
            return
        if self._queue_command_thread is not None:
            self._queue_command_thread.deleteLater()
        self._queue_command_thread = self._queue_command_worker = None
        self._queue_worker_command = None
        for lease in reversed(self._queue_command_leases):
            lease.release()
        self._queue_command_leases = []
        pending, self._queue_pending_finalize = self._queue_pending_finalize, None
        if pending is not None:
            if worker.error:
                pending['reason'] = (pending.get('reason') or 'Queue stopped') + ': ' + worker.error
            if 'recovery remains unverified' in worker.error or command is None and not worker.ok:
                pending['state'] = 'error'
            if self._queue_native_session is not None and command is not None:
                pending['return_to_safe'] = True
            self._finalize_queue_run(**pending)
        elif not worker.ok:
            self._finalize_queue_run('error', reason=f'Queue {command.command_type} failed: {worker.error}', return_to_safe=False)
        else:
            if self._queue_native_session is not None:
                if command.command_type == 'NativeStartup':
                    self._start_queue_command_worker(QueueCommand('NativeLoadPark'),
                        self._measurement_backend.park_queue_station, [], recover_on_error=False)
                    return
                if command.command_type in {'NativeLoadPark', 'Flip'}:
                    self._confirm_native_tray(command)
                    return
                if command.command_type == 'Meas':
                    self._queue_loaded_command = command
                    self._run_measure_command()
                    return
            self._set_queue_position_status(command.command_type, command)
            self._save_queue_state()
            if self._queue_paused:
                self.set_flow_state('paused')
            else:
                self._queue_advance_timer.start(0)

    def _confirm_native_tray(self, command):
        flip = command.command_type == 'Flip'
        message = (f'Place specimens from {command.file_id} in the tray with their arrows reversed.' if flip
                   else 'Load the specimen tray using the registered slots and the selected file orientations.')
        answer = QtWidgets.QMessageBox.question(self, 'Flip specimen tray' if flip else 'Load specimen tray',
            message + '\n\nConfirm the tray is secured and your hands are clear before continuing.',
            QtWidgets.QMessageBox.StandardButton.Yes | QtWidgets.QMessageBox.StandardButton.No,
            QtWidgets.QMessageBox.StandardButton.No)
        if answer != QtWidgets.QMessageBox.StandardButton.Yes:
            self._finalize_queue_run('halted', reason='Operator did not confirm the specimen tray.')
            return
        action = lambda: self._measurement_backend.record_queue_tray_confirmation(
            'flip' if flip else 'load', command.file_id, operator=self.config.general.operator)
        self._start_queue_command_worker(QueueCommand('NativeTrayConfirmed'), action, [], recover_on_error=False)

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

        if self._uses_native_queue():
            required = ['prepare_queue_lifetime', 'start_queue_lifetime', 'finish_queue_lifetime',
                        'park_queue_station', 'record_queue_tray_confirmation', 'queue_file_direction',
                        'load_queue_specimen', 'measure_queue_holder']
            return [name for name in required if not callable(getattr(self._measurement_backend, name, None))]

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
        if self._queue_native_session is not None and command_type in {'Goto', 'NativeTrayConfirmed'}:
            self.set_position('Loading corner')
            self.set_status('Tray confirmation recorded.' if command_type == 'NativeTrayConfirmed'
                            else 'Table parked at the loading corner.')
            return
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
        if self._queue_native_session is not None:
            try:
                session, vacuum = self._queue_native_session, self._vacuum_backend
                root = session.store._queue(session.token)
                session.store.verify_history(root)
                geometry = getattr(self._measurement_backend, '_queue_specimen_geometry', None)
                expected_grip = geometry is not None and not getattr(geometry, 'is_holder', False)
                if (session._lease is None or vacuum._queue_binding is None
                        or vacuum._queue_binding.session is not session or vacuum.output_state_known is not True
                        or vacuum.is_pump_on() is not True or vacuum.is_valve_connected() is not expected_grip
                        or root['profile']['stage_profiles']['vacuum'] != vacuum.queue_station_binding()
                        or (root['stage'] is not None and root['stage']['status'] == 'pending')):
                    return 'Original native queue vacuum state or stage is unverified.'
                return None
            except Exception as exc:
                return str(exc)
        try:
            require_vacuum_ready(
                self._vacuum_backend,
                warn_threshold=float(self.config.vacuum.warn_threshold),
            )
        except Exception as exc:
            return str(exc)
        return None

    def _queue_squid_fault_reason(self) -> str | None:
        if self._queue_native_session is not None:
            # Native connection/reset and instrument commands belong to the
            # original acquisition stage, never the GUI diagnostic adapter.
            return None
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
        command = self._queue_current_command
        if command is not None and command.command_type == 'Meas' and command.sample_name != sample:
            self._set_queue_sample_status(command.sample_name, 'Error')
            self._finalize_queue_run('error', reason='Measurement completion does not match the original queued specimen.')
            return
        if self._queue_pending_finalize is not None and self._queue_command_thread is None:
            pending, self._queue_pending_finalize = self._queue_pending_finalize, None
            self._set_queue_sample_status(sample, 'Error')
            self._finalize_queue_run(**pending)
            return
        had_error = (
            self._measurement.take_last_run_error()
            if hasattr(self._measurement, "take_last_run_error")
            else False
        )
        if aborted:
            self._set_queue_sample_status(sample, 'Error')
            self._finalize_queue_run("halted", reason=f"Queue stopped after sample {sample}.")
            return
        if had_error:
            self._set_queue_sample_status(sample, 'Error')
            self._finalize_queue_run("error", reason=f"Queue sample {sample} failed with an error.")
            return
        self._set_queue_sample_status(sample, 'Done')
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
        if getattr(self, '_shutdown_cleanup_requested', False):
            raise DeviceOwnershipError('Shutdown cleanup is in progress; wait for active hardware owners to finish.')
        backend = getattr(self,"_measurement_backend",None)
        from rapidpy_common.hardware_safety import HardwareSafetyError
        try:
            dc_recovery = resource == 'changer' and owner == 'dc_motors_panel' and getattr(getattr(self, '_dc_motor_backend', None), 'can_recover_pending', False) is True
            vacuum_recovery = resource == 'vacuum' and owner == 'vacuum_panel' and getattr(getattr(self, '_vacuum_backend', None), 'can_recover_pending', False) is True
            irm_recovery = resource == 'af_demag' and owner == 'irm_panel'
            unresolved = any(getattr(backend, name, False) is True for name in ('has_unresolved_hardware_fault', 'has_unresolved_pulse_fault', 'has_unresolved_rotation_fault'))
        except HardwareSafetyError as exc:
            raise DeviceOwnershipError(str(exc)) from exc
        native_reentrant = owner == 'queue_workflow' and self._allows_native_queue_reentrant(resource, owner)
        if resource in {'measurement', 'changer', 'af_demag', 'vacuum', 'squid', 'susceptibility'} and not (irm_recovery or dc_recovery or vacuum_recovery or native_reentrant) and unresolved:
            raise DeviceOwnershipError("Hardware safe state is unverified. Recover the unfinished operation with its original panel or helper before using other controls.")
        return self._ownership.acquire(resource, owner, allow_reentrant=allow_reentrant)

    def acquire_devices(self, resources, owner, *, allow_reentrant=True):
        """Validate each fault guard, then atomically reserve the whole circuit."""
        resources = tuple(resources)
        probes = []
        try:
            for resource in resources:
                probes.append(self.acquire_device(resource, owner, allow_reentrant=allow_reentrant))
        finally:
            for probe in reversed(probes):
                probe.release()
        return self._ownership.acquire_many(resources, owner, allow_reentrant=allow_reentrant)

    def _allows_native_queue_reentrant(self, resource, owner):
        session = self._queue_native_session
        if session is None or owner != 'queue_workflow' or self._ownership.owner_of(resource) != owner:
            return False
        try:
            startup = self._queue_native_startup
            backend = self._measurement_backend
            root = session.store._queue(session.token)
            session.store.verify_history(root)
            resources = ('measurement', 'changer', 'af_demag', 'vacuum', 'squid', 'susceptibility')
            return (self._uses_native_queue() and startup is not None and startup.session is session
                and startup.backend is backend and startup.completed is True
                and backend._queue_instruments is startup.instruments
                and startup.instruments.session is session and startup.instruments.is_original()
                and self._queue_command_thread is None and session._lease is not None and session._owner is None
                and backend._safety_store is session.child_store and root['status'] == 'pending'
                and root['stage'] is not None and root['stage']['status'] in {'verified', 'held'}
                and root['profile']['stage_profiles']['acquisition'] == backend._acquisition_safety_profile()
                and all(root['profile']['stage_profiles'][family] == backend._safety_profile()
                        for family in ('af', 'arm', 'pulse', 'rrm'))
                and all(self._ownership.owner_of(item) == owner for item in resources))
        except Exception:
            return False

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
        self._run_sequence_labels = list(labels if labels is not None else self._sequence_labels)
        self._run_current_idx = 0
        self._run_start_time = datetime.now()
        self._run_timer.start()
        self._update_runtime_display(current_idx=0, running=True)
        self.set_flow_state("running")

    def stop_run(self) -> None:
        """End the measurement run — stop the countdown timer."""
        self._run_timer.stop()
        self._run_sequence_labels = None
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
        self._retry_queue_close_action = dm.addAction('Retry Queue &Shutdown', self.retry_queue_shutdown)
        self._retry_queue_close_action.setEnabled(False)
        self._retry_queue_close_action.setToolTip('Finish closing the original connections after a queue shutdown failure.')
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
        previous_susceptibility = getattr(self, "_susceptibility_backend", None)
        if previous_susceptibility is not None:
            disconnect = getattr(previous_susceptibility, "disconnect", None)
            if callable(disconnect):
                try:
                    disconnect()
                except Exception:
                    pass
        self._susceptibility_backend = build_backend_or_unavailable(
            "Susceptibility bridge",
            build_susceptibility_backend,
            self.config.susceptibility,
            nocomm=nocomm,
        )
        self._dc_motor_backend = build_backend_or_unavailable(
            "DC motors",
            build_dcmotor_backend,
            config=self.config,
            port=(self.config.changer.port or "COM3").strip(),
            baud=int(self.config.changer.baud or 9600),
            nocomm=nocomm,
        )

    # ── No-Comm toggle ────────────────────────────────────────────────────────
    def _operating_mode_change_blocker(self, on):
        if self._shutdown_cleanup_requested or self._has_active_automation() or self._owned_dialog_leases or self._external_process_leases:
            return 'Close hardware diagnostics and finish active work before changing operator or operating mode.'
        from rapidpy_common.hardware_safety import HardwareSafetyStore, default_safety_path, HardwareSafetyError
        try:
            if HardwareSafetyStore(default_safety_path()).pending() is not None and on:
                return 'Recover the unfinished hardware operation before entering No-Comm. Hardware mode remains available for original-panel recovery.'
        except HardwareSafetyError as exc:
            return str(exc)
        return ''

    def _on_nocomm_toggled(self, on: bool) -> None:
        reason = self._operating_mode_change_blocker(on)
        if reason:
            self._nocomm_btn.blockSignals(True)
            self._nocomm_btn.setChecked(bool(self.config.general.nocomm))
            self._nocomm_btn.blockSignals(False)
            self.set_status(reason)
            return
        self.config.general.nocomm = bool(on)
        self.config.save()
        self._rebuild_diagnostic_backends(nocomm=bool(on))
        self._measurement_backend = build_measurement_backend(
            self.config,
            susceptibility_backend=self._susceptibility_backend,
        )
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
        leases = []
        try:
            for resource in ("measurement", "changer", "af_demag", "irm"):
                leases.append(self.acquire_device(resource, "irm_panel", allow_reentrant=False))
            from .manual_treatment import ManualArmTreatment
            manual_arm = None
            if not self.config.general.nocomm:
                manual_arm = ManualArmTreatment(self._measurement_backend, Path(self.config.general.data_dir) / "manual_treatments")
            dialog = IrmArmDialog(self, backend=manual_arm or self._irm_arm_backend, manual_arm=manual_arm)
            dialog.setWindowIcon(self.windowIcon())
            dialog.exec()
        except (DeviceOwnershipError, RuntimeError) as exc:
            QtWidgets.QMessageBox.warning(self, "IRM / ARM Control", str(exc))
        finally:
            for lease in reversed(leases):
                lease.release()

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
            modal=False,
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
        reason = vrm_launch_block_reason(
            active_automation=self._has_active_automation(),
            measurement_owner=self._ownership.owner_of("measurement"),
            squid_owner=self._ownership.owner_of("squid"),
        )
        if reason:
            self.set_status(reason)
            QtWidgets.QMessageBox.warning(self, "VRM launcher", reason)
            return

        owner = "vrm_logger_process"
        leases: list[object] = []
        try:
            leases.append(self.acquire_device("measurement", owner, allow_reentrant=False))
            leases.append(self.acquire_device("squid", owner, allow_reentrant=False))
        except DeviceOwnershipError as exc:
            for lease in reversed(leases):
                lease.release()
            QtWidgets.QMessageBox.warning(self, "VRM launcher", str(exc))
            return

        try:
            self._release_squid_connections_for_vrm()
        except Exception as exc:
            for lease in reversed(leases):
                lease.release()
            message = f"VRM cannot claim the SQUID serial port: {exc}"
            self.set_status(message)
            QtWidgets.QMessageBox.warning(self, "VRM launcher", message)
            return

        sample_id = self._current_sample if self._current_sample != "UNKNOWN" else ""
        data_dir = Path(self.config.general.data_dir) if self.config.general.data_dir else None
        intended_output_dir = data_dir / sample_id if data_dir and sample_id else data_dir
        context = build_vrm_launch_context(
            active_automation=False,
            nocomm=bool(getattr(self.config.general, "nocomm", False)),
            sample_id=sample_id,
            operator=self.config.general.operator,
            intended_output_dir=intended_output_dir,
            software_version=software_version(),
            config_hash=config_fingerprint(self.config),
        )
        try:
            context_path = write_vrm_launch_context(context=context)
        except OSError as exc:
            for lease in reversed(leases):
                lease.release()
            QtWidgets.QMessageBox.warning(
                self,
                "VRM launcher",
                f"Unable to prepare VRM run context: {exc}",
            )
            return
        process = self._launch_external_tool(
            target_path="vrm_logger/main.py",
            module="vrm_logger",
            app_name="VRM Logger",
            env={VRM_CONTEXT_ENV: str(context_path)},
        )
        if process is None:
            for lease in reversed(leases):
                lease.release()
            return
        self._hold_external_process_leases(process, leases)

    def _release_squid_connections_for_vrm(self) -> None:
        """Close retained main-app SQUID clients before handing off the port."""

        backends = (
            ("measurement backend", self._measurement_backend),
            ("diagnostic backend", self._squid_backend),
        )
        seen: set[int] = set()
        for label, backend in backends:
            if backend is None or id(backend) in seen:
                continue
            seen.add(id(backend))
            release = getattr(backend, "release_squid_for_external_tool", None)
            if callable(release):
                release()
                continue
            connected_provider = getattr(backend, "is_connected", None)
            connected = bool(connected_provider()) if callable(connected_provider) else False
            if not connected:
                continue
            disconnect = getattr(backend, "disconnect", None)
            if not callable(disconnect):
                raise RuntimeError(f"{label} is connected and has no disconnect operation")
            disconnect()
            if callable(connected_provider) and bool(connected_provider()):
                raise RuntimeError(f"{label} remained connected after disconnect")

    def _hold_external_process_leases(self, process: object, leases: list[object]) -> None:
        """Keep shared resources reserved until a launched helper exits."""

        key = id(process)
        timer = QtCore.QTimer(self)
        timer.setInterval(500)

        def release_if_finished() -> None:
            poll = getattr(process, "poll", None)
            if not callable(poll) or poll() is None:
                return
            timer.stop()
            for lease in reversed(leases):
                lease.release()
            self._external_process_leases.pop(key, None)
            self.set_status("VRM Logger closed; measurement and SQUID ownership released.")

        timer.timeout.connect(release_if_finished)
        self._external_process_leases[key] = (process, leases, timer)
        timer.start()

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
    ) -> subprocess.Popen | None:
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
                return None

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
            return None

        try:
            process_env = os.environ.copy()
            if env:
                process_env.update(env)
            process = subprocess.Popen(
                list(launch.command),
                cwd=str(launch.cwd) if launch.cwd is not None else None,
                env=process_env,
            )
            self.set_status(f"Launched {app_name} from {launch.source}.")
            return process
        except Exception as exc:  # pragma: no cover - platform/environment dependent
            QtWidgets.QMessageBox.warning(
                self,
                f"{app_name} launcher",
                f"Unable to launch {app_name}: {exc}",
            )
            return None

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
            lease = self.acquire_device(resource, owner, allow_reentrant=False)
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
                self._owned_dialogs[resource] = dlg
                blocked = getattr(dlg, 'shutdown_blocked', None)
                if blocked is not None:
                    blocked.connect(self._diagnostic_shutdown_blocked)

                @QtCore.Slot(object)
                def _release_owner(_obj: object) -> None:
                    self._owned_dialogs.pop(resource, None)
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

    def _defer_window_callback(self, delay_ms, callback):
        """Cancel deferred callbacks when their owning window is deleted."""
        timer = QtCore.QTimer(self)
        timer.setSingleShot(True)
        timer.timeout.connect(callback)
        timer.timeout.connect(timer.deleteLater)
        timer.start(delay_ms)

    def closeEvent(self, event) -> None:
        """Confirm safe shutdown and persist layout settings before closing."""
        if not self._shutdown_cleanup_requested:
            if not self._confirm_shutdown(prompt=True, on_close=True):
                event.ignore()
                return
            self._shutdown_cleanup_requested = True

        for resource, dialog in list(self._owned_dialogs.items()):
            if resource not in self._shutdown_requested_dialogs:
                self._shutdown_requested_dialogs.add(resource)
                dialog.close()
        if not self._shutdown_cleanup_requested:
            event.ignore()
            return
        if (self._queue_native_session is not None and not self._queue_active
                and self._queue_command_thread is None and not self._measurement.is_active()):
            self._diagnostic_shutdown_blocked('Recover the original unfinished native queue before closing.')
            event.ignore()
            return
        if self._owned_dialog_leases or self._has_active_automation():
            event.ignore()
            self.set_status('Shutdown is waiting for hardware workers and diagnostic cleanup.')
            if not self._shutdown_retry_pending:
                self._shutdown_retry_pending = True
                self._shutdown_timer.start(100)
            return

        self._shutdown_timer.stop()
        self._save_layout_state()
        super().closeEvent(event)

    @QtCore.Slot()
    def _retry_shutdown(self):
        self._shutdown_retry_pending = False
        if self._shutdown_cleanup_requested:
            self.close()

    @QtCore.Slot(str)
    def _diagnostic_shutdown_blocked(self, reason):
        self._shutdown_cleanup_requested = False
        self._shutdown_retry_pending = False
        self._shutdown_requested_dialogs.clear()
        self._shutdown_timer.stop()
        self.set_status('Shutdown remains open for diagnostic recovery: ' + reason)

    def _has_active_automation(self) -> bool:
        """Return true when a live measurement or queue run is active."""
        measurement_active = bool(
            hasattr(self._measurement, "is_active") and self._measurement.is_active()
        )
        return (measurement_active or bool(self._queue_active) or self._queue_command_thread is not None
                or self._queue_native_session is not None)

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
        reason = self._operating_mode_change_blocker(False)
        if reason:
            QtWidgets.QMessageBox.warning(
                self,
                "Change Operator",
                reason,
            )
            return
        dialog = LoginDialog(self)
        if dialog.exec() != QtWidgets.QDialog.DialogCode.Accepted:
            return
        reason = self._operating_mode_change_blocker(bool(dialog.nocomm))
        if reason:
            self.set_status(reason)
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
            self._defer_window_callback(350, self._launch_startup_guide_if_enabled)

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
            source_file=str(dialog.source_path.resolve()) if dialog.source_path is not None else '',
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
        retry_action = getattr(self, '_retry_queue_close_action', None)
        if retry_action is not None:
            retry_action.setEnabled(self._queue_native_session is not None and not self._queue_active
                and self._queue_command_thread is None and not self._measurement.is_active())
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
        if self._queue_active or (hasattr(self._measurement, "is_active") and self._measurement.is_active()):
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
        window._defer_window_callback(75, window._fit_window_to_current_screen)
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

    for delay_ms in (0, 150, 350, 750, 1400):
        window._defer_window_callback(delay_ms, _enforce_startup_fit)
    return app.exec()
