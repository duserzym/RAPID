from __future__ import annotations

import csv
import json
import sys
from dataclasses import asdict, dataclass
from pathlib import Path

import pyqtgraph as pg
from PySide6 import QtCore, QtGui, QtWidgets


def _bootstrap_common_imports() -> None:
    root = Path(__file__).resolve().parents[2]
    if str(root) not in sys.path:
        sys.path.insert(0, str(root))


_bootstrap_common_imports()
from rapidpy_common.adwin_af import (  # noqa: E402
    AdwinAFController,
    AdwinBoardConfig,
    AdwinCoilLimits,
    AdwinDenseCaptureRequest,
    AdwinDenseCaptureResult,
    AdwinError,
    AdwinRampRequest,
)
from rapidpy_common.ui import apply_card_shadow, apply_liquid_glass_theme, set_app_icon  # noqa: E402


@dataclass(slots=True)
class CoilTuningConfig:
    axial_res_freq: float = 877.0
    axial_max_ramp: float = 0.5
    axial_max_monitor: float = 0.5
    trans_res_freq: float = 1000.0
    trans_max_ramp: float = 0.5
    trans_max_monitor: float = 0.5


@dataclass(slots=True)
class AutoTuneConfig:
    low_freq: float = 760.0
    high_freq: float = 1050.0
    step_freq: float = 5.0
    hold_ms: int = 500
    io_rate_hz: float = 25000.0


@dataclass(slots=True)
class ClipCaptureConfig:
    sine_freq_hz: float = 877.0
    amplitude_v: float = 0.5
    duration_ms: int = 220
    io_rate_hz: float = 25000.0
    ramp_up_slope_vps: float = 200.0
    ramp_down_slope_vps: float = 200.0
    ramp_down_periods: int = 2


@dataclass(slots=True)
class BackendConfig:
    board_num: int = 1
    bin_folder: str = ""
    boot_file: str = "ADwin9.btl"
    process_file: str = ""
    ramp_dac_chan: int = 1
    monitor_adc_chan: int = 1
    axial_relay_bit: int = 0
    trans_relay_bit: int = 1


@dataclass(slots=True)
class SweepPoint:
    freq_hz: float
    monitor_peak_v: float
    ramp_peak_v: float
    points_per_period: float


@dataclass(slots=True)
class AutoTuneResult:
    best_freq_hz: float
    best_monitor_peak_v: float
    points: list[SweepPoint]


@dataclass(slots=True)
class CaptureEnvelope:
    label: str
    freq_hz: float
    amplitude_v: float
    capture: AdwinDenseCaptureResult


COIL_CONFIG_PATH = Path.home() / ".rapidpy_af_tuner.json"
BACKEND_CONFIG_PATH = Path.home() / ".rapidpy_af_backend.json"
AUTOTUNE_CONFIG_PATH = Path.home() / ".rapidpy_af_autotune.json"
CLIP_CAPTURE_PATH = Path.home() / ".rapidpy_af_clip_capture.json"


def _repo_root() -> Path:
    return Path(__file__).resolve().parents[3]


def _default_process_file() -> str:
    candidate = _repo_root() / "VB6" / "ADwin" / "sineout.T91"
    return str(candidate) if candidate.exists() else ""


def _assets_dir() -> Path:
    return Path(__file__).resolve().parents[1] / "assets"


def _copy_config(obj):
    return type(obj)(**asdict(obj))


def load_dataclass(path: Path, cls):
    if not path.exists():
        return cls()
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
        defaults = asdict(cls())
        for key in defaults:
            if key in data:
                defaults[key] = data[key]
        return cls(**defaults)
    except (OSError, json.JSONDecodeError, TypeError):
        return cls()


def save_dataclass(path: Path, obj) -> None:
    try:
        path.write_text(json.dumps(asdict(obj), indent=2, sort_keys=True), encoding="utf-8")
    except OSError:
        return


class AutoTuneWorker(QtCore.QObject):
    progress = QtCore.Signal(str)
    finished = QtCore.Signal(object)
    failed = QtCore.Signal(str)

    def __init__(
        self,
        backend: BackendConfig,
        limits: AdwinCoilLimits,
        low_freq: float,
        high_freq: float,
        step_freq: float,
        hold_ms: int,
        io_rate_hz: float,
        ramp_peak_v: float,
        monitor_peak_v: float,
        coil: str,
        parent: QtCore.QObject | None = None,
    ) -> None:
        super().__init__(parent)
        self._backend = _copy_config(backend)
        self._limits = limits
        self._low = low_freq
        self._high = high_freq
        self._step = step_freq
        self._hold_ms = hold_ms
        self._io_rate_hz = io_rate_hz
        self._ramp_peak_v = ramp_peak_v
        self._monitor_peak_v = monitor_peak_v
        self._coil = coil
        self._abort = False

    @QtCore.Slot()
    def run(self) -> None:
        try:
            board = AdwinBoardConfig(
                board_num=self._backend.board_num,
                bin_folder=self._backend.bin_folder,
                boot_file=self._backend.boot_file,
                process_file=self._backend.process_file or _default_process_file(),
                ramp_dac_chan=self._backend.ramp_dac_chan,
                monitor_adc_chan=self._backend.monitor_adc_chan,
                axial_relay_bit=self._backend.axial_relay_bit,
                trans_relay_bit=self._backend.trans_relay_bit,
            )
            controller = AdwinAFController(board=board, limits=self._limits)
            duration_s = max(float(self._hold_ms) / 1000.0, 0.001)
            slope = self._ramp_peak_v / duration_s

            freq = self._low
            best_freq = self._low
            best_amp = -1.0
            points: list[SweepPoint] = []

            while freq <= self._high + 1e-9:
                if self._abort:
                    self.progress.emit("Auto-tune aborted.")
                    break

                self.progress.emit(f"Sweeping {freq:.4f} Hz…")
                result = controller.run_ramp(
                    AdwinRampRequest(
                        slope_up=slope,
                        slope_down=slope,
                        peak_monitor_voltage=self._monitor_peak_v,
                        sine_freq_hz=freq,
                        ramp_peak_voltage=self._ramp_peak_v,
                        active_coil=self._coil,
                        ramp_mode=3,
                        hold_ms=self._hold_ms,
                        ramp_down_mode=1,
                        io_rate_hz=self._io_rate_hz,
                        noise_level=5,
                    )
                )
                points.append(
                    SweepPoint(
                        freq_hz=freq,
                        monitor_peak_v=result.monitor_peak_v,
                        ramp_peak_v=result.ramp_peak_v,
                        points_per_period=result.points_per_period,
                    )
                )
                self.progress.emit(
                    f"{freq:.4f} Hz -> monitor_peak={result.monitor_peak_v:.5f} V, "
                    f"ramp_peak={result.ramp_peak_v:.5f} V, {result.points_per_period:.1f} samples/cycle"
                )
                if result.monitor_peak_v > best_amp:
                    best_amp = result.monitor_peak_v
                    best_freq = freq
                freq += self._step

            self.finished.emit(
                AutoTuneResult(
                    best_freq_hz=best_freq,
                    best_monitor_peak_v=max(best_amp, 0.0),
                    points=points,
                )
            )
        except Exception as exc:
            self.failed.emit(str(exc))

    def abort(self) -> None:
        self._abort = True


class DenseCaptureWorker(QtCore.QObject):
    progress = QtCore.Signal(str)
    capture_ready = QtCore.Signal(object)
    finished = QtCore.Signal()
    failed = QtCore.Signal(str)

    def __init__(
        self,
        backend: BackendConfig,
        limits: AdwinCoilLimits,
        coil: str,
        freq_hz: float,
        amplitude_v: float,
        duration_ms: int,
        io_rate_hz: float,
        ramp_up_slope_vps: float,
        ramp_down_slope_vps: float,
        ramp_down_periods: int,
        label: str,
        parent: QtCore.QObject | None = None,
    ) -> None:
        super().__init__(parent)
        self._backend = _copy_config(backend)
        self._limits = limits
        self._coil = coil
        self._freq_hz = freq_hz
        self._amplitude_v = amplitude_v
        self._duration_ms = duration_ms
        self._io_rate_hz = io_rate_hz
        self._ramp_up_slope_vps = ramp_up_slope_vps
        self._ramp_down_slope_vps = ramp_down_slope_vps
        self._ramp_down_periods = ramp_down_periods
        self._label = label
        self._stop = False

    @QtCore.Slot()
    def run(self) -> None:
        try:
            board = AdwinBoardConfig(
                board_num=self._backend.board_num,
                bin_folder=self._backend.bin_folder,
                boot_file=self._backend.boot_file,
                process_file=self._backend.process_file or _default_process_file(),
                ramp_dac_chan=self._backend.ramp_dac_chan,
                monitor_adc_chan=self._backend.monitor_adc_chan,
                axial_relay_bit=self._backend.axial_relay_bit,
                trans_relay_bit=self._backend.trans_relay_bit,
            )
            controller = AdwinAFController(board=board, limits=self._limits)
            controller.set_af_relays(self._coil, one_chan_on=True)
            self.progress.emit(
                f"{self._label}: board-timed capture at {self._freq_hz:.3f} Hz, "
                f"{self._amplitude_v:.3f} V, {self._io_rate_hz:.0f} Hz IO rate"
            )
            capture = controller.run_dense_loopback(
                AdwinDenseCaptureRequest(
                    sine_freq_hz=self._freq_hz,
                    amplitude_v=self._amplitude_v,
                    io_rate_hz=self._io_rate_hz,
                    duration_s=max(float(self._duration_ms) / 1000.0, 0.05),
                    dac_chan=board.ramp_dac_chan,
                    adc_chan=board.monitor_adc_chan,
                    noise_level=5,
                    ramp_mode=3,
                    ramp_up_slope_vps=self._ramp_up_slope_vps,
                    ramp_down_slope_vps=self._ramp_down_slope_vps,
                    ramp_down_periods=max(1, self._ramp_down_periods),
                ),
                should_stop=lambda: self._stop,
            )
            self.capture_ready.emit(
                CaptureEnvelope(
                    label=self._label,
                    freq_hz=self._freq_hz,
                    amplitude_v=self._amplitude_v,
                    capture=capture,
                )
            )
            self.finished.emit()
        except Exception as exc:
            self.failed.emit(str(exc))

    def stop(self) -> None:
        self._stop = True


class MainWindow(QtWidgets.QMainWindow):
    def __init__(self) -> None:
        super().__init__()
        self.setWindowTitle("RapidPy AF Tuner")

        self._coil_config = load_dataclass(COIL_CONFIG_PATH, CoilTuningConfig)
        self._runtime_coil_config = _copy_config(self._coil_config)
        self._backend_config = load_dataclass(BACKEND_CONFIG_PATH, BackendConfig)
        if not self._backend_config.process_file:
            self._backend_config.process_file = _default_process_file()
        self._autotune_config = load_dataclass(AUTOTUNE_CONFIG_PATH, AutoTuneConfig)
        self._clip_capture_config = load_dataclass(CLIP_CAPTURE_PATH, ClipCaptureConfig)

        self._ctrl: AdwinAFController | None = None
        self._connected = False
        self._last_version: int = 0
        self._worker_thread: QtCore.QThread | None = None
        self._worker: QtCore.QObject | None = None
        self._task_kind: str | None = None
        self._queued_capture_freq: float | None = None

        self._wave_time: list[float] = []
        self._wave_dac: list[float] = []
        self._wave_adc: list[float] = []
        self._sweep_freqs: list[float] = []
        self._sweep_monitor: list[float] = []
        self._sweep_ramp: list[float] = []
        self._last_export_kind = "waveform"

        self._build_ui()
        self._load_into_widgets()
        self._set_coil_locked(False, announce=False)
        self._set_backend_status("Not connected", "#6c5a53")
        self._set_comm_summary("Communication check not run yet.", "#6c5a53")
        self._set_relay_status("Relay state unknown.", "#6c5a53")
        self._reset_plot_view()

        compact_font = QtGui.QFont(self.font())
        compact_size = compact_font.pointSizeF()
        if compact_size > 0:
            compact_font.setPointSizeF(max(9.0, compact_size - 0.4))
            self.setFont(compact_font)

        set_app_icon(self, "af_tuner_icon.ico", _assets_dir())
        self.setMinimumSize(1480, 860)
        self.resize(1680, 940)

    # ------------------------------------------------------------------
    # UI construction
    # ------------------------------------------------------------------
    def _build_ui(self) -> None:
        root = QtWidgets.QWidget(self)
        self.setCentralWidget(root)
        layout = QtWidgets.QVBoxLayout(root)
        layout.setContentsMargins(16, 16, 16, 16)
        layout.setSpacing(12)

        splitter = QtWidgets.QSplitter(QtCore.Qt.Orientation.Horizontal)
        splitter.setChildrenCollapsible(False)
        splitter.setHandleWidth(8)
        layout.addWidget(splitter, stretch=1)

        left_scroll = QtWidgets.QScrollArea()
        left_scroll.setObjectName("panelScroll")
        left_scroll.setWidgetResizable(True)
        left_scroll.setFrameShape(QtWidgets.QFrame.Shape.NoFrame)
        left_scroll.setHorizontalScrollBarPolicy(QtCore.Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        left_scroll.setVerticalScrollBarPolicy(QtCore.Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        left_scroll.setMinimumHeight(100)

        left_host = QtWidgets.QWidget()
        left_layout = QtWidgets.QVBoxLayout(left_host)
        left_layout.setContentsMargins(0, 0, 6, 0)
        left_layout.setSpacing(12)

        tuner_card = self._build_tuner_card()
        backend_card = self._build_backend_card()
        console_card = self._build_console_card()
        for card in (tuner_card, backend_card, console_card):
            apply_card_shadow(card)
            left_layout.addWidget(card)
        left_layout.addStretch(1)
        left_scroll.setWidget(left_host)

        plot_card = self._build_plot_card()
        apply_card_shadow(plot_card)

        splitter.addWidget(left_scroll)
        splitter.addWidget(plot_card)
        splitter.setStretchFactor(0, 0)
        splitter.setStretchFactor(1, 1)
        splitter.setSizes([720, 900])

        self.setStyleSheet(
            self.styleSheet()
            + """
            QSplitter::handle { background: rgba(122, 2, 25, 0.08); border-radius: 4px; }
            QSplitter::handle:hover { background: rgba(122, 2, 25, 0.18); }
            QPlainTextEdit#console {
                color: #d5c8ab;
                border: 1px solid rgba(214, 181, 112, 0.22);
            }
            QLabel#hintText {
                color: #6c5a53;
                font-size: 10px;
            }
            QLabel#plotSummary {
                background: rgba(255, 255, 255, 0.76);
                border: 1px solid rgba(122, 2, 25, 0.12);
                border-radius: 16px;
                padding: 8px 10px;
                color: #4d3a39;
            }
            QFrame#card QPushButton {
                background: rgba(255, 255, 255, 0.96);
                color: #2f2827;
                border: 1px solid rgba(122, 2, 25, 0.48);
                border-radius: 12px;
                padding: 8px 14px;
                font-weight: 600;
                min-height: 36px;
            }
            QFrame#card QPushButton:hover {
                background: rgba(255, 252, 246, 1.0);
                border: 1px solid rgba(122, 2, 25, 0.68);
            }
            QFrame#card QPushButton:pressed {
                background: rgba(237, 228, 215, 0.98);
            }
            QFrame#card QPushButton:disabled {
                background: rgba(233, 228, 217, 0.92);
                color: #8b7a73;
                border: 1px solid rgba(122, 2, 25, 0.18);
            }
            QFrame#card QPushButton#accent {
                background: #1a237e;
                color: #fff9eb;
                border: 1px solid rgba(255, 255, 255, 0.22);
                font-weight: 700;
            }
            QFrame#card QPushButton#accent:hover {
                background: #2433a1;
            }
            QFrame#card QPushButton#accent:pressed {
                background: #121a61;
            }
            QFrame#card QPushButton#accent:disabled {
                background: rgba(26, 35, 126, 0.42);
                color: rgba(255, 249, 235, 0.76);
            }
            QRadioButton#coilChoice {
                background: rgba(255, 255, 255, 0.9);
                color: #2f2827;
                border: 1px solid rgba(122, 2, 25, 0.22);
                border-radius: 14px;
                padding: 10px 14px;
                font-weight: 700;
            }
            QRadioButton#coilChoice:hover {
                background: rgba(255, 255, 255, 0.98);
                border: 1px solid rgba(26, 35, 126, 0.34);
            }
            QRadioButton#coilChoice:checked {
                background: #1a237e;
                color: #fff9eb;
                border: 1px solid rgba(255, 255, 255, 0.28);
            }
            QRadioButton#coilChoice:checked:hover {
                background: #2433a1;
            }
            QRadioButton#coilChoice:disabled {
                background: rgba(233, 228, 217, 0.92);
                color: #71615b;
                border: 1px solid rgba(122, 2, 25, 0.12);
            }
            QRadioButton#coilChoice::indicator {
                width: 0px;
                height: 0px;
            }
            QCheckBox#lockChoice {
                background: rgba(255, 255, 255, 0.92);
                color: #2f2827;
                border: 1px solid rgba(122, 2, 25, 0.3);
                border-radius: 14px;
                padding: 10px 14px;
                font-weight: 700;
            }
            QCheckBox#lockChoice:hover {
                border: 1px solid rgba(122, 2, 25, 0.48);
                background: rgba(255, 255, 255, 0.98);
            }
            QCheckBox#lockChoice:checked {
                background: #7a0219;
                color: #fff9eb;
                border: 1px solid rgba(255, 255, 255, 0.26);
            }
            QCheckBox#lockChoice:disabled {
                background: rgba(233, 228, 217, 0.92);
                color: #71615b;
                border: 1px solid rgba(122, 2, 25, 0.12);
            }
            QCheckBox#lockChoice::indicator {
                width: 0px;
                height: 0px;
            }
            QFrame#card QPushButton#connectAction {
                background: #7a0219;
                color: #fff9eb;
                border: 1px solid rgba(255, 255, 255, 0.22);
                font-weight: 700;
            }
            QFrame#card QPushButton#connectAction:hover {
                background: #95132a;
            }
            QFrame#card QPushButton#connectAction:pressed {
                background: #5f0214;
            }
            QFrame#card QPushButton#connectAction:disabled {
                background: rgba(122, 2, 25, 0.38);
                color: rgba(255, 249, 235, 0.74);
            }
            """
        )

    def _build_tuner_card(self) -> QtWidgets.QFrame:
        frame = QtWidgets.QFrame()
        frame.setObjectName("card")
        layout = QtWidgets.QVBoxLayout(frame)
        layout.setContentsMargins(18, 18, 18, 18)
        layout.setSpacing(12)

        title = QtWidgets.QLabel("AF Coil Tuner")
        title.setObjectName("title")
        subtitle = QtWidgets.QLabel(
            "Migrated from VB6 AF Tuner / ClipTest workflow: runtime apply/save, relay-aware tuning, and board-timed waveform capture."
        )
        subtitle.setObjectName("subtitle")
        subtitle.setWordWrap(True)
        layout.addWidget(title)
        layout.addWidget(subtitle)

        header_row = QtWidgets.QHBoxLayout()
        self.coil_group = QtWidgets.QButtonGroup(self)
        self.axial_radio = QtWidgets.QRadioButton("Axial Coil")
        self.trans_radio = QtWidgets.QRadioButton("Transverse Coil")
        self.axial_radio.setObjectName("coilChoice")
        self.trans_radio.setObjectName("coilChoice")
        self.axial_radio.setMinimumHeight(42)
        self.trans_radio.setMinimumHeight(42)
        self.axial_radio.setChecked(True)
        self.coil_group.addButton(self.axial_radio)
        self.coil_group.addButton(self.trans_radio)
        self.chk_lock_coils = QtWidgets.QCheckBox("Lock Coil Selection")
        self.chk_lock_coils.setObjectName("lockChoice")
        self.chk_lock_coils.setMinimumHeight(42)
        header_row.addWidget(self.axial_radio)
        header_row.addWidget(self.trans_radio)
        header_row.addStretch(1)
        header_row.addWidget(self.chk_lock_coils)
        layout.addLayout(header_row)

        self.status = QtWidgets.QLabel("Ready")
        self.status.setObjectName("valuePill")
        self.status.setAlignment(QtCore.Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(self.status)

        quick_comm_row = QtWidgets.QGridLayout()
        quick_comm_row.setHorizontalSpacing(8)
        quick_comm_row.setVerticalSpacing(8)
        quick_comm_row.setColumnStretch(0, 1)
        quick_comm_row.setColumnStretch(1, 1)
        self.quick_backend_status = QtWidgets.QLabel("Not connected")
        self.quick_backend_status.setObjectName("valuePill")
        self.quick_backend_status.setAlignment(QtCore.Qt.AlignmentFlag.AlignCenter)
        self.quick_backend_status.setWordWrap(True)
        self.quick_backend_status.setMinimumWidth(0)
        self.quick_backend_status.setSizePolicy(
            QtWidgets.QSizePolicy.Policy.Ignored,
            QtWidgets.QSizePolicy.Policy.Preferred,
        )
        self.quick_relay_status = QtWidgets.QLabel("Relay state unknown")
        self.quick_relay_status.setObjectName("valuePill")
        self.quick_relay_status.setAlignment(QtCore.Qt.AlignmentFlag.AlignCenter)
        self.quick_relay_status.setWordWrap(True)
        self.quick_relay_status.setMinimumWidth(0)
        self.quick_relay_status.setSizePolicy(
            QtWidgets.QSizePolicy.Policy.Ignored,
            QtWidgets.QSizePolicy.Policy.Preferred,
        )
        self.quick_connect_btn = QtWidgets.QPushButton("Connect / Boot")
        self.quick_connect_btn.setObjectName("connectAction")
        self.quick_connect_btn.setMinimumHeight(38)
        self.quick_check_btn = QtWidgets.QPushButton("Check Comm")
        self.quick_check_btn.setMinimumHeight(38)
        quick_comm_row.addWidget(self.quick_backend_status, 0, 0, 1, 2)
        quick_comm_row.addWidget(self.quick_relay_status, 1, 0, 1, 2)
        quick_comm_row.addWidget(self.quick_connect_btn, 2, 0)
        quick_comm_row.addWidget(self.quick_check_btn, 2, 1)
        layout.addLayout(quick_comm_row)

        body_row = QtWidgets.QHBoxLayout()
        body_row.setSpacing(12)
        left_column = QtWidgets.QVBoxLayout()
        left_column.setSpacing(12)
        right_column = QtWidgets.QVBoxLayout()
        right_column.setSpacing(12)

        tuning_box = QtWidgets.QGroupBox("Resonance And Voltage Limits")
        tuning_form = QtWidgets.QFormLayout(tuning_box)
        tuning_form.setSpacing(8)

        self.old_freq = QtWidgets.QDoubleSpinBox()
        self.old_freq.setReadOnly(True)
        self.old_freq.setButtonSymbols(QtWidgets.QAbstractSpinBox.ButtonSymbols.NoButtons)
        self.old_freq.setDecimals(4)
        self.old_freq.setRange(0.0, 250000.0)
        self.new_freq = QtWidgets.QDoubleSpinBox()
        self.new_freq.setRange(0.001, 250000.0)
        self.new_freq.setDecimals(4)

        self.old_ramp = QtWidgets.QDoubleSpinBox()
        self.old_ramp.setReadOnly(True)
        self.old_ramp.setButtonSymbols(QtWidgets.QAbstractSpinBox.ButtonSymbols.NoButtons)
        self.old_ramp.setDecimals(4)
        self.old_ramp.setRange(0.0, 10.0)
        self.new_ramp = QtWidgets.QDoubleSpinBox()
        self.new_ramp.setRange(0.001, 10.0)
        self.new_ramp.setDecimals(4)

        self.old_monitor = QtWidgets.QDoubleSpinBox()
        self.old_monitor.setReadOnly(True)
        self.old_monitor.setButtonSymbols(QtWidgets.QAbstractSpinBox.ButtonSymbols.NoButtons)
        self.old_monitor.setDecimals(4)
        self.old_monitor.setRange(0.0, 10.0)
        self.new_monitor = QtWidgets.QDoubleSpinBox()
        self.new_monitor.setRange(0.001, 10.0)
        self.new_monitor.setDecimals(4)

        tuning_form.addRow("Saved Resonance Freq", self.old_freq)
        tuning_form.addRow("Runtime Resonance Freq", self.new_freq)
        tuning_form.addRow("Saved Max Ramp", self.old_ramp)
        tuning_form.addRow("Runtime Max Ramp", self.new_ramp)
        tuning_form.addRow("Saved Max Monitor", self.old_monitor)
        tuning_form.addRow("Runtime Max Monitor", self.new_monitor)
        left_column.addWidget(tuning_box)

        tuning_actions = QtWidgets.QWidget()
        buttons = QtWidgets.QGridLayout(tuning_actions)
        buttons.setContentsMargins(0, 0, 0, 0)
        buttons.setHorizontalSpacing(8)
        buttons.setVerticalSpacing(8)
        self.apply_freq_btn = QtWidgets.QPushButton("Apply Freq")
        self.apply_volt_btn = QtWidgets.QPushButton("Apply Max Volt")
        self.save_freq_btn = QtWidgets.QPushButton("Save Freq")
        self.save_volt_btn = QtWidgets.QPushButton("Save Max Volt")
        buttons.addWidget(self.apply_freq_btn, 0, 0)
        buttons.addWidget(self.apply_volt_btn, 0, 1)
        buttons.addWidget(self.save_freq_btn, 1, 0)
        buttons.addWidget(self.save_volt_btn, 1, 1)
        left_column.addWidget(tuning_actions)

        autotune_box = QtWidgets.QGroupBox("Auto-Tune Sweep")
        autotune_form = QtWidgets.QFormLayout(autotune_box)
        self.low_freq = QtWidgets.QDoubleSpinBox()
        self.low_freq.setRange(0.001, 5000.0)
        self.low_freq.setDecimals(4)
        self.high_freq = QtWidgets.QDoubleSpinBox()
        self.high_freq.setRange(0.001, 5000.0)
        self.high_freq.setDecimals(4)
        self.step_freq = QtWidgets.QDoubleSpinBox()
        self.step_freq.setRange(0.001, 1000.0)
        self.step_freq.setDecimals(4)
        self.hold_ms = QtWidgets.QSpinBox()
        self.hold_ms.setRange(0, 60000)
        self.io_rate = QtWidgets.QDoubleSpinBox()
        self.io_rate.setRange(500.0, 100000.0)
        self.io_rate.setDecimals(1)
        autotune_form.addRow("Sweep Low", self.low_freq)
        autotune_form.addRow("Sweep High", self.high_freq)
        autotune_form.addRow("Step Size", self.step_freq)
        autotune_form.addRow("Peak Hang (ms)", self.hold_ms)
        autotune_form.addRow("Board IO Rate (Hz)", self.io_rate)
        right_column.addWidget(autotune_box)

        autotune_actions = QtWidgets.QWidget()
        autotune_actions_layout = QtWidgets.QVBoxLayout(autotune_actions)
        autotune_actions_layout.setContentsMargins(0, 0, 0, 0)
        autotune_actions_layout.setSpacing(8)
        self.autotune_hint = QtWidgets.QLabel(
            "Axial tuning often lands near 877 Hz. At 25 kHz IO rate that gives about 28.5 samples/cycle, which is dense enough for a readable board-timed preview."
        )
        self.autotune_hint.setObjectName("hintText")
        self.autotune_hint.setWordWrap(True)
        autotune_actions_layout.addWidget(self.autotune_hint)

        self.auto_tune_btn = QtWidgets.QPushButton("Start Auto-Tune")
        self.auto_tune_btn.setObjectName("accent")
        autotune_actions_layout.addWidget(self.auto_tune_btn)
        right_column.addWidget(autotune_actions)

        capture_box = QtWidgets.QGroupBox("Clip / Firing Capture")
        capture_form = QtWidgets.QFormLayout(capture_box)
        self.clip_freq = QtWidgets.QDoubleSpinBox()
        self.clip_freq.setRange(1.0, 5000.0)
        self.clip_freq.setDecimals(4)
        self.clip_amp = QtWidgets.QDoubleSpinBox()
        self.clip_amp.setRange(0.001, 10.0)
        self.clip_amp.setDecimals(4)
        self.clip_duration_ms = QtWidgets.QSpinBox()
        self.clip_duration_ms.setRange(50, 60000)
        self.clip_io_rate = QtWidgets.QDoubleSpinBox()
        self.clip_io_rate.setRange(500.0, 100000.0)
        self.clip_io_rate.setDecimals(1)
        self.clip_ramp_up = QtWidgets.QDoubleSpinBox()
        self.clip_ramp_up.setRange(0.1, 100000.0)
        self.clip_ramp_up.setDecimals(2)
        self.clip_ramp_down = QtWidgets.QDoubleSpinBox()
        self.clip_ramp_down.setRange(0.1, 100000.0)
        self.clip_ramp_down.setDecimals(2)
        self.clip_down_periods = QtWidgets.QSpinBox()
        self.clip_down_periods.setRange(1, 20)
        capture_form.addRow("Capture Freq", self.clip_freq)
        capture_form.addRow("Target Amp", self.clip_amp)
        capture_form.addRow("Steady Time (ms)", self.clip_duration_ms)
        capture_form.addRow("Capture IO Rate", self.clip_io_rate)
        capture_form.addRow("Ramp Up Slope", self.clip_ramp_up)
        capture_form.addRow("Ramp Down Slope", self.clip_ramp_down)
        capture_form.addRow("Ramp Down Periods", self.clip_down_periods)
        left_column.addWidget(capture_box)

        capture_actions = QtWidgets.QWidget()
        capture_actions_layout = QtWidgets.QVBoxLayout(capture_actions)
        capture_actions_layout.setContentsMargins(0, 0, 0, 0)
        capture_actions_layout.setSpacing(8)
        self.capture_hint = QtWidgets.QLabel(
            "This uses the same ADwin-side dense capture path as the comms tester, so the waveform plot reflects board timing rather than host polling."
        )
        self.capture_hint.setObjectName("hintText")
        self.capture_hint.setWordWrap(True)
        capture_actions_layout.addWidget(self.capture_hint)

        self.capture_btn = QtWidgets.QPushButton("Run Clip / Firing Capture")
        self.capture_btn.setObjectName("accent")
        capture_actions_layout.addWidget(self.capture_btn)
        right_column.addWidget(capture_actions)

        left_column.addStretch(1)
        right_column.addStretch(1)
        body_row.addLayout(left_column, 1)
        body_row.addLayout(right_column, 1)
        layout.addLayout(body_row)

        self.coil_group.buttonClicked.connect(self._on_coil_changed)
        self.chk_lock_coils.toggled.connect(self._on_lock_toggled)
        self.quick_connect_btn.clicked.connect(self._connect_backend)
        self.quick_check_btn.clicked.connect(self._check_communication)
        self.apply_freq_btn.clicked.connect(self._apply_freq)
        self.apply_volt_btn.clicked.connect(self._apply_volts)
        self.save_freq_btn.clicked.connect(self._save_freq)
        self.save_volt_btn.clicked.connect(self._save_volts)
        self.auto_tune_btn.clicked.connect(self._toggle_autotune)
        self.capture_btn.clicked.connect(self._toggle_capture)

        return frame

    def _build_backend_card(self) -> QtWidgets.QFrame:
        frame = QtWidgets.QFrame()
        frame.setObjectName("card")
        layout = QtWidgets.QVBoxLayout(frame)
        layout.setContentsMargins(18, 18, 18, 18)
        layout.setSpacing(12)

        title = QtWidgets.QLabel("ADwin Backend")
        title.setObjectName("title")
        layout.addWidget(title)

        self.backend_status = QtWidgets.QLabel("Not connected")
        self.backend_status.setObjectName("valuePill")
        self.backend_status.setAlignment(QtCore.Qt.AlignmentFlag.AlignCenter)
        self.backend_status.setWordWrap(True)
        self.backend_status.setMinimumWidth(0)
        self.backend_status.setSizePolicy(
            QtWidgets.QSizePolicy.Policy.Ignored,
            QtWidgets.QSizePolicy.Policy.Preferred,
        )
        layout.addWidget(self.backend_status)

        self.comm_summary = QtWidgets.QLabel("Communication check not run yet.")
        self.comm_summary.setObjectName("valuePill")
        self.comm_summary.setAlignment(QtCore.Qt.AlignmentFlag.AlignCenter)
        self.comm_summary.setWordWrap(True)
        self.comm_summary.setMinimumWidth(0)
        self.comm_summary.setSizePolicy(
            QtWidgets.QSizePolicy.Policy.Ignored,
            QtWidgets.QSizePolicy.Policy.Preferred,
        )
        layout.addWidget(self.comm_summary)

        self.relay_status = QtWidgets.QLabel("Relay state unknown.")
        self.relay_status.setObjectName("valuePill")
        self.relay_status.setAlignment(QtCore.Qt.AlignmentFlag.AlignCenter)
        self.relay_status.setWordWrap(True)
        self.relay_status.setMinimumWidth(0)
        self.relay_status.setSizePolicy(
            QtWidgets.QSizePolicy.Policy.Ignored,
            QtWidgets.QSizePolicy.Policy.Preferred,
        )
        layout.addWidget(self.relay_status)

        form = QtWidgets.QFormLayout()
        form.setSpacing(8)
        self.board_num = QtWidgets.QSpinBox()
        self.board_num.setRange(1, 64)
        self.bin_folder = QtWidgets.QLineEdit()
        self.boot_file = QtWidgets.QLineEdit()
        self.process_file = QtWidgets.QLineEdit()
        self.ramp_chan = QtWidgets.QSpinBox()
        self.ramp_chan.setRange(1, 128)
        self.monitor_chan = QtWidgets.QSpinBox()
        self.monitor_chan.setRange(1, 128)
        self.axial_relay_bit = QtWidgets.QSpinBox()
        self.axial_relay_bit.setRange(0, 5)
        self.trans_relay_bit = QtWidgets.QSpinBox()
        self.trans_relay_bit.setRange(0, 5)

        self.btn_browse_bin = QtWidgets.QPushButton("Browse…")
        self.btn_browse_boot = QtWidgets.QPushButton("Browse…")
        self.btn_browse_process = QtWidgets.QPushButton("Browse…")
        self.btn_browse_bin.setMinimumWidth(96)
        self.btn_browse_boot.setMinimumWidth(96)
        self.btn_browse_process.setMinimumWidth(96)

        bin_row = QtWidgets.QHBoxLayout()
        bin_row.addWidget(self.bin_folder)
        bin_row.addWidget(self.btn_browse_bin)
        boot_row = QtWidgets.QHBoxLayout()
        boot_row.addWidget(self.boot_file)
        boot_row.addWidget(self.btn_browse_boot)
        proc_row = QtWidgets.QHBoxLayout()
        proc_row.addWidget(self.process_file)
        proc_row.addWidget(self.btn_browse_process)

        form.addRow("Board #", self.board_num)
        form.addRow("Bin Folder", bin_row)
        form.addRow("Boot File", boot_row)
        form.addRow("Process File", proc_row)
        form.addRow("Ramp DAC Chan", self.ramp_chan)
        form.addRow("Monitor ADC Chan", self.monitor_chan)
        form.addRow("Axial Relay Bit", self.axial_relay_bit)
        form.addRow("Trans Relay Bit", self.trans_relay_bit)
        layout.addLayout(form)

        buttons = QtWidgets.QGridLayout()
        self.connect_btn = QtWidgets.QPushButton("Connect / Boot")
        self.connect_btn.setObjectName("connectAction")
        self.check_comm_btn = QtWidgets.QPushButton("Check Communication")
        self.apply_relay_btn = QtWidgets.QPushButton("Apply Active Relay")
        self.relays_off_btn = QtWidgets.QPushButton("All Relays Off")
        self.save_backend_btn = QtWidgets.QPushButton("Save Backend")
        buttons.addWidget(self.connect_btn, 0, 0, 1, 2)
        buttons.addWidget(self.check_comm_btn, 1, 0, 1, 2)
        buttons.addWidget(self.apply_relay_btn, 2, 0)
        buttons.addWidget(self.relays_off_btn, 2, 1)
        buttons.addWidget(self.save_backend_btn, 3, 0, 1, 2)
        layout.addLayout(buttons)

        hint = QtWidgets.QLabel(
            "Use the compiled ADwin process file `sineout.T91` for both tuning sweeps and dense firing capture. Saving backend settings persists the runtime mapping for the next session."
        )
        hint.setObjectName("hintText")
        hint.setWordWrap(True)
        layout.addWidget(hint)

        self.btn_browse_bin.clicked.connect(self._browse_bin_folder)
        self.btn_browse_boot.clicked.connect(self._browse_boot_file)
        self.btn_browse_process.clicked.connect(self._browse_process_file)
        self.connect_btn.clicked.connect(self._connect_backend)
        self.check_comm_btn.clicked.connect(self._check_communication)
        self.apply_relay_btn.clicked.connect(self._apply_relays)
        self.relays_off_btn.clicked.connect(self._all_relays_off)
        self.save_backend_btn.clicked.connect(self._save_backend)

        return frame

    def _build_console_card(self) -> QtWidgets.QFrame:
        frame = QtWidgets.QFrame()
        frame.setObjectName("card")
        layout = QtWidgets.QVBoxLayout(frame)
        layout.setContentsMargins(18, 18, 18, 18)
        layout.setSpacing(10)

        label = QtWidgets.QLabel("Run Console")
        label.setObjectName("title")
        layout.addWidget(label)

        self.console = QtWidgets.QPlainTextEdit()
        self.console.setObjectName("console")
        self.console.setReadOnly(True)
        self.console.setMinimumHeight(130)
        self.console.setSizePolicy(
            QtWidgets.QSizePolicy.Policy.Preferred,
            QtWidgets.QSizePolicy.Policy.Expanding,
        )
        layout.addWidget(self.console, stretch=1)
        return frame

    def _build_plot_card(self) -> QtWidgets.QFrame:
        frame = QtWidgets.QFrame()
        frame.setObjectName("card")
        layout = QtWidgets.QVBoxLayout(frame)
        layout.setContentsMargins(18, 18, 18, 18)
        layout.setSpacing(12)

        top_row = QtWidgets.QHBoxLayout()
        title = QtWidgets.QLabel("Board Capture And Sweep Results")
        title.setObjectName("title")
        top_row.addWidget(title)
        top_row.addStretch(1)
        self.save_plot_btn = QtWidgets.QPushButton("Save Plot")
        self.save_csv_btn = QtWidgets.QPushButton("Save CSV")
        self.clear_plot_btn = QtWidgets.QPushButton("Clear Plot")
        top_row.addWidget(self.save_plot_btn)
        top_row.addWidget(self.save_csv_btn)
        top_row.addWidget(self.clear_plot_btn)
        layout.addLayout(top_row)

        help_lbl = QtWidgets.QLabel(
            "Top chart shows board-timed firing capture (DAC output and ADC monitor). Bottom chart shows the auto-tune frequency sweep response. Left-drag zooms, right-drag pans, and double-click resets the plot under the cursor."
        )
        help_lbl.setObjectName("hintText")
        help_lbl.setWordWrap(True)
        layout.addWidget(help_lbl)

        self.plot_summary = QtWidgets.QLabel(
            "No capture loaded yet. Auto-tune fills the response chart and clip/firing capture fills the waveform chart."
        )
        self.plot_summary.setObjectName("plotSummary")
        self.plot_summary.setWordWrap(True)
        layout.addWidget(self.plot_summary)

        self.plot_splitter = QtWidgets.QSplitter(QtCore.Qt.Orientation.Vertical)
        self.plot_splitter.setChildrenCollapsible(False)
        self.plot_splitter.setHandleWidth(6)
        layout.addWidget(self.plot_splitter, stretch=1)

        self.wave_plot = self._create_plot_widget("Time", "Voltage", "ms", "V")
        self.sweep_plot = self._create_plot_widget("Frequency", "Peak Voltage", "Hz", "V")
        self.sweep_plot.setMinimumHeight(220)
        self.plot_splitter.addWidget(self.wave_plot)
        self.plot_splitter.addWidget(self.sweep_plot)
        self.plot_splitter.setSizes([520, 260])

        self.wave_plot.addLegend(offset=(12, 12))
        self.wave_dac_curve = self.wave_plot.plot(name="DAC out", pen=pg.mkPen("#0f766e", width=1.8))
        self.wave_adc_curve = self.wave_plot.plot(name="ADC in", pen=pg.mkPen("#b45309", width=1.8))
        for curve in (self.wave_dac_curve, self.wave_adc_curve):
            curve.setClipToView(True)
            curve.setDownsampling(auto=True, method="peak")

        self.sweep_plot.addLegend(offset=(12, 12))
        self.sweep_monitor_curve = self.sweep_plot.plot(
            name="Monitor peak",
            pen=pg.mkPen("#7a0219", width=2.0),
            symbol="o",
            symbolSize=5,
            symbolBrush="#7a0219",
            symbolPen=None,
        )
        self.sweep_ramp_curve = self.sweep_plot.plot(
            name="Ramp peak",
            pen=pg.mkPen("#c9a227", width=1.5, style=QtCore.Qt.PenStyle.DashLine),
            symbol=None,
        )
        self.sweep_monitor_curve.setClipToView(True)
        self.sweep_monitor_curve.setDownsampling(auto=True, method="peak")
        self.sweep_ramp_curve.setClipToView(True)
        self.sweep_ramp_curve.setDownsampling(auto=True, method="peak")

        self.wave_plot.scene().sigMouseClicked.connect(lambda event: self._on_plot_mouse_clicked(self.wave_plot, event))
        self.sweep_plot.scene().sigMouseClicked.connect(lambda event: self._on_plot_mouse_clicked(self.sweep_plot, event))

        self.save_plot_btn.clicked.connect(self._save_plot_image)
        self.save_csv_btn.clicked.connect(self._save_plot_csv)
        self.clear_plot_btn.clicked.connect(self._clear_plots)
        self._reset_plot_view()
        return frame

    def _create_plot_widget(self, bottom_label: str, left_label: str, bottom_units: str, left_units: str) -> pg.PlotWidget:
        plot = pg.PlotWidget()
        plot.setBackground("#fffaf3")
        plot.showGrid(x=True, y=True, alpha=0.18)
        plot.setMouseEnabled(x=True, y=True)
        plot.getViewBox().setMouseMode(pg.ViewBox.RectMode)
        plot.getPlotItem().setMenuEnabled(False)
        plot.getPlotItem().setClipToView(True)
        plot.getPlotItem().setDownsampling(mode="peak")
        axis_pen = pg.mkPen("#7b6a63", width=1)
        for axis_name in ("bottom", "left"):
            axis = plot.getAxis(axis_name)
            axis.setPen(axis_pen)
            axis.setTextPen(axis_pen)
            axis.enableAutoSIPrefix(False)
        label_style = {"color": "#5a4741", "font-size": "11px"}
        plot.setLabel("bottom", bottom_label, units=bottom_units, **label_style)
        plot.setLabel("left", left_label, units=left_units, **label_style)
        return plot

    # ------------------------------------------------------------------
    # Data and status helpers
    # ------------------------------------------------------------------
    def _is_axial(self) -> bool:
        return self.axial_radio.isChecked()

    def _active_coil_name(self) -> str:
        return "axial" if self._is_axial() else "transverse"

    def _runtime_limits(self) -> AdwinCoilLimits:
        return AdwinCoilLimits(
            axial_ramp_max=self._runtime_coil_config.axial_max_ramp,
            axial_monitor_max=self._runtime_coil_config.axial_max_monitor,
            trans_ramp_max=self._runtime_coil_config.trans_max_ramp,
            trans_monitor_max=self._runtime_coil_config.trans_max_monitor,
        )

    def _active_runtime_freq(self) -> float:
        return (
            self._runtime_coil_config.axial_res_freq
            if self._is_axial()
            else self._runtime_coil_config.trans_res_freq
        )

    def _active_saved_freq(self) -> float:
        return self._coil_config.axial_res_freq if self._is_axial() else self._coil_config.trans_res_freq

    def _append(self, text: str) -> None:
        stamp = QtCore.QDateTime.currentDateTime().toString("HH:mm:ss")
        self.console.appendPlainText(f"[{stamp}] {text}")
        scrollbar = self.console.verticalScrollBar()
        scrollbar.setValue(scrollbar.maximum())
        self.status.setText(text)

    def _set_backend_status(self, text: str, color: str) -> None:
        self.backend_status.setText(text)
        self.backend_status.setStyleSheet(
            f"background: rgba(255,255,255,0.82); border: 1px solid rgba(122,2,25,0.16); "
            f"border-radius: 16px; padding: 8px 10px; font-weight: 650; color: {color};"
        )
        if hasattr(self, "quick_backend_status"):
            self.quick_backend_status.setText(text)
            self.quick_backend_status.setStyleSheet(
                f"background: rgba(255,255,255,0.82); border: 1px solid rgba(122,2,25,0.16); "
                f"border-radius: 16px; padding: 8px 10px; font-weight: 650; color: {color};"
            )

    def _set_comm_summary(self, text: str, color: str) -> None:
        self.comm_summary.setText(text)
        self.comm_summary.setStyleSheet(
            f"background: rgba(255,255,255,0.82); border: 1px solid rgba(122,2,25,0.16); "
            f"border-radius: 16px; padding: 8px 10px; font-weight: 650; color: {color};"
        )

    def _set_relay_status(self, text: str, color: str) -> None:
        self.relay_status.setText(text)
        self.relay_status.setStyleSheet(
            f"background: rgba(255,255,255,0.82); border: 1px solid rgba(122,2,25,0.16); "
            f"border-radius: 16px; padding: 8px 10px; font-weight: 650; color: {color};"
        )
        if hasattr(self, "quick_relay_status"):
            self.quick_relay_status.setText(text)
            self.quick_relay_status.setStyleSheet(
                f"background: rgba(255,255,255,0.82); border: 1px solid rgba(122,2,25,0.16); "
                f"border-radius: 16px; padding: 8px 10px; font-weight: 650; color: {color};"
            )

    def _set_busy(self, busy: bool) -> None:
        self.auto_tune_btn.setEnabled(True)
        self.capture_btn.setEnabled(True)
        self.connect_btn.setEnabled(not busy)
        self.check_comm_btn.setEnabled(not busy)
        self.quick_connect_btn.setEnabled(not busy)
        self.quick_check_btn.setEnabled(not busy)
        self.apply_relay_btn.setEnabled(not busy)
        self.relays_off_btn.setEnabled(not busy)
        self.save_backend_btn.setEnabled(not busy)
        if busy and self._task_kind == "sweep":
            self.auto_tune_btn.setText("Stop Auto-Tune")
            self.capture_btn.setEnabled(False)
        elif busy and self._task_kind == "capture":
            self.capture_btn.setText("Stop Capture")
            self.auto_tune_btn.setEnabled(False)
        else:
            self.auto_tune_btn.setText("Start Auto-Tune")
            self.capture_btn.setText("Run Clip / Firing Capture")

    def _load_into_widgets(self) -> None:
        if self._is_axial():
            self.old_freq.setValue(self._coil_config.axial_res_freq)
            self.new_freq.setValue(self._runtime_coil_config.axial_res_freq)
            self.old_ramp.setValue(self._coil_config.axial_max_ramp)
            self.new_ramp.setValue(self._runtime_coil_config.axial_max_ramp)
            self.old_monitor.setValue(self._coil_config.axial_max_monitor)
            self.new_monitor.setValue(self._runtime_coil_config.axial_max_monitor)
        else:
            self.old_freq.setValue(self._coil_config.trans_res_freq)
            self.new_freq.setValue(self._runtime_coil_config.trans_res_freq)
            self.old_ramp.setValue(self._coil_config.trans_max_ramp)
            self.new_ramp.setValue(self._runtime_coil_config.trans_max_ramp)
            self.old_monitor.setValue(self._coil_config.trans_max_monitor)
            self.new_monitor.setValue(self._runtime_coil_config.trans_max_monitor)

        self.low_freq.setValue(self._autotune_config.low_freq)
        self.high_freq.setValue(self._autotune_config.high_freq)
        self.step_freq.setValue(self._autotune_config.step_freq)
        self.hold_ms.setValue(self._autotune_config.hold_ms)
        self.io_rate.setValue(self._autotune_config.io_rate_hz)

        self.clip_freq.setValue(self._active_runtime_freq())
        self.clip_amp.setValue(self._clip_capture_config.amplitude_v)
        self.clip_duration_ms.setValue(self._clip_capture_config.duration_ms)
        self.clip_io_rate.setValue(self._clip_capture_config.io_rate_hz)
        self.clip_ramp_up.setValue(self._clip_capture_config.ramp_up_slope_vps)
        self.clip_ramp_down.setValue(self._clip_capture_config.ramp_down_slope_vps)
        self.clip_down_periods.setValue(self._clip_capture_config.ramp_down_periods)

        self.board_num.setValue(self._backend_config.board_num)
        self.bin_folder.setText(self._backend_config.bin_folder)
        self.boot_file.setText(self._backend_config.boot_file)
        self.process_file.setText(self._backend_config.process_file or _default_process_file())
        self.ramp_chan.setValue(self._backend_config.ramp_dac_chan)
        self.monitor_chan.setValue(self._backend_config.monitor_adc_chan)
        self.axial_relay_bit.setValue(self._backend_config.axial_relay_bit)
        self.trans_relay_bit.setValue(self._backend_config.trans_relay_bit)

    def _build_backend_from_widgets(self) -> BackendConfig:
        process_path = self.process_file.text().strip() or _default_process_file()
        return BackendConfig(
            board_num=self.board_num.value(),
            bin_folder=self.bin_folder.text().strip(),
            boot_file=self.boot_file.text().strip() or "ADwin9.btl",
            process_file=process_path,
            ramp_dac_chan=self.ramp_chan.value(),
            monitor_adc_chan=self.monitor_chan.value(),
            axial_relay_bit=self.axial_relay_bit.value(),
            trans_relay_bit=self.trans_relay_bit.value(),
        )

    def _set_coil_locked(self, locked: bool, announce: bool = True) -> None:
        self.chk_lock_coils.blockSignals(True)
        self.chk_lock_coils.setChecked(locked)
        self.chk_lock_coils.blockSignals(False)
        self.axial_radio.setEnabled(not locked or self.axial_radio.isChecked())
        self.trans_radio.setEnabled(not locked or self.trans_radio.isChecked())
        if announce:
            self._append("Coil selection locked." if locked else "Coil selection unlocked.")

    # ------------------------------------------------------------------
    # Backend actions
    # ------------------------------------------------------------------
    def _browse_bin_folder(self) -> None:
        start = self.bin_folder.text().strip() or str(Path.home())
        path = QtWidgets.QFileDialog.getExistingDirectory(self, "Select ADwin Bin/BTL Folder", start)
        if path:
            self.bin_folder.setText(path)

    def _browse_boot_file(self) -> None:
        start = self.bin_folder.text().strip() or str(Path.home())
        path, _ = QtWidgets.QFileDialog.getOpenFileName(
            self,
            "Select ADwin Boot File",
            start,
            "ADwin boot files (*.btl);;All files (*)",
        )
        if not path:
            return
        file_path = Path(path)
        self.bin_folder.setText(str(file_path.parent))
        self.boot_file.setText(file_path.name)

    def _browse_process_file(self) -> None:
        start = self.process_file.text().strip() or str(_repo_root())
        path, _ = QtWidgets.QFileDialog.getOpenFileName(
            self,
            "Select ADwin Process File",
            start,
            "ADwin process files (*.T9* *.T91 *.abp);;All files (*)",
        )
        if path:
            self.process_file.setText(path)

    def _connect_backend(self) -> None:
        self._backend_config = self._build_backend_from_widgets()
        if not self._backend_config.process_file:
            QtWidgets.QMessageBox.warning(
                self,
                "Missing Process File",
                "The ADwin process file is not set. Point it at VB6/ADwin/sineout.T91 before connecting.",
            )
            return
        try:
            controller = AdwinAFController(
                board=AdwinBoardConfig(**asdict(self._backend_config)),
                limits=self._runtime_limits(),
            )
            try:
                version = controller.test_version()
            except Exception:
                version = 0
            relay_word = controller.set_af_relays(self._active_coil_name(), one_chan_on=True)
            try:
                version_after = controller.test_version()
            except Exception:
                version_after = 0
            self._last_version = max(version, version_after)
            self._ctrl = controller
            self._connected = True
            if self._last_version:
                self._set_backend_status(f"Connected (v{self._last_version})", "#0f766e")
            else:
                self._set_backend_status("Connected (I/O ready)", "#0f766e")
            self._append(
                f"ADwin ready on board {self._backend_config.board_num}; active coil relay word 0x{relay_word:02X}."
            )
            self._update_comm_snapshot(log_result=False, relay_word=relay_word)
        except Exception as exc:
            self._ctrl = None
            self._connected = False
            self._last_version = 0
            self._set_backend_status("Connection failed", "#991b1b")
            self._set_comm_summary("Communication not established.", "#991b1b")
            self._set_relay_status("Relay state unavailable.", "#991b1b")
            self._append(f"[ERROR] ADwin connect/boot failed: {exc}")

    def _check_communication(self) -> None:
        if self._ctrl is None:
            self._connect_backend()
            if self._ctrl is None:
                return
        self._update_comm_snapshot(log_result=True)

    def _update_comm_snapshot(self, log_result: bool, relay_word: int | None = None) -> None:
        if self._ctrl is None:
            self._set_comm_summary("Communication not established.", "#991b1b")
            self._set_relay_status("Relay state unavailable.", "#991b1b")
            return

        try:
            try:
                version = self._ctrl.test_version()
            except Exception:
                version = 0
            digout_word = self._ctrl.get_digout() if relay_word is None else relay_word
            adc_v = self._ctrl.get_adc(self.monitor_chan.value())
            self._last_version = max(self._last_version, version)
            version_text = f"v{version}" if version else "I/O ready"
            self._set_comm_summary(
                f"Board {self.board_num.value()} responding: {version_text}, ADC ch{self.monitor_chan.value()} = {adc_v:+.4f} V",
                "#0f766e",
            )
            self._set_relay_status(
                f"Requested coil: {self._active_coil_name().title()} | Digout = 0x{int(digout_word) & 0x3F:02X}",
                "#1a237e",
            )
            if log_result:
                self._append(
                    f"Communication established: board={self.board_num.value()} {version_text}, "
                    f"digout=0x{int(digout_word) & 0x3F:02X}, adc={adc_v:+.4f} V"
                )
        except Exception as exc:
            self._set_comm_summary("Communication check failed.", "#991b1b")
            self._set_relay_status("Relay state unavailable.", "#991b1b")
            if log_result:
                self._append(f"[ERROR] Communication check failed: {exc}")

    def _save_backend(self) -> None:
        cfg = self._build_backend_from_widgets()
        self._backend_config = cfg
        save_dataclass(BACKEND_CONFIG_PATH, cfg)
        self._append("Saved ADwin backend settings.")

    def _apply_relays(self) -> None:
        if self._ctrl is None:
            self._connect_backend()
            if self._ctrl is None:
                return
        try:
            self._ctrl.board = AdwinBoardConfig(**asdict(self._build_backend_from_widgets()))
            relay_word = self._ctrl.set_af_relays(self._active_coil_name(), one_chan_on=True)
            self._update_comm_snapshot(log_result=False, relay_word=relay_word)
            self._append(f"AF relays set for {self._active_coil_name()} coil (word 0x{relay_word:02X}).")
        except Exception as exc:
            self._append(f"[ERROR] Relay update failed: {exc}")

    def _all_relays_off(self) -> None:
        if self._ctrl is None:
            self._connect_backend()
            if self._ctrl is None:
                return
        try:
            self._ctrl.set_af_relays("off", one_chan_on=True)
            self._update_comm_snapshot(log_result=False, relay_word=0)
            self._append("All AF relays turned off.")
        except Exception as exc:
            self._append(f"[ERROR] Failed to turn all relays off: {exc}")

    # ------------------------------------------------------------------
    # Coil and value actions
    # ------------------------------------------------------------------
    def _on_lock_toggled(self, checked: bool) -> None:
        self._set_coil_locked(checked)

    def _on_coil_changed(self, *_args) -> None:
        if self.chk_lock_coils.isChecked():
            return
        self._load_into_widgets()
        self._apply_relays()

    def _apply_freq(self) -> None:
        freq = self.new_freq.value()
        if freq <= 0:
            QtWidgets.QMessageBox.warning(self, "Invalid Frequency", "Resonance frequency must be positive.")
            return
        if self._is_axial():
            self._runtime_coil_config.axial_res_freq = freq
        else:
            self._runtime_coil_config.trans_res_freq = freq
        self.clip_freq.setValue(freq)
        self._append(f"Applied {self._active_coil_name()} resonance frequency: {freq:.4f} Hz")

    def _apply_volts(self) -> None:
        ramp = self.new_ramp.value()
        monitor = self.new_monitor.value()
        if ramp <= 0 or monitor <= 0:
            QtWidgets.QMessageBox.warning(self, "Invalid Voltages", "Max ramp and monitor voltages must be positive.")
            return
        if self._is_axial():
            self._runtime_coil_config.axial_max_ramp = ramp
            self._runtime_coil_config.axial_max_monitor = monitor
        else:
            self._runtime_coil_config.trans_max_ramp = ramp
            self._runtime_coil_config.trans_max_monitor = monitor
        self._append(
            f"Applied {self._active_coil_name()} runtime max voltages: ramp={ramp:.4f} V, monitor={monitor:.4f} V"
        )

    def _save_freq(self) -> None:
        self._apply_freq()
        freq = self.new_freq.value()
        if self._is_axial():
            self._coil_config.axial_res_freq = freq
            self.old_freq.setValue(freq)
        else:
            self._coil_config.trans_res_freq = freq
            self.old_freq.setValue(freq)
        save_dataclass(COIL_CONFIG_PATH, self._coil_config)
        self._append("Frequency saved to local config.")

    def _save_volts(self) -> None:
        self._apply_volts()
        ramp = self.new_ramp.value()
        monitor = self.new_monitor.value()
        if self._is_axial():
            self._coil_config.axial_max_ramp = ramp
            self._coil_config.axial_max_monitor = monitor
            self.old_ramp.setValue(ramp)
            self.old_monitor.setValue(monitor)
        else:
            self._coil_config.trans_max_ramp = ramp
            self._coil_config.trans_max_monitor = monitor
            self.old_ramp.setValue(ramp)
            self.old_monitor.setValue(monitor)
        save_dataclass(COIL_CONFIG_PATH, self._coil_config)
        self._append("Max voltages saved to local config.")

    # ------------------------------------------------------------------
    # Task launchers
    # ------------------------------------------------------------------
    def _toggle_autotune(self) -> None:
        if self._task_kind == "sweep" and isinstance(self._worker, AutoTuneWorker):
            self._worker.abort()
            self._append("Stopping auto-tune…")
            return
        if self._task_kind is not None:
            self._append("Another ADwin operation is already running.")
            return

        low = self.low_freq.value()
        high = self.high_freq.value()
        step = self.step_freq.value()
        if low <= 0 or high <= 0 or step <= 0 or high < low:
            QtWidgets.QMessageBox.warning(
                self,
                "Invalid Sweep",
                "Sweep values must satisfy 0 < low <= high and step > 0.",
            )
            return

        self._apply_freq()
        self._apply_volts()
        self._save_backend()
        self._autotune_config.low_freq = low
        self._autotune_config.high_freq = high
        self._autotune_config.step_freq = step
        self._autotune_config.hold_ms = self.hold_ms.value()
        self._autotune_config.io_rate_hz = self.io_rate.value()
        save_dataclass(AUTOTUNE_CONFIG_PATH, self._autotune_config)

        worker = AutoTuneWorker(
            backend=self._backend_config,
            limits=self._runtime_limits(),
            low_freq=low,
            high_freq=high,
            step_freq=step,
            hold_ms=self.hold_ms.value(),
            io_rate_hz=self.io_rate.value(),
            ramp_peak_v=self.new_ramp.value(),
            monitor_peak_v=self.new_monitor.value(),
            coil=self._active_coil_name(),
        )
        thread = QtCore.QThread(self)
        worker.moveToThread(thread)
        thread.started.connect(worker.run)
        worker.progress.connect(self._append)
        worker.finished.connect(self._on_autotune_finished)
        worker.failed.connect(self._on_task_failed)
        worker.finished.connect(thread.quit)
        worker.failed.connect(thread.quit)
        thread.finished.connect(self._cleanup_worker)

        self._task_kind = "sweep"
        self._worker = worker
        self._worker_thread = thread
        self._queued_capture_freq = None
        self._set_busy(True)
        self._set_coil_locked(True)
        self._append("Auto-tune started.")
        thread.start()

    def _toggle_capture(self) -> None:
        if self._task_kind == "capture" and isinstance(self._worker, DenseCaptureWorker):
            self._worker.stop()
            self._append("Stopping firing capture…")
            return
        if self._task_kind is not None:
            self._append("Another ADwin operation is already running.")
            return
        self._start_capture(
            freq_hz=self.clip_freq.value(),
            label="Clip / firing capture",
        )

    def _start_capture(self, freq_hz: float, label: str) -> None:
        if freq_hz <= 0 or self.clip_amp.value() <= 0:
            QtWidgets.QMessageBox.warning(self, "Invalid Capture", "Capture frequency and amplitude must be positive.")
            return

        self._apply_volts()
        self._save_backend()
        self._clip_capture_config.sine_freq_hz = freq_hz
        self._clip_capture_config.amplitude_v = self.clip_amp.value()
        self._clip_capture_config.duration_ms = self.clip_duration_ms.value()
        self._clip_capture_config.io_rate_hz = self.clip_io_rate.value()
        self._clip_capture_config.ramp_up_slope_vps = self.clip_ramp_up.value()
        self._clip_capture_config.ramp_down_slope_vps = self.clip_ramp_down.value()
        self._clip_capture_config.ramp_down_periods = self.clip_down_periods.value()
        save_dataclass(CLIP_CAPTURE_PATH, self._clip_capture_config)

        worker = DenseCaptureWorker(
            backend=self._backend_config,
            limits=self._runtime_limits(),
            coil=self._active_coil_name(),
            freq_hz=freq_hz,
            amplitude_v=self.clip_amp.value(),
            duration_ms=self.clip_duration_ms.value(),
            io_rate_hz=self.clip_io_rate.value(),
            ramp_up_slope_vps=self.clip_ramp_up.value(),
            ramp_down_slope_vps=self.clip_ramp_down.value(),
            ramp_down_periods=self.clip_down_periods.value(),
            label=label,
        )
        thread = QtCore.QThread(self)
        worker.moveToThread(thread)
        thread.started.connect(worker.run)
        worker.progress.connect(self._append)
        worker.capture_ready.connect(self._on_capture_ready)
        worker.finished.connect(self._on_capture_finished)
        worker.failed.connect(self._on_task_failed)
        worker.finished.connect(thread.quit)
        worker.failed.connect(thread.quit)
        thread.finished.connect(self._cleanup_worker)

        self._task_kind = "capture"
        self._worker = worker
        self._worker_thread = thread
        self._set_busy(True)
        self._set_coil_locked(True)
        self._append(
            f"{label} started at {freq_hz:.4f} Hz; {self.clip_io_rate.value():.0f} Hz IO rate = "
            f"{self.clip_io_rate.value() / max(freq_hz, 0.1):.1f} samples/cycle."
        )
        thread.start()

    # ------------------------------------------------------------------
    # Task callbacks
    # ------------------------------------------------------------------
    @QtCore.Slot(object)
    def _on_autotune_finished(self, result: object) -> None:
        tune_result = result if isinstance(result, AutoTuneResult) else None
        if tune_result is None:
            return
        self._sweep_freqs = [point.freq_hz for point in tune_result.points]
        self._sweep_monitor = [point.monitor_peak_v for point in tune_result.points]
        self._sweep_ramp = [point.ramp_peak_v for point in tune_result.points]
        self._refresh_sweep_plot()
        if tune_result.best_freq_hz > 0:
            self.new_freq.setValue(tune_result.best_freq_hz)
            self._runtime_coil_config.axial_res_freq = (
                tune_result.best_freq_hz if self._is_axial() else self._runtime_coil_config.axial_res_freq
            )
            self._runtime_coil_config.trans_res_freq = (
                tune_result.best_freq_hz if not self._is_axial() else self._runtime_coil_config.trans_res_freq
            )
            self.clip_freq.setValue(tune_result.best_freq_hz)
            self._queued_capture_freq = tune_result.best_freq_hz
        self._last_export_kind = "sweep"
        self.plot_summary.setText(
            f"Auto-tune sweep completed. Best frequency = {tune_result.best_freq_hz:.4f} Hz, "
            f"best monitor peak = {tune_result.best_monitor_peak_v:.5f} V. A board-timed waveform preview will run next."
        )
        self._append(
            f"Auto-tune completed. Best freq={tune_result.best_freq_hz:.5f} Hz, "
            f"monitor_peak={tune_result.best_monitor_peak_v:.6f} V"
        )

    @QtCore.Slot(object)
    def _on_capture_ready(self, payload: object) -> None:
        envelope = payload if isinstance(payload, CaptureEnvelope) else None
        if envelope is None:
            return
        capture = envelope.capture
        self._wave_time = list(capture.time_s)
        self._wave_dac = list(capture.dac_v)
        self._wave_adc = list(capture.adc_v)
        self._refresh_wave_plot()
        observed_ramp = max((abs(value) for value in capture.dac_v), default=0.0)
        observed_monitor = max((abs(value) for value in capture.adc_v), default=0.0)
        self.new_ramp.setValue(observed_ramp)
        self.new_monitor.setValue(observed_monitor)
        self._last_export_kind = "waveform"
        self.plot_summary.setText(
            f"{envelope.label}: {len(capture.time_s)} points, dt={capture.timestep_s * 1000.0:.3f} ms, "
            f"{capture.points_per_period:.1f} samples/cycle. Observed ramp={observed_ramp:.4f} V, "
            f"monitor={observed_monitor:.4f} V loaded into the runtime fields for review."
        )
        self._append(
            f"{envelope.label} captured {len(capture.time_s)} points; dt={capture.timestep_s * 1000.0:.3f} ms, "
            f"samples/cycle={capture.points_per_period:.1f}."
        )

    @QtCore.Slot()
    def _on_capture_finished(self) -> None:
        self._append("ADwin firing capture finished.")

    @QtCore.Slot(str)
    def _on_task_failed(self, message: str) -> None:
        QtWidgets.QMessageBox.critical(self, "ADwin Operation Error", message)
        self._append(f"[ERROR] {message}")
        self._queued_capture_freq = None

    @QtCore.Slot()
    def _cleanup_worker(self) -> None:
        finished_kind = self._task_kind
        if self._worker_thread is not None:
            self._worker_thread.deleteLater()
        if self._worker is not None:
            self._worker.deleteLater()
        self._worker_thread = None
        self._worker = None
        self._task_kind = None
        self._set_busy(False)
        self._set_coil_locked(False, announce=False)

        queued = self._queued_capture_freq
        self._queued_capture_freq = None
        if finished_kind == "sweep" and queued is not None:
            QtCore.QTimer.singleShot(0, lambda: self._start_capture(queued, "Best-frequency waveform preview"))

    # ------------------------------------------------------------------
    # Plot helpers
    # ------------------------------------------------------------------
    def _refresh_wave_plot(self) -> None:
        time_ms = [value * 1000.0 for value in self._wave_time]
        self.wave_dac_curve.setData(time_ms, self._wave_dac)
        self.wave_adc_curve.setData(time_ms, self._wave_adc)
        self._reset_wave_plot_view()

    def _refresh_sweep_plot(self) -> None:
        self.sweep_monitor_curve.setData(self._sweep_freqs, self._sweep_monitor)
        self.sweep_ramp_curve.setData(self._sweep_freqs, self._sweep_ramp)
        self._reset_sweep_plot_view()

    def _clear_plots(self) -> None:
        self._wave_time.clear()
        self._wave_dac.clear()
        self._wave_adc.clear()
        self._sweep_freqs.clear()
        self._sweep_monitor.clear()
        self._sweep_ramp.clear()
        self.wave_dac_curve.setData([], [])
        self.wave_adc_curve.setData([], [])
        self.sweep_monitor_curve.setData([], [])
        self.sweep_ramp_curve.setData([], [])
        self.plot_summary.setText("Plots cleared.")
        self._reset_plot_view()

    def _reset_plot_view(self) -> None:
        self._reset_wave_plot_view()
        self._reset_sweep_plot_view()

    def _reset_wave_plot_view(self) -> None:
        duration_ms = max(float(self.clip_duration_ms.value()), 50.0)
        amplitude = max(self.clip_amp.value(), self.new_monitor.value(), self.new_ramp.value(), 0.5)
        y_pad = max(0.25, amplitude * 1.25)
        self.wave_plot.setXRange(0.0, duration_ms, padding=0.01)
        self.wave_plot.setYRange(-y_pad, y_pad, padding=0.04)

    def _reset_sweep_plot_view(self) -> None:
        low = min(self.low_freq.value(), self.high_freq.value())
        high = max(self.low_freq.value(), self.high_freq.value())
        peak = max(self._sweep_monitor + self._sweep_ramp + [self.new_monitor.value(), 0.5])
        self.sweep_plot.setXRange(low, high if high > low else low + 1.0, padding=0.03)
        self.sweep_plot.setYRange(0.0, peak * 1.2, padding=0.05)

    def _on_plot_mouse_clicked(self, plot: pg.PlotWidget, event) -> None:
        if not event.double():
            return
        if plot.getViewBox().sceneBoundingRect().contains(event.scenePos()):
            plot.enableAutoRange(axis=pg.ViewBox.XYAxes, enable=True)
            plot.getViewBox().autoRange()
            plot.enableAutoRange(axis=pg.ViewBox.XYAxes, enable=False)

    def _capture_basename(self, stem: str) -> str:
        stamp = QtCore.QDateTime.currentDateTime().toString("yyyyMMdd-HHmmss")
        return f"af-tuner-{stem}-{stamp}"

    def _save_plot_image(self) -> None:
        target = self.wave_plot if self._wave_time else self.sweep_plot
        if not self._wave_time and not self._sweep_freqs:
            self._append("[WARNING] No plot data is loaded yet.")
            return
        default_path = str(Path.home() / f"{self._capture_basename(self._last_export_kind)}.png")
        path, _ = QtWidgets.QFileDialog.getSaveFileName(
            self,
            "Save AF tuner plot",
            default_path,
            "PNG image (*.png);;JPEG image (*.jpg *.jpeg);;All files (*)",
        )
        if not path:
            return
        pixmap = target.grab()
        if not pixmap.save(path):
            self._append(f"[ERROR] Failed to save plot image: {path}")
            return
        self._append(f"Saved plot image: {path}")

    def _save_plot_csv(self) -> None:
        if self._wave_time:
            default_path = str(Path.home() / f"{self._capture_basename('waveform')}.csv")
            path, _ = QtWidgets.QFileDialog.getSaveFileName(
                self,
                "Save waveform CSV",
                default_path,
                "CSV files (*.csv);;All files (*)",
            )
            if not path:
                return
            with open(path, "w", newline="", encoding="utf-8") as handle:
                writer = csv.writer(handle)
                writer.writerow(["time_s", "dac_v", "adc_v"])
                writer.writerows(zip(self._wave_time, self._wave_dac, self._wave_adc))
            self._append(f"Saved waveform CSV: {path}")
            return

        if self._sweep_freqs:
            default_path = str(Path.home() / f"{self._capture_basename('sweep')}.csv")
            path, _ = QtWidgets.QFileDialog.getSaveFileName(
                self,
                "Save sweep CSV",
                default_path,
                "CSV files (*.csv);;All files (*)",
            )
            if not path:
                return
            with open(path, "w", newline="", encoding="utf-8") as handle:
                writer = csv.writer(handle)
                writer.writerow(["freq_hz", "monitor_peak_v", "ramp_peak_v"])
                writer.writerows(zip(self._sweep_freqs, self._sweep_monitor, self._sweep_ramp))
            self._append(f"Saved sweep CSV: {path}")
            return

        self._append("[WARNING] No plot data is loaded yet.")


def main() -> int:
    pg.setConfigOptions(antialias=False, foreground="#4d3a39")
    app = QtWidgets.QApplication(sys.argv)
    apply_liquid_glass_theme(app)
    set_app_icon(app, "af_tuner_icon.ico", _assets_dir())
    window = MainWindow()
    window.show()
    return app.exec()
