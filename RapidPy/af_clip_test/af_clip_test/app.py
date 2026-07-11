from __future__ import annotations

import csv
import json
import math
import sys
from dataclasses import asdict, dataclass
from pathlib import Path

import numpy as np
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
)
from rapidpy_common.ui import apply_card_shadow, apply_liquid_glass_theme, set_app_icon  # noqa: E402


@dataclass(slots=True)
class CoilTuningConfig:
    axial_res_freq: float = 877.0
    axial_max_ramp: float = 0.5
    axial_max_monitor: float = 0.5
    trans_res_freq: float = 324.0
    trans_max_ramp: float = 0.5
    trans_max_monitor: float = 0.5


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
class ClipTestConfig:
    min_amp_v: float = 0.0
    max_amp_v: float = 10.0
    sine_freq_hz: float = 877.0
    scan_points: int = 21
    io_rate_hz: float = 25000.0
    duration_ms: int = 220
    ramp_up_slope_vps: float = 200.0
    ramp_down_slope_vps: float = 200.0
    ramp_down_periods: int = 2


@dataclass(slots=True)
class ClipPoint:
    ramp_voltage_v: float
    monitor_amplitude_v: float
    residual_rms_v: float


@dataclass(slots=True)
class ClipScanResult:
    coil: str
    sine_freq_hz: float
    avg_points: list[ClipPoint]
    up_points: list[ClipPoint]
    down_points: list[ClipPoint]
    suggested_point: ClipPoint | None
    preview_capture: AdwinDenseCaptureResult | None


COIL_CONFIG_PATH = Path.home() / ".rapidpy_af_tuner.json"
BACKEND_CONFIG_PATH = Path.home() / ".rapidpy_af_backend.json"
CLIP_TEST_CONFIG_PATH = Path.home() / ".rapidpy_af_clip_test.json"


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
        payload = json.loads(path.read_text(encoding="utf-8"))
        values = asdict(cls())
        for key in values:
            if key in payload:
                values[key] = payload[key]
        return cls(**values)
    except (OSError, TypeError, json.JSONDecodeError):
        return cls()


def save_dataclass(path: Path, obj) -> None:
    try:
        path.write_text(json.dumps(asdict(obj), indent=2, sort_keys=True), encoding="utf-8")
    except OSError:
        return


def _linspace(start: float, stop: float, count: int) -> list[float]:
    if count <= 1:
        return [float(stop)]
    step = (float(stop) - float(start)) / float(count - 1)
    return [float(start) + step * index for index in range(count)]


def _fit_sine_metrics(samples: list[float], timestep_s: float, freq_hz: float) -> tuple[float, float]:
    if len(samples) < 12 or timestep_s <= 0 or freq_hz <= 0:
        return float("nan"), float("nan")
    values = np.asarray(samples, dtype=float)
    if not np.isfinite(values).all():
        return float("nan"), float("nan")
    time_s = np.arange(values.size, dtype=float) * float(timestep_s)
    omega = 2.0 * math.pi * float(freq_hz)
    design = np.column_stack((np.ones_like(time_s), np.sin(omega * time_s), np.cos(omega * time_s)))
    try:
        coeffs, *_ = np.linalg.lstsq(design, values, rcond=None)
    except np.linalg.LinAlgError:
        return float("nan"), float("nan")
    fit = design @ coeffs
    amplitude = float(np.hypot(coeffs[1], coeffs[2]))
    residual_rms = float(np.sqrt(np.mean(np.square(values - fit))))
    return amplitude, residual_rms


def _combine_passes(amplitudes: list[float], up_points: list[ClipPoint], down_points: list[ClipPoint]) -> list[ClipPoint]:
    up_map = {round(point.ramp_voltage_v, 6): point for point in up_points}
    down_map = {round(point.ramp_voltage_v, 6): point for point in down_points}
    combined: list[ClipPoint] = []
    for amplitude in amplitudes:
        key = round(amplitude, 6)
        monitors: list[float] = []
        residuals: list[float] = []
        for point in (up_map.get(key), down_map.get(key)):
            if point is None:
                continue
            if math.isfinite(point.monitor_amplitude_v):
                monitors.append(point.monitor_amplitude_v)
            if math.isfinite(point.residual_rms_v):
                residuals.append(point.residual_rms_v)
        combined.append(
            ClipPoint(
                ramp_voltage_v=amplitude,
                monitor_amplitude_v=float(np.mean(monitors)) if monitors else float("nan"),
                residual_rms_v=float(np.mean(residuals)) if residuals else float("nan"),
            )
        )
    return combined


def _suggest_limit(points: list[ClipPoint]) -> ClipPoint | None:
    valid = [
        point
        for point in points
        if math.isfinite(point.monitor_amplitude_v) and math.isfinite(point.residual_rms_v)
    ]
    if not valid:
        return None

    baseline_count = max(3, min(6, len(valid)))
    baseline = float(np.median([point.residual_rms_v for point in valid[:baseline_count]]))
    threshold = max(baseline * 3.0, baseline + 0.01)
    safe_points = [point for point in valid if point.residual_rms_v <= threshold]
    if safe_points:
        return safe_points[-1]
    return min(valid, key=lambda point: point.residual_rms_v)


class AutoClipWorker(QtCore.QObject):
    progress = QtCore.Signal(str)
    result_ready = QtCore.Signal(object)
    finished = QtCore.Signal()
    failed = QtCore.Signal(str)

    def __init__(
        self,
        backend: BackendConfig,
        limits: AdwinCoilLimits,
        coil: str,
        config: ClipTestConfig,
        parent: QtCore.QObject | None = None,
    ) -> None:
        super().__init__(parent)
        self._backend = _copy_config(backend)
        self._limits = limits
        self._coil = coil
        self._config = _copy_config(config)
        self._stop = False

    def stop(self) -> None:
        self._stop = True

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

            amplitudes = _linspace(self._config.min_amp_v, self._config.max_amp_v, self._config.scan_points)
            duration_s = max(
                float(self._config.duration_ms) / 1000.0,
                max(6.0 / max(self._config.sine_freq_hz, 0.1), 0.08),
            )

            up_points: list[ClipPoint] = []
            down_points: list[ClipPoint] = []
            preview_capture: AdwinDenseCaptureResult | None = None

            for pass_name, sweep in (("up", amplitudes), ("down", list(reversed(amplitudes)))):
                point_total = len(sweep)
                for index, amplitude in enumerate(sweep, start=1):
                    if self._stop:
                        raise AdwinError("Clipping test stopped by user.")
                    capture = controller.run_dense_loopback(
                        AdwinDenseCaptureRequest(
                            sine_freq_hz=self._config.sine_freq_hz,
                            amplitude_v=amplitude,
                            io_rate_hz=self._config.io_rate_hz,
                            duration_s=duration_s,
                            dac_chan=board.ramp_dac_chan,
                            adc_chan=board.monitor_adc_chan,
                            ramp_mode=3,
                            ramp_up_slope_vps=self._config.ramp_up_slope_vps,
                            ramp_down_slope_vps=self._config.ramp_down_slope_vps,
                            ramp_down_periods=max(1, self._config.ramp_down_periods),
                        ),
                        should_stop=lambda: self._stop,
                    )
                    monitor_amp_v, residual_rms_v = _fit_sine_metrics(
                        capture.adc_v,
                        capture.timestep_s,
                        self._config.sine_freq_hz,
                    )
                    point = ClipPoint(amplitude, monitor_amp_v, residual_rms_v)
                    if pass_name == "up":
                        up_points.append(point)
                        if index == point_total:
                            preview_capture = capture
                    else:
                        down_points.append(point)
                    self.progress.emit(
                        f"Clip {pass_name} {index}/{point_total}: ramp {amplitude:.3f} V, "
                        f"monitor {monitor_amp_v:.4f} V, residual {residual_rms_v:.5f} V RMS"
                    )

            avg_points = _combine_passes(amplitudes, up_points, down_points)
            suggested_point = _suggest_limit(avg_points)
            self.result_ready.emit(
                ClipScanResult(
                    coil=self._coil,
                    sine_freq_hz=self._config.sine_freq_hz,
                    avg_points=avg_points,
                    up_points=up_points,
                    down_points=down_points,
                    suggested_point=suggested_point,
                    preview_capture=preview_capture,
                )
            )
            self.finished.emit()
        except Exception as exc:
            self.failed.emit(str(exc))
            self.finished.emit()


class MainWindow(QtWidgets.QMainWindow):
    def __init__(self) -> None:
        super().__init__()
        self.setWindowTitle("RapidPy AF Clip Test")

        self._coil_config = load_dataclass(COIL_CONFIG_PATH, CoilTuningConfig)
        self._backend_config = load_dataclass(BACKEND_CONFIG_PATH, BackendConfig)
        if not self._backend_config.process_file:
            self._backend_config.process_file = _default_process_file()
        self._clip_test_config = load_dataclass(CLIP_TEST_CONFIG_PATH, ClipTestConfig)

        self._fit_ramp_by_coil = {
            "axial": self._coil_config.axial_max_ramp,
            "transverse": self._coil_config.trans_max_ramp,
        }
        self._fit_monitor_by_coil = {
            "axial": self._coil_config.axial_max_monitor,
            "transverse": self._coil_config.trans_max_monitor,
        }

        self._ctrl: AdwinAFController | None = None
        self._connected = False
        self._last_version = 0
        self._worker_thread: QtCore.QThread | None = None
        self._worker: AutoClipWorker | None = None
        self._last_result: ClipScanResult | None = None

        self._build_ui()
        self._load_into_widgets()
        self._set_coil_locked(False, announce=False)
        self._set_backend_status("Not connected", "#6c5a53")
        self._set_comm_summary("Communication check not run yet.", "#6c5a53")
        self._set_relay_status("Relay state unknown.", "#6c5a53")
        self._clear_plot_state()

        compact_font = QtGui.QFont(self.font())
        compact_size = compact_font.pointSizeF()
        if compact_size > 0:
            compact_font.setPointSizeF(max(9.0, compact_size - 0.4))
            self.setFont(compact_font)

        set_app_icon(self, "af_clip_test_icon.ico", _assets_dir())
        self.setMinimumSize(1500, 860)
        self.resize(1660, 940)

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

        clip_card = self._build_clip_card()
        backend_card = self._build_backend_card()
        console_card = self._build_console_card()
        for card in (clip_card, backend_card, console_card):
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
        splitter.setSizes([560, 1040])

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

    def _build_clip_card(self) -> QtWidgets.QFrame:
        frame = QtWidgets.QFrame()
        frame.setObjectName("card")
        layout = QtWidgets.QVBoxLayout(frame)
        layout.setContentsMargins(18, 18, 18, 18)
        layout.setSpacing(12)

        title = QtWidgets.QLabel("AF Clip Test")
        title.setObjectName("title")
        subtitle = QtWidgets.QLabel(
            "VB6-style clipping test: sweep ramp voltage up and down, fit the monitor sine response, and review residual RMS against monitor amplitude before saving new coil limits."
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
        self.quick_backend_status.setSizePolicy(QtWidgets.QSizePolicy.Policy.Ignored, QtWidgets.QSizePolicy.Policy.Preferred)

        self.quick_relay_status = QtWidgets.QLabel("Relay state unknown")
        self.quick_relay_status.setObjectName("valuePill")
        self.quick_relay_status.setAlignment(QtCore.Qt.AlignmentFlag.AlignCenter)
        self.quick_relay_status.setWordWrap(True)
        self.quick_relay_status.setMinimumWidth(0)
        self.quick_relay_status.setSizePolicy(QtWidgets.QSizePolicy.Policy.Ignored, QtWidgets.QSizePolicy.Policy.Preferred)

        self.quick_connect_btn = QtWidgets.QPushButton("Connect / Boot")
        self.quick_connect_btn.setObjectName("connectAction")
        self.quick_check_btn = QtWidgets.QPushButton("Check Comm")
        quick_comm_row.addWidget(self.quick_backend_status, 0, 0, 1, 2)
        quick_comm_row.addWidget(self.quick_relay_status, 1, 0, 1, 2)
        quick_comm_row.addWidget(self.quick_connect_btn, 2, 0)
        quick_comm_row.addWidget(self.quick_check_btn, 2, 1)
        layout.addLayout(quick_comm_row)

        limits_box = QtWidgets.QGroupBox("Active Coil Limits")
        limits_form = QtWidgets.QFormLayout(limits_box)
        limits_form.setSpacing(8)
        self.saved_res_freq = QtWidgets.QLabel("0.0000 Hz")
        self.saved_res_freq.setObjectName("valuePill")
        self.saved_max_ramp = QtWidgets.QLabel("0.0000 V")
        self.saved_max_ramp.setObjectName("valuePill")
        self.saved_max_monitor = QtWidgets.QLabel("0.0000 V")
        self.saved_max_monitor.setObjectName("valuePill")
        self.fit_max_ramp = QtWidgets.QDoubleSpinBox()
        self.fit_max_ramp.setRange(0.0, 10.0)
        self.fit_max_ramp.setDecimals(4)
        self.fit_max_monitor = QtWidgets.QDoubleSpinBox()
        self.fit_max_monitor.setRange(0.0, 10.0)
        self.fit_max_monitor.setDecimals(4)
        limits_form.addRow("Saved Resonance", self.saved_res_freq)
        limits_form.addRow("Saved Max Ramp", self.saved_max_ramp)
        limits_form.addRow("Saved Max Monitor", self.saved_max_monitor)
        limits_form.addRow("Suggested Max Ramp", self.fit_max_ramp)
        limits_form.addRow("Suggested Max Monitor", self.fit_max_monitor)
        layout.addWidget(limits_box)

        scan_box = QtWidgets.QGroupBox("Auto Clipping Test")
        scan_form = QtWidgets.QFormLayout(scan_box)
        scan_form.setSpacing(8)
        self.min_clip_amp = QtWidgets.QDoubleSpinBox()
        self.min_clip_amp.setRange(0.0, 10.0)
        self.min_clip_amp.setDecimals(4)
        self.max_clip_amp = QtWidgets.QDoubleSpinBox()
        self.max_clip_amp.setRange(0.0, 10.0)
        self.max_clip_amp.setDecimals(4)
        self.clip_freq = QtWidgets.QDoubleSpinBox()
        self.clip_freq.setRange(0.001, 5000.0)
        self.clip_freq.setDecimals(4)
        self.scan_points = QtWidgets.QSpinBox()
        self.scan_points.setRange(3, 200)
        self.io_rate = QtWidgets.QDoubleSpinBox()
        self.io_rate.setRange(500.0, 100000.0)
        self.io_rate.setDecimals(1)
        self.duration_ms = QtWidgets.QSpinBox()
        self.duration_ms.setRange(50, 5000)
        self.ramp_up_slope = QtWidgets.QDoubleSpinBox()
        self.ramp_up_slope.setRange(0.1, 100000.0)
        self.ramp_up_slope.setDecimals(2)
        self.ramp_down_slope = QtWidgets.QDoubleSpinBox()
        self.ramp_down_slope.setRange(0.1, 100000.0)
        self.ramp_down_slope.setDecimals(2)
        self.ramp_down_periods = QtWidgets.QSpinBox()
        self.ramp_down_periods.setRange(1, 20)
        scan_form.addRow("Scan Start Voltage", self.min_clip_amp)
        scan_form.addRow("Scan Max Voltage", self.max_clip_amp)
        scan_form.addRow("Clipping Sine Freq", self.clip_freq)
        scan_form.addRow("Ramp Points", self.scan_points)
        scan_form.addRow("Board IO Rate", self.io_rate)
        scan_form.addRow("Steady Capture (ms)", self.duration_ms)
        scan_form.addRow("Ramp Up Slope", self.ramp_up_slope)
        scan_form.addRow("Ramp Down Slope", self.ramp_down_slope)
        scan_form.addRow("Ramp Down Periods", self.ramp_down_periods)
        layout.addWidget(scan_box)

        hint = QtWidgets.QLabel(
            "The scan follows the VB6 procedure: ramp up, ramp down, fit the monitor sine on each step, then average the two passes before suggesting a new maximum voltage."
        )
        hint.setObjectName("hintText")
        hint.setWordWrap(True)
        layout.addWidget(hint)

        button_row = QtWidgets.QGridLayout()
        self.run_clip_btn = QtWidgets.QPushButton("Start Clipping Auto-Test")
        self.run_clip_btn.setObjectName("accent")
        self.save_limits_btn = QtWidgets.QPushButton("Save Active Limits")
        button_row.addWidget(self.run_clip_btn, 0, 0)
        button_row.addWidget(self.save_limits_btn, 0, 1)
        layout.addLayout(button_row)

        self.coil_group.buttonClicked.connect(self._on_coil_changed)
        self.chk_lock_coils.toggled.connect(self._on_lock_toggled)
        self.quick_connect_btn.clicked.connect(self._connect_backend)
        self.quick_check_btn.clicked.connect(self._check_communication)
        self.run_clip_btn.clicked.connect(self._toggle_clip_test)
        self.save_limits_btn.clicked.connect(self._save_active_limits)

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
        self.backend_status.setSizePolicy(QtWidgets.QSizePolicy.Policy.Ignored, QtWidgets.QSizePolicy.Policy.Preferred)
        layout.addWidget(self.backend_status)

        self.comm_summary = QtWidgets.QLabel("Communication check not run yet.")
        self.comm_summary.setObjectName("valuePill")
        self.comm_summary.setAlignment(QtCore.Qt.AlignmentFlag.AlignCenter)
        self.comm_summary.setWordWrap(True)
        self.comm_summary.setMinimumWidth(0)
        self.comm_summary.setSizePolicy(QtWidgets.QSizePolicy.Policy.Ignored, QtWidgets.QSizePolicy.Policy.Preferred)
        layout.addWidget(self.comm_summary)

        self.relay_status = QtWidgets.QLabel("Relay state unknown.")
        self.relay_status.setObjectName("valuePill")
        self.relay_status.setAlignment(QtCore.Qt.AlignmentFlag.AlignCenter)
        self.relay_status.setWordWrap(True)
        self.relay_status.setMinimumWidth(0)
        self.relay_status.setSizePolicy(QtWidgets.QSizePolicy.Policy.Ignored, QtWidgets.QSizePolicy.Policy.Preferred)
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

        self.btn_browse_bin = QtWidgets.QPushButton("Browse...")
        self.btn_browse_boot = QtWidgets.QPushButton("Browse...")
        self.btn_browse_process = QtWidgets.QPushButton("Browse...")
        self.btn_browse_bin.setMinimumWidth(96)
        self.btn_browse_boot.setMinimumWidth(96)
        self.btn_browse_process.setMinimumWidth(96)

        bin_row = QtWidgets.QHBoxLayout()
        bin_row.addWidget(self.bin_folder)
        bin_row.addWidget(self.btn_browse_bin)
        boot_row = QtWidgets.QHBoxLayout()
        boot_row.addWidget(self.boot_file)
        boot_row.addWidget(self.btn_browse_boot)
        process_row = QtWidgets.QHBoxLayout()
        process_row.addWidget(self.process_file)
        process_row.addWidget(self.btn_browse_process)

        form.addRow("Board #", self.board_num)
        form.addRow("Bin Folder", bin_row)
        form.addRow("Boot File", boot_row)
        form.addRow("Process File", process_row)
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
            "Use the same ADwin process file as the tuner. The clipping scan reuses the shared RAPID ADwin backend and the same relay mapping."
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
        self.console.setSizePolicy(QtWidgets.QSizePolicy.Policy.Preferred, QtWidgets.QSizePolicy.Policy.Expanding)
        layout.addWidget(self.console, stretch=1)
        return frame

    def _build_plot_card(self) -> QtWidgets.QFrame:
        frame = QtWidgets.QFrame()
        frame.setObjectName("card")
        layout = QtWidgets.QVBoxLayout(frame)
        layout.setContentsMargins(18, 18, 18, 18)
        layout.setSpacing(12)

        top_row = QtWidgets.QHBoxLayout()
        title = QtWidgets.QLabel("Clip Curve And Waveform Preview")
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

        hint = QtWidgets.QLabel(
            "Top chart shows the VB6-style clipping curve: residual RMS versus ramp voltage on the left axis, monitor amplitude on the right. Bottom chart previews the highest-voltage waveform capture."
        )
        hint.setObjectName("hintText")
        hint.setWordWrap(True)
        layout.addWidget(hint)

        self.plot_summary = QtWidgets.QLabel(
            "No clipping scan loaded yet. Run the clip test to populate the RMS and monitor curves, then save new limits if the result looks correct."
        )
        self.plot_summary.setObjectName("plotSummary")
        self.plot_summary.setWordWrap(True)
        layout.addWidget(self.plot_summary)

        self.plot_splitter = QtWidgets.QSplitter(QtCore.Qt.Orientation.Vertical)
        self.plot_splitter.setChildrenCollapsible(False)
        self.plot_splitter.setHandleWidth(8)
        layout.addWidget(self.plot_splitter, stretch=1)

        self.clip_plot = pg.PlotWidget()
        self.wave_plot = pg.PlotWidget()
        self.clip_plot.setMinimumHeight(320)
        self.wave_plot.setMinimumHeight(220)
        self.plot_splitter.addWidget(self.clip_plot)
        self.plot_splitter.addWidget(self.wave_plot)
        self.plot_splitter.setSizes([470, 280])

        self._configure_plot_widget(self.clip_plot)
        self._configure_plot_widget(self.wave_plot)

        clip_item = self.clip_plot.getPlotItem()
        clip_item.setLabel("left", "Residual RMS", units="V")
        clip_item.setLabel("bottom", "Ramp Voltage", units="V")
        clip_item.showAxis("right")
        clip_item.getAxis("right").setLabel("Monitor Amplitude", units="V")
        clip_item.addLegend(offset=(12, 12))

        self.monitor_view = pg.ViewBox()
        clip_item.scene().addItem(self.monitor_view)
        clip_item.getAxis("right").linkToView(self.monitor_view)
        self.monitor_view.setXLink(clip_item.vb)
        clip_item.vb.sigResized.connect(self._sync_clip_views)

        self.residual_curve = clip_item.plot(name="Residual RMS", pen=pg.mkPen("#1A237E", width=2.4))
        self.monitor_curve = pg.PlotCurveItem(name="Monitor Amplitude", pen=pg.mkPen("#7A0219", width=2.4))
        self.monitor_view.addItem(self.monitor_curve)
        self.residual_curve.setClipToView(True)

        wave_item = self.wave_plot.getPlotItem()
        wave_item.setLabel("left", "Voltage", units="V")
        wave_item.setLabel("bottom", "Time", units="ms")
        wave_item.addLegend(offset=(12, 12))
        self.wave_dac_curve = wave_item.plot(name="DAC", pen=pg.mkPen("#7A0219", width=2.0))
        self.wave_adc_curve = wave_item.plot(name="ADC", pen=pg.mkPen("#1A237E", width=2.0))
        self.wave_dac_curve.setClipToView(True)
        self.wave_adc_curve.setClipToView(True)

        self.save_plot_btn.clicked.connect(self._save_plot_image)
        self.save_csv_btn.clicked.connect(self._save_plot_csv)
        self.clear_plot_btn.clicked.connect(self._clear_plot_state)

        return frame

    def _configure_plot_widget(self, plot: pg.PlotWidget) -> None:
        plot.setBackground((0, 0, 0, 0))
        item = plot.getPlotItem()
        item.showGrid(x=True, y=True, alpha=0.15)
        item.getAxis("left").setTextPen(pg.mkPen("#4D3A39"))
        item.getAxis("bottom").setTextPen(pg.mkPen("#4D3A39"))
        item.getAxis("left").setPen(pg.mkPen("#8B6F68"))
        item.getAxis("bottom").setPen(pg.mkPen("#8B6F68"))
        item.getAxis("left").enableAutoSIPrefix(False)
        item.getAxis("bottom").enableAutoSIPrefix(False)

    def _sync_clip_views(self) -> None:
        clip_item = self.clip_plot.getPlotItem()
        self.monitor_view.setGeometry(clip_item.vb.sceneBoundingRect())
        self.monitor_view.linkedViewChanged(clip_item.vb, self.monitor_view.XAxis)

    def _is_axial(self) -> bool:
        return self.axial_radio.isChecked()

    def _active_coil_name(self) -> str:
        return "axial" if self._is_axial() else "transverse"

    def _active_saved_freq(self) -> float:
        return self._coil_config.axial_res_freq if self._is_axial() else self._coil_config.trans_res_freq

    def _active_saved_ramp_limit(self) -> float:
        return self._coil_config.axial_max_ramp if self._is_axial() else self._coil_config.trans_max_ramp

    def _active_saved_monitor_limit(self) -> float:
        return self._coil_config.axial_max_monitor if self._is_axial() else self._coil_config.trans_max_monitor

    def _runtime_limits(self) -> AdwinCoilLimits:
        return AdwinCoilLimits(
            axial_ramp_max=self._coil_config.axial_max_ramp,
            axial_monitor_max=self._coil_config.axial_max_monitor,
            trans_ramp_max=self._coil_config.trans_max_ramp,
            trans_monitor_max=self._coil_config.trans_max_monitor,
        )

    def _append(self, text: str) -> None:
        stamp = QtCore.QDateTime.currentDateTime().toString("HH:mm:ss")
        self.console.appendPlainText(f"[{stamp}] {text}")
        scrollbar = self.console.verticalScrollBar()
        scrollbar.setValue(scrollbar.maximum())
        self.status.setText(text)

    def _set_backend_status(self, text: str, color: str) -> None:
        style = (
            f"background: rgba(255,255,255,0.82); border: 1px solid rgba(122,2,25,0.16); "
            f"border-radius: 16px; padding: 8px 10px; font-weight: 650; color: {color};"
        )
        self.backend_status.setText(text)
        self.backend_status.setStyleSheet(style)
        self.quick_backend_status.setText(text)
        self.quick_backend_status.setStyleSheet(style)

    def _set_comm_summary(self, text: str, color: str) -> None:
        self.comm_summary.setText(text)
        self.comm_summary.setStyleSheet(
            f"background: rgba(255,255,255,0.82); border: 1px solid rgba(122,2,25,0.16); "
            f"border-radius: 16px; padding: 8px 10px; font-weight: 650; color: {color};"
        )

    def _set_relay_status(self, text: str, color: str) -> None:
        style = (
            f"background: rgba(255,255,255,0.82); border: 1px solid rgba(122,2,25,0.16); "
            f"border-radius: 16px; padding: 8px 10px; font-weight: 650; color: {color};"
        )
        self.relay_status.setText(text)
        self.relay_status.setStyleSheet(style)
        self.quick_relay_status.setText(text)
        self.quick_relay_status.setStyleSheet(style)

    def _set_busy(self, busy: bool) -> None:
        for button in (
            self.connect_btn,
            self.check_comm_btn,
            self.quick_connect_btn,
            self.quick_check_btn,
            self.apply_relay_btn,
            self.relays_off_btn,
            self.save_backend_btn,
            self.save_limits_btn,
        ):
            button.setEnabled(not busy)
        self.run_clip_btn.setText("Stop Clipping Test") if busy else self.run_clip_btn.setText("Start Clipping Auto-Test")

    def _load_into_widgets(self) -> None:
        self.board_num.setValue(self._backend_config.board_num)
        self.bin_folder.setText(self._backend_config.bin_folder)
        self.boot_file.setText(self._backend_config.boot_file)
        self.process_file.setText(self._backend_config.process_file or _default_process_file())
        self.ramp_chan.setValue(self._backend_config.ramp_dac_chan)
        self.monitor_chan.setValue(self._backend_config.monitor_adc_chan)
        self.axial_relay_bit.setValue(self._backend_config.axial_relay_bit)
        self.trans_relay_bit.setValue(self._backend_config.trans_relay_bit)

        self.min_clip_amp.setValue(self._clip_test_config.min_amp_v)
        self.max_clip_amp.setValue(self._clip_test_config.max_amp_v)
        self.clip_freq.setValue(self._active_saved_freq())
        self.scan_points.setValue(self._clip_test_config.scan_points)
        self.io_rate.setValue(self._clip_test_config.io_rate_hz)
        self.duration_ms.setValue(self._clip_test_config.duration_ms)
        self.ramp_up_slope.setValue(self._clip_test_config.ramp_up_slope_vps)
        self.ramp_down_slope.setValue(self._clip_test_config.ramp_down_slope_vps)
        self.ramp_down_periods.setValue(self._clip_test_config.ramp_down_periods)
        self._refresh_limit_widgets()

    def _refresh_limit_widgets(self) -> None:
        coil = self._active_coil_name()
        self.saved_res_freq.setText(f"{self._active_saved_freq():.4f} Hz")
        self.saved_max_ramp.setText(f"{self._active_saved_ramp_limit():.4f} V")
        self.saved_max_monitor.setText(f"{self._active_saved_monitor_limit():.4f} V")
        self.fit_max_ramp.setValue(float(self._fit_ramp_by_coil[coil]))
        self.fit_max_monitor.setValue(float(self._fit_monitor_by_coil[coil]))
        self.clip_freq.setValue(self._active_saved_freq())

    def _build_backend_from_widgets(self) -> BackendConfig:
        return BackendConfig(
            board_num=self.board_num.value(),
            bin_folder=self.bin_folder.text().strip(),
            boot_file=self.boot_file.text().strip() or "ADwin9.btl",
            process_file=self.process_file.text().strip() or _default_process_file(),
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
                "Point the ADwin process file at VB6/ADwin/sineout.T91 before connecting.",
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
            self._ctrl = controller
            self._connected = True
            self._last_version = max(self._last_version, version)
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
        self._backend_config = self._build_backend_from_widgets()
        save_dataclass(BACKEND_CONFIG_PATH, self._backend_config)
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

    def _on_lock_toggled(self, checked: bool) -> None:
        self._set_coil_locked(checked)

    def _on_coil_changed(self, *_args) -> None:
        if self.chk_lock_coils.isChecked():
            return
        self._refresh_limit_widgets()
        self._apply_relays()

    def _toggle_clip_test(self) -> None:
        if self._worker is not None:
            self._worker.stop()
            self._append("Stopping clipping test...")
            return
        self._start_clip_test()

    def _start_clip_test(self) -> None:
        min_amp = self.min_clip_amp.value()
        max_amp = self.max_clip_amp.value()
        freq_hz = self.clip_freq.value()
        if min_amp >= max_amp:
            QtWidgets.QMessageBox.warning(
                self,
                "Invalid Scan Range",
                "Scan start voltage must be lower than the scan max voltage.",
            )
            return
        if freq_hz <= 0:
            QtWidgets.QMessageBox.warning(
                self,
                "Invalid Frequency",
                "Clipping sine frequency must be greater than zero.",
            )
            return
        self._clip_test_config.min_amp_v = min_amp
        self._clip_test_config.max_amp_v = max_amp
        self._clip_test_config.sine_freq_hz = freq_hz
        self._clip_test_config.scan_points = self.scan_points.value()
        self._clip_test_config.io_rate_hz = self.io_rate.value()
        self._clip_test_config.duration_ms = self.duration_ms.value()
        self._clip_test_config.ramp_up_slope_vps = self.ramp_up_slope.value()
        self._clip_test_config.ramp_down_slope_vps = self.ramp_down_slope.value()
        self._clip_test_config.ramp_down_periods = self.ramp_down_periods.value()
        save_dataclass(CLIP_TEST_CONFIG_PATH, self._clip_test_config)

        self._backend_config = self._build_backend_from_widgets()
        worker = AutoClipWorker(
            backend=self._backend_config,
            limits=self._runtime_limits(),
            coil=self._active_coil_name(),
            config=self._clip_test_config,
        )
        thread = QtCore.QThread(self)
        worker.moveToThread(thread)
        thread.started.connect(worker.run)
        worker.progress.connect(self._append)
        worker.result_ready.connect(self._on_clip_result_ready)
        worker.failed.connect(self._on_worker_failed)
        worker.finished.connect(thread.quit)
        thread.finished.connect(self._cleanup_worker)

        self._worker = worker
        self._worker_thread = thread
        self._set_busy(True)
        self._append(
            f"Starting clipping test for {self._active_coil_name()} coil at {freq_hz:.4f} Hz over {self.scan_points.value()} ramp points."
        )
        thread.start()

    def _on_clip_result_ready(self, result: ClipScanResult) -> None:
        self._last_result = result
        if result.suggested_point is not None:
            self._fit_ramp_by_coil[result.coil] = result.suggested_point.ramp_voltage_v
            self._fit_monitor_by_coil[result.coil] = result.suggested_point.monitor_amplitude_v
            if result.coil == self._active_coil_name():
                self.fit_max_ramp.setValue(result.suggested_point.ramp_voltage_v)
                self.fit_max_monitor.setValue(result.suggested_point.monitor_amplitude_v)
        self._refresh_clip_plot()
        self._refresh_wave_plot()
        if result.suggested_point is not None:
            self.plot_summary.setText(
                f"Suggested limit for {result.coil.title()} coil: ramp {result.suggested_point.ramp_voltage_v:.4f} V, "
                f"monitor {result.suggested_point.monitor_amplitude_v:.4f} V, residual {result.suggested_point.residual_rms_v:.5f} V RMS."
            )
            self._append(
                f"Clipping test complete. Suggested {result.coil} limit: ramp {result.suggested_point.ramp_voltage_v:.4f} V, "
                f"monitor {result.suggested_point.monitor_amplitude_v:.4f} V."
            )
        else:
            self.plot_summary.setText("Clipping test complete, but no valid points were available to suggest a new limit.")
            self._append("Clipping test complete, but no valid points were available to suggest a new limit.")

    def _on_worker_failed(self, message: str) -> None:
        lowered = message.lower()
        if "stopped by user" in lowered:
            self._append(message)
            return
        self._append(f"[ERROR] Clipping test failed: {message}")

    def _cleanup_worker(self) -> None:
        if self._worker_thread is not None:
            self._worker_thread.deleteLater()
        if self._worker is not None:
            self._worker.deleteLater()
        self._worker_thread = None
        self._worker = None
        self._set_busy(False)

    def _refresh_clip_plot(self) -> None:
        if self._last_result is None:
            self.residual_curve.setData([], [])
            self.monitor_curve.setData([], [])
            return
        ramp_v = [point.ramp_voltage_v for point in self._last_result.avg_points]
        residual_v = [point.residual_rms_v for point in self._last_result.avg_points]
        monitor_v = [point.monitor_amplitude_v for point in self._last_result.avg_points]
        self.residual_curve.setData(ramp_v, residual_v)
        self.monitor_curve.setData(ramp_v, monitor_v)
        self.clip_plot.getPlotItem().enableAutoRange()
        self.monitor_view.enableAutoRange()

    def _refresh_wave_plot(self) -> None:
        if self._last_result is None or self._last_result.preview_capture is None:
            self.wave_dac_curve.setData([], [])
            self.wave_adc_curve.setData([], [])
            return
        capture = self._last_result.preview_capture
        time_ms = [value * 1000.0 for value in capture.time_s]
        self.wave_dac_curve.setData(time_ms, capture.dac_v)
        self.wave_adc_curve.setData(time_ms, capture.adc_v)
        self.wave_plot.getPlotItem().enableAutoRange()

    def _save_active_limits(self) -> None:
        coil = self._active_coil_name()
        ramp_limit = self.fit_max_ramp.value()
        monitor_limit = self.fit_max_monitor.value()
        if ramp_limit <= 0 or monitor_limit <= 0:
            QtWidgets.QMessageBox.warning(
                self,
                "Invalid Limits",
                "Suggested max ramp and max monitor values must both be positive before saving.",
            )
            return
        self._fit_ramp_by_coil[coil] = ramp_limit
        self._fit_monitor_by_coil[coil] = monitor_limit
        if coil == "axial":
            self._coil_config.axial_max_ramp = ramp_limit
            self._coil_config.axial_max_monitor = monitor_limit
        else:
            self._coil_config.trans_max_ramp = ramp_limit
            self._coil_config.trans_max_monitor = monitor_limit
        save_dataclass(COIL_CONFIG_PATH, self._coil_config)
        self._refresh_limit_widgets()
        self._append(
            f"Saved {coil} clip limits: max ramp {ramp_limit:.4f} V, max monitor {monitor_limit:.4f} V."
        )

    def _clear_plot_state(self) -> None:
        self._last_result = None
        self.residual_curve.setData([], [])
        self.monitor_curve.setData([], [])
        self.wave_dac_curve.setData([], [])
        self.wave_adc_curve.setData([], [])
        self.plot_summary.setText(
            "No clipping scan loaded yet. Run the clip test to populate the RMS and monitor curves, then save new limits if the result looks correct."
        )

    def _save_plot_image(self) -> None:
        path, _ = QtWidgets.QFileDialog.getSaveFileName(
            self,
            "Save Plot Image",
            str(Path.home() / "af_clip_test.png"),
            "PNG image (*.png)",
        )
        if not path:
            return
        if not self.plot_splitter.grab().save(path):
            self._append(f"[ERROR] Failed to save plot image: {path}")
            return
        self._append(f"Saved plot image to {path}")

    def _save_plot_csv(self) -> None:
        if self._last_result is None:
            QtWidgets.QMessageBox.information(self, "No Data", "Run a clipping test before exporting CSV data.")
            return
        path, _ = QtWidgets.QFileDialog.getSaveFileName(
            self,
            "Save Clip CSV",
            str(Path.home() / "af_clip_test.csv"),
            "CSV files (*.csv)",
        )
        if not path:
            return
        up_map = {round(point.ramp_voltage_v, 6): point for point in self._last_result.up_points}
        down_map = {round(point.ramp_voltage_v, 6): point for point in self._last_result.down_points}
        try:
            with open(path, "w", newline="", encoding="utf-8") as handle:
                writer = csv.writer(handle)
                writer.writerow(
                    [
                        "ramp_voltage_v",
                        "avg_monitor_amplitude_v",
                        "avg_residual_rms_v",
                        "up_monitor_amplitude_v",
                        "up_residual_rms_v",
                        "down_monitor_amplitude_v",
                        "down_residual_rms_v",
                    ]
                )
                for point in self._last_result.avg_points:
                    key = round(point.ramp_voltage_v, 6)
                    up_point = up_map.get(key)
                    down_point = down_map.get(key)
                    writer.writerow(
                        [
                            f"{point.ramp_voltage_v:.6f}",
                            f"{point.monitor_amplitude_v:.6f}" if math.isfinite(point.monitor_amplitude_v) else "",
                            f"{point.residual_rms_v:.6f}" if math.isfinite(point.residual_rms_v) else "",
                            f"{up_point.monitor_amplitude_v:.6f}" if up_point and math.isfinite(up_point.monitor_amplitude_v) else "",
                            f"{up_point.residual_rms_v:.6f}" if up_point and math.isfinite(up_point.residual_rms_v) else "",
                            f"{down_point.monitor_amplitude_v:.6f}" if down_point and math.isfinite(down_point.monitor_amplitude_v) else "",
                            f"{down_point.residual_rms_v:.6f}" if down_point and math.isfinite(down_point.residual_rms_v) else "",
                        ]
                    )
            self._append(f"Saved clip CSV to {path}")
        except OSError as exc:
            self._append(f"[ERROR] Failed to save clip CSV: {exc}")


def main() -> int:
    app = QtWidgets.QApplication(sys.argv)
    apply_liquid_glass_theme(app)
    window = MainWindow()
    set_app_icon(window, "af_clip_test_icon.ico", _assets_dir())
    window.show()
    return app.exec()
