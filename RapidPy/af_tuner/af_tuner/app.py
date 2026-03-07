from __future__ import annotations

import json
import sys
from dataclasses import asdict, dataclass
from pathlib import Path

from PySide6 import QtCore, QtWidgets


def _bootstrap_common_imports() -> None:
    root = Path(__file__).resolve().parents[2]
    if str(root) not in sys.path:
        sys.path.insert(0, str(root))


_bootstrap_common_imports()
from rapidpy_common.adwin_af import (  # noqa: E402
    AdwinAFController,
    AdwinBoardConfig,
    AdwinCoilLimits,
    AdwinRampRequest,
)
from rapidpy_common.ui import apply_card_shadow, apply_liquid_glass_theme  # noqa: E402


@dataclass(slots=True)
class CoilTuningConfig:
    axial_res_freq: float = 1000.0
    axial_max_ramp: float = 1.0
    axial_max_monitor: float = 1.0
    trans_res_freq: float = 1000.0
    trans_max_ramp: float = 1.0
    trans_max_monitor: float = 1.0


@dataclass(slots=True)
class AutoTuneConfig:
    low_freq: float = 900.0
    high_freq: float = 1100.0
    step_freq: float = 10.0
    hold_ms: int = 500
    io_rate_hz: float = 1000.0


@dataclass(slots=True)
class BackendConfig:
    board_num: int = 1
    bin_folder: str = ""
    boot_file: str = "ADwin9.btl"
    process_file: str = "AF_Ramp_System.abp"
    ramp_dac_chan: int = 1
    monitor_adc_chan: int = 1
    axial_relay_bit: int = 0
    trans_relay_bit: int = 1


COIL_CONFIG_PATH = Path.home() / ".rapidpy_af_tuner.json"
BACKEND_CONFIG_PATH = Path.home() / ".rapidpy_af_backend.json"
AUTOTUNE_CONFIG_PATH = Path.home() / ".rapidpy_af_autotune.json"


class AutoTuneWorker(QtCore.QObject):
    progress = QtCore.Signal(str)
    finished = QtCore.Signal(float, float)
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
        self._backend = backend
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
            controller = AdwinAFController(
                board=AdwinBoardConfig(
                    board_num=self._backend.board_num,
                    bin_folder=self._backend.bin_folder,
                    boot_file=self._backend.boot_file,
                    process_file=self._backend.process_file,
                    ramp_dac_chan=self._backend.ramp_dac_chan,
                    monitor_adc_chan=self._backend.monitor_adc_chan,
                    axial_relay_bit=self._backend.axial_relay_bit,
                    trans_relay_bit=self._backend.trans_relay_bit,
                ),
                limits=self._limits,
            )
            duration_s = max(float(self._hold_ms) / 1000.0, 0.001)
            slope = self._ramp_peak_v / duration_s

            best_freq = self._low
            best_amp = -1.0
            freq = self._low

            while freq <= self._high + 1e-9:
                if self._abort:
                    self.progress.emit("Auto-tune aborted.")
                    self.finished.emit(best_freq, best_amp)
                    return

                self.progress.emit(f"Sweeping {freq:.4f} Hz...")
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

                amp = result.monitor_peak_v
                self.progress.emit(
                    f"{freq:.4f} Hz -> monitor_peak={amp:.5f} V, ramp_peak={result.ramp_peak_v:.5f} V"
                )
                if amp > best_amp:
                    best_amp = amp
                    best_freq = freq

                freq += self._step

            self.finished.emit(best_freq, best_amp)
        except Exception as exc:
            self.failed.emit(str(exc))

    def abort(self) -> None:
        self._abort = True


def load_dataclass(path: Path, cls):
    if not path.exists():
        return cls()
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
        base = asdict(cls())
        for key in base:
            if key in data:
                base[key] = data[key]
        return cls(**base)
    except (OSError, json.JSONDecodeError):
        return cls()


def save_dataclass(path: Path, obj) -> None:
    try:
        path.write_text(json.dumps(asdict(obj), indent=2, sort_keys=True), encoding="utf-8")
    except OSError:
        return


class MainWindow(QtWidgets.QMainWindow):
    def __init__(self) -> None:
        super().__init__()
        self.setWindowTitle("RapidPy AF Tuner")
        self.resize(1180, 760)

        self._coil_config = load_dataclass(COIL_CONFIG_PATH, CoilTuningConfig)
        self._backend_config = load_dataclass(BACKEND_CONFIG_PATH, BackendConfig)
        self._autotune_config = load_dataclass(AUTOTUNE_CONFIG_PATH, AutoTuneConfig)

        self._thread: QtCore.QThread | None = None
        self._worker: AutoTuneWorker | None = None

        self._build_ui()
        self._load_into_widgets()

    def _build_ui(self) -> None:
        root = QtWidgets.QWidget(self)
        self.setCentralWidget(root)
        layout = QtWidgets.QHBoxLayout(root)
        layout.setContentsMargins(16, 16, 16, 16)
        layout.setSpacing(14)

        controls = QtWidgets.QFrame()
        controls.setObjectName("card")
        controls_layout = QtWidgets.QVBoxLayout(controls)
        controls_layout.setContentsMargins(18, 18, 18, 18)
        controls_layout.setSpacing(10)

        title = QtWidgets.QLabel("AF Coil Tuner")
        title.setObjectName("title")
        subtitle = QtWidgets.QLabel("VB6-style coil save/apply + ADWIN auto-tune")
        subtitle.setObjectName("subtitle")
        controls_layout.addWidget(title)
        controls_layout.addWidget(subtitle)

        self.coil_group = QtWidgets.QButtonGroup(self)
        coil_row = QtWidgets.QHBoxLayout()
        self.axial_radio = QtWidgets.QRadioButton("Axial Coil")
        self.trans_radio = QtWidgets.QRadioButton("Transverse Coil")
        self.axial_radio.setChecked(True)
        self.coil_group.addButton(self.axial_radio)
        self.coil_group.addButton(self.trans_radio)
        coil_row.addWidget(self.axial_radio)
        coil_row.addWidget(self.trans_radio)
        controls_layout.addLayout(coil_row)

        self.old_freq = QtWidgets.QDoubleSpinBox()
        self.old_freq.setReadOnly(True)
        self.old_freq.setButtonSymbols(QtWidgets.QAbstractSpinBox.NoButtons)
        self.new_freq = QtWidgets.QDoubleSpinBox()
        self.new_freq.setRange(0.001, 250_000)
        self.new_freq.setDecimals(4)

        self.old_ramp = QtWidgets.QDoubleSpinBox()
        self.old_ramp.setReadOnly(True)
        self.old_ramp.setButtonSymbols(QtWidgets.QAbstractSpinBox.NoButtons)
        self.new_ramp = QtWidgets.QDoubleSpinBox()
        self.new_ramp.setRange(0.001, 10.0)
        self.new_ramp.setDecimals(4)

        self.old_monitor = QtWidgets.QDoubleSpinBox()
        self.old_monitor.setReadOnly(True)
        self.old_monitor.setButtonSymbols(QtWidgets.QAbstractSpinBox.NoButtons)
        self.new_monitor = QtWidgets.QDoubleSpinBox()
        self.new_monitor.setRange(0.001, 10.0)
        self.new_monitor.setDecimals(4)

        form = QtWidgets.QFormLayout()
        form.addRow("Old Resonance Freq", self.old_freq)
        form.addRow("New Resonance Freq", self.new_freq)
        form.addRow("Old Max Ramp", self.old_ramp)
        form.addRow("New Max Ramp", self.new_ramp)
        form.addRow("Old Max Monitor", self.old_monitor)
        form.addRow("New Max Monitor", self.new_monitor)
        controls_layout.addLayout(form)

        self.low_freq = QtWidgets.QDoubleSpinBox()
        self.low_freq.setRange(0.001, 250_000)
        self.low_freq.setDecimals(4)
        self.high_freq = QtWidgets.QDoubleSpinBox()
        self.high_freq.setRange(0.001, 250_000)
        self.high_freq.setDecimals(4)
        self.step_freq = QtWidgets.QDoubleSpinBox()
        self.step_freq.setRange(0.001, 20_000)
        self.step_freq.setDecimals(4)
        self.hold_ms = QtWidgets.QSpinBox()
        self.hold_ms.setRange(0, 60_000)
        self.io_rate = QtWidgets.QDoubleSpinBox()
        self.io_rate.setRange(1.0, 200_000)
        self.io_rate.setDecimals(3)

        tune_form = QtWidgets.QFormLayout()
        tune_form.addRow("Sweep Low Freq", self.low_freq)
        tune_form.addRow("Sweep High Freq", self.high_freq)
        tune_form.addRow("Sweep Step", self.step_freq)
        tune_form.addRow("Peak Hang (ms)", self.hold_ms)
        tune_form.addRow("Board IO Rate (Hz)", self.io_rate)
        controls_layout.addLayout(tune_form)

        button_grid = QtWidgets.QGridLayout()
        self.apply_freq_btn = QtWidgets.QPushButton("Apply Freq")
        self.apply_volt_btn = QtWidgets.QPushButton("Apply Max Volt")
        self.save_freq_btn = QtWidgets.QPushButton("Save Freq")
        self.save_volt_btn = QtWidgets.QPushButton("Save Max Volt")
        self.auto_tune_btn = QtWidgets.QPushButton("Start Auto-Tune")
        self.auto_tune_btn.setObjectName("accent")
        button_grid.addWidget(self.apply_freq_btn, 0, 0)
        button_grid.addWidget(self.apply_volt_btn, 0, 1)
        button_grid.addWidget(self.save_freq_btn, 1, 0)
        button_grid.addWidget(self.save_volt_btn, 1, 1)
        button_grid.addWidget(self.auto_tune_btn, 2, 0, 1, 2)
        controls_layout.addLayout(button_grid)

        self.status = QtWidgets.QLabel("Ready")
        self.status.setObjectName("valuePill")
        controls_layout.addWidget(self.status)

        self.console = QtWidgets.QPlainTextEdit()
        self.console.setReadOnly(True)
        self.console.setObjectName("console")
        controls_layout.addWidget(self.console, stretch=1)

        right = QtWidgets.QFrame()
        right.setObjectName("card")
        right_layout = QtWidgets.QVBoxLayout(right)
        right_layout.setContentsMargins(18, 18, 18, 18)
        right_layout.setSpacing(10)

        right_title = QtWidgets.QLabel("ADWIN Backend")
        right_title.setObjectName("title")
        right_title.setStyleSheet("font-size:20px;")
        right_layout.addWidget(right_title)

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

        backend_form = QtWidgets.QFormLayout()
        backend_form.addRow("Board #", self.board_num)
        backend_form.addRow("Bin Folder", self.bin_folder)
        backend_form.addRow("Boot File", self.boot_file)
        backend_form.addRow("Process File", self.process_file)
        backend_form.addRow("Ramp DAC Chan", self.ramp_chan)
        backend_form.addRow("Monitor ADC Chan", self.monitor_chan)
        backend_form.addRow("Axial Relay Bit", self.axial_relay_bit)
        backend_form.addRow("Trans Relay Bit", self.trans_relay_bit)
        right_layout.addLayout(backend_form)

        self.save_backend_btn = QtWidgets.QPushButton("Save Backend Settings")
        right_layout.addWidget(self.save_backend_btn)
        right_layout.addStretch(1)

        layout.addWidget(controls, stretch=2)
        layout.addWidget(right, stretch=1)

        apply_card_shadow(controls)
        apply_card_shadow(right)

        self.coil_group.buttonClicked.connect(self._on_coil_changed)
        self.apply_freq_btn.clicked.connect(self._apply_freq)
        self.apply_volt_btn.clicked.connect(self._apply_volts)
        self.save_freq_btn.clicked.connect(self._save_freq)
        self.save_volt_btn.clicked.connect(self._save_volts)
        self.auto_tune_btn.clicked.connect(self._toggle_autotune)
        self.save_backend_btn.clicked.connect(self._save_backend)

    def _is_axial(self) -> bool:
        return self.axial_radio.isChecked()

    def _load_into_widgets(self) -> None:
        if self._is_axial():
            self.old_freq.setValue(self._coil_config.axial_res_freq)
            self.new_freq.setValue(self._coil_config.axial_res_freq)
            self.old_ramp.setValue(self._coil_config.axial_max_ramp)
            self.new_ramp.setValue(self._coil_config.axial_max_ramp)
            self.old_monitor.setValue(self._coil_config.axial_max_monitor)
            self.new_monitor.setValue(self._coil_config.axial_max_monitor)
        else:
            self.old_freq.setValue(self._coil_config.trans_res_freq)
            self.new_freq.setValue(self._coil_config.trans_res_freq)
            self.old_ramp.setValue(self._coil_config.trans_max_ramp)
            self.new_ramp.setValue(self._coil_config.trans_max_ramp)
            self.old_monitor.setValue(self._coil_config.trans_max_monitor)
            self.new_monitor.setValue(self._coil_config.trans_max_monitor)

        self.low_freq.setValue(self._autotune_config.low_freq)
        self.high_freq.setValue(self._autotune_config.high_freq)
        self.step_freq.setValue(self._autotune_config.step_freq)
        self.hold_ms.setValue(self._autotune_config.hold_ms)
        self.io_rate.setValue(self._autotune_config.io_rate_hz)

        self.board_num.setValue(self._backend_config.board_num)
        self.bin_folder.setText(self._backend_config.bin_folder)
        self.boot_file.setText(self._backend_config.boot_file)
        self.process_file.setText(self._backend_config.process_file)
        self.ramp_chan.setValue(self._backend_config.ramp_dac_chan)
        self.monitor_chan.setValue(self._backend_config.monitor_adc_chan)
        self.axial_relay_bit.setValue(self._backend_config.axial_relay_bit)
        self.trans_relay_bit.setValue(self._backend_config.trans_relay_bit)

    def _append(self, text: str) -> None:
        self.console.appendPlainText(text)
        self.status.setText(text)

    def _active_coil_name(self) -> str:
        return "axial" if self._is_axial() else "transverse"

    def _active_limits(self) -> AdwinCoilLimits:
        return AdwinCoilLimits(
            axial_ramp_max=self._coil_config.axial_max_ramp,
            axial_monitor_max=self._coil_config.axial_max_monitor,
            trans_ramp_max=self._coil_config.trans_max_ramp,
            trans_monitor_max=self._coil_config.trans_max_monitor,
        )

    def _apply_freq(self) -> None:
        self._append(
            f"Applied {self._active_coil_name()} resonance frequency: {self.new_freq.value():.4f} Hz"
        )

    def _apply_volts(self) -> None:
        self._append(
            f"Applied {self._active_coil_name()} max voltages: "
            f"ramp={self.new_ramp.value():.4f} V, monitor={self.new_monitor.value():.4f} V"
        )

    def _save_freq(self) -> None:
        if self.new_freq.value() <= 0:
            QtWidgets.QMessageBox.warning(self, "Invalid Value", "Resonance frequency must be positive.")
            return
        if self._is_axial():
            self._coil_config.axial_res_freq = self.new_freq.value()
            self.old_freq.setValue(self._coil_config.axial_res_freq)
        else:
            self._coil_config.trans_res_freq = self.new_freq.value()
            self.old_freq.setValue(self._coil_config.trans_res_freq)
        save_dataclass(COIL_CONFIG_PATH, self._coil_config)
        self._append("Frequency saved to local config.")

    def _save_volts(self) -> None:
        if self.new_ramp.value() <= 0 or self.new_monitor.value() <= 0:
            QtWidgets.QMessageBox.warning(self, "Invalid Value", "Max ramp and monitor voltages must be positive.")
            return
        if self._is_axial():
            self._coil_config.axial_max_ramp = self.new_ramp.value()
            self._coil_config.axial_max_monitor = self.new_monitor.value()
            self.old_ramp.setValue(self._coil_config.axial_max_ramp)
            self.old_monitor.setValue(self._coil_config.axial_max_monitor)
        else:
            self._coil_config.trans_max_ramp = self.new_ramp.value()
            self._coil_config.trans_max_monitor = self.new_monitor.value()
            self.old_ramp.setValue(self._coil_config.trans_max_ramp)
            self.old_monitor.setValue(self._coil_config.trans_max_monitor)
        save_dataclass(COIL_CONFIG_PATH, self._coil_config)
        self._append("Max voltages saved to local config.")

    def _save_backend(self) -> None:
        self._backend_config.board_num = self.board_num.value()
        self._backend_config.bin_folder = self.bin_folder.text().strip()
        self._backend_config.boot_file = self.boot_file.text().strip()
        self._backend_config.process_file = self.process_file.text().strip()
        self._backend_config.ramp_dac_chan = self.ramp_chan.value()
        self._backend_config.monitor_adc_chan = self.monitor_chan.value()
        self._backend_config.axial_relay_bit = self.axial_relay_bit.value()
        self._backend_config.trans_relay_bit = self.trans_relay_bit.value()
        save_dataclass(BACKEND_CONFIG_PATH, self._backend_config)
        self._append("Saved ADWIN backend settings.")

    def _on_coil_changed(self, *_args) -> None:
        self._load_into_widgets()
        self._save_backend()
        self._apply_relays()

    def _apply_relays(self) -> None:
        try:
            controller = AdwinAFController(
                board=AdwinBoardConfig(
                    board_num=self._backend_config.board_num,
                    bin_folder=self._backend_config.bin_folder,
                    boot_file=self._backend_config.boot_file,
                    process_file=self._backend_config.process_file,
                    ramp_dac_chan=self._backend_config.ramp_dac_chan,
                    monitor_adc_chan=self._backend_config.monitor_adc_chan,
                    axial_relay_bit=self._backend_config.axial_relay_bit,
                    trans_relay_bit=self._backend_config.trans_relay_bit,
                ),
                limits=self._active_limits(),
            )
            bit_value = controller.set_af_relays(self._active_coil_name())
            self._append(f"AF relays set for {self._active_coil_name()} coil (bit value {bit_value}).")
        except Exception as exc:
            self._append(f"AF relay update skipped: {exc}")

    def _toggle_autotune(self) -> None:
        if self._thread is not None and self._worker is not None:
            self._worker.abort()
            self.auto_tune_btn.setEnabled(False)
            self._append("Stopping auto-tune...")
            return

        low = self.low_freq.value()
        high = self.high_freq.value()
        step = self.step_freq.value()
        if low <= 0 or high <= 0 or step <= 0 or high < low:
            QtWidgets.QMessageBox.warning(self, "Invalid Sweep", "Sweep values must satisfy 0 < low <= high and step > 0.")
            return

        self._autotune_config.low_freq = low
        self._autotune_config.high_freq = high
        self._autotune_config.step_freq = step
        self._autotune_config.hold_ms = self.hold_ms.value()
        self._autotune_config.io_rate_hz = self.io_rate.value()
        save_dataclass(AUTOTUNE_CONFIG_PATH, self._autotune_config)

        self._save_backend()
        self._apply_relays()

        self._thread = QtCore.QThread(self)
        self._worker = AutoTuneWorker(
            backend=self._backend_config,
            limits=self._active_limits(),
            low_freq=low,
            high_freq=high,
            step_freq=step,
            hold_ms=self.hold_ms.value(),
            io_rate_hz=self.io_rate.value(),
            ramp_peak_v=self.new_ramp.value(),
            monitor_peak_v=self.new_monitor.value(),
            coil=self._active_coil_name(),
        )
        self._worker.moveToThread(self._thread)

        self._thread.started.connect(self._worker.run)
        self._worker.progress.connect(self._append)
        self._worker.finished.connect(self._autotune_finished)
        self._worker.failed.connect(self._autotune_failed)
        self._worker.finished.connect(self._thread.quit)
        self._worker.failed.connect(self._thread.quit)
        self._thread.finished.connect(self._cleanup_worker)

        self.auto_tune_btn.setText("Stop Auto-Tune")
        self.auto_tune_btn.setEnabled(True)
        self._append("Auto-tune started.")
        self._thread.start()

    @QtCore.Slot(float, float)
    def _autotune_finished(self, best_freq: float, best_amp: float) -> None:
        if best_freq > 0:
            self.new_freq.setValue(best_freq)
        self._append(f"Auto-tune completed. Best freq={best_freq:.5f} Hz, monitor_peak={best_amp:.6f} V")

    @QtCore.Slot(str)
    def _autotune_failed(self, message: str) -> None:
        QtWidgets.QMessageBox.critical(self, "Auto-Tune Error", message)
        self._append(f"Auto-tune failed: {message}")

    @QtCore.Slot()
    def _cleanup_worker(self) -> None:
        self.auto_tune_btn.setText("Start Auto-Tune")
        self.auto_tune_btn.setEnabled(True)

        if self._worker is not None:
            self._worker.deleteLater()
        if self._thread is not None:
            self._thread.deleteLater()

        self._worker = None
        self._thread = None


def main() -> int:
    app = QtWidgets.QApplication(sys.argv)
    apply_liquid_glass_theme(app)
    window = MainWindow()
    window.show()
    return app.exec()
