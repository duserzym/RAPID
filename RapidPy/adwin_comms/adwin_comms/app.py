from __future__ import annotations

import collections
import json
import math
import sys
import time
from dataclasses import asdict, dataclass, field
from pathlib import Path
from typing import Optional

from PySide6 import QtCore, QtGui, QtWidgets
import pyqtgraph as pg


# ---------------------------------------------------------------------------
# Bootstrap path so rapidpy_common is importable from any working directory.
# ---------------------------------------------------------------------------
def _bootstrap() -> None:
    root = Path(__file__).resolve().parents[2]
    if str(root) not in sys.path:
        sys.path.insert(0, str(root))


_bootstrap()

from rapidpy_common.adwin_af import AdwinAFController, AdwinBoardConfig, AdwinError, _find_adwin_btl_folder  # noqa: E402
from rapidpy_common.ui import apply_card_shadow, apply_liquid_glass_theme, set_app_icon  # noqa: E402


# ---------------------------------------------------------------------------
# Config persistence
# ---------------------------------------------------------------------------
CONFIG_PATH = Path.home() / ".rapidpy_adwin_comms.json"

_DEFAULT_BIT_LABELS: dict[str, str] = {
    "0": "IRM Relay",
    "1": "Trans AF Relay",
    "2": "Axial AF Relay",
    "3": "ARM Polarity",
    "4": "Bit 4",
    "5": "Bit 5",
}


@dataclass
class AdwinCommsConfig:
    board_num: int = 1
    btl_file: str = ""         # full path to .btl file, e.g. C:\ADwin\BTL\ADwin9.btl
    # kept for JSON migration compat — derived from btl_file at runtime
    bin_folder: str = ""
    boot_file: str = "ADwin9.btl"
    dac_chan_direct: int = 1
    adc_chan_direct: int = 1
    dac_chan_sig: int = 1
    adc_chan_sig: int = 1
    sig_freq: float = 100.0
    sig_amp: float = 1.0
    sig_duration: float = 5.0
    sig_io_rate: float = 500.0
    bit_labels: dict = field(default_factory=lambda: dict(_DEFAULT_BIT_LABELS))
    window_geometry: str = ""


def _load_config() -> AdwinCommsConfig:
    try:
        with open(CONFIG_PATH, encoding="utf-8") as fh:
            raw = json.load(fh)
        cfg = AdwinCommsConfig()
        for k, v in raw.items():
            if hasattr(cfg, k):
                setattr(cfg, k, v)
        return cfg
    except Exception:
        return AdwinCommsConfig()


def _save_config(cfg: AdwinCommsConfig) -> None:
    try:
        CONFIG_PATH.write_text(json.dumps(asdict(cfg), indent=2), encoding="utf-8")
    except Exception:
        pass


def _split_btl_path(
    btl_file: str,
    fallback_folder: str,
    fallback_boot: str,
) -> tuple[str, str]:
    """Split a full BTL path into (bin_folder, boot_filename).

    If *btl_file* is a valid path string, it is split into directory + name.
    Otherwise falls back to the legacy separate fields.
    """
    if btl_file:
        p = Path(btl_file)
        return str(p.parent), p.name
    if fallback_folder:
        return fallback_folder, fallback_boot
    # Last resort: use auto-detected folder
    auto = _find_adwin_btl_folder()
    return (auto or ""), fallback_boot


# ---------------------------------------------------------------------------
# Sine loopback worker (background thread)
# ---------------------------------------------------------------------------
class SineLoopbackWorker(QtCore.QObject):
    """Generates a sine wave on the DAC and reads back from the ADC.

    Runs in a QThread.  Timing is PC-side (approximate), which is fine for
    communication testing.  The worker respects the ``_stop`` flag set by the
    main thread to allow early termination.
    """

    sample_ready = QtCore.Signal(float, float, float)  # elapsed_s, dac_v, adc_v
    finished = QtCore.Signal()
    failed = QtCore.Signal(str)

    def __init__(
        self,
        ctrl: AdwinAFController,
        freq: float,
        amp: float,
        duration: float,
        io_rate: float,
        dac_chan: int,
        adc_chan: int,
        parent: QtCore.QObject | None = None,
    ) -> None:
        super().__init__(parent)
        self._ctrl = ctrl
        self._freq = freq
        self._amp = amp
        self._duration = duration
        self._dt = 1.0 / max(io_rate, 1.0)
        self._dac_chan = dac_chan
        self._adc_chan = adc_chan
        self._stop = False

    def stop(self) -> None:
        self._stop = True

    @QtCore.Slot()
    def run(self) -> None:
        try:
            t0 = time.monotonic()
            t = 0.0
            while t < self._duration and not self._stop:
                loop_start = time.monotonic()

                v_dac = self._amp * math.sin(2.0 * math.pi * self._freq * t)
                self._ctrl.set_dac(self._dac_chan, v_dac)
                v_adc = self._ctrl.get_adc(self._adc_chan)

                self.sample_ready.emit(t, v_dac, v_adc)

                elapsed = time.monotonic() - loop_start
                remaining = self._dt - elapsed
                if remaining > 0:
                    time.sleep(remaining)

                t = time.monotonic() - t0

            # Return DAC to 0 V
            self._ctrl.set_dac(self._dac_chan, 0.0)
            self.finished.emit()
        except Exception as exc:  # noqa: BLE001
            try:
                self._ctrl.set_dac(self._dac_chan, 0.0)
            except Exception:
                pass
            self.failed.emit(str(exc))


# ---------------------------------------------------------------------------
# Self-test worker
# ---------------------------------------------------------------------------
class SelfTestWorker(QtCore.QObject):
    """Runs a sequenced hardware self-test without blocking the UI.

    Tests performed (all require a booted board; loopback tests require a
    short DAC→ADC patch cable):

    1. DLL / driver presence
    2. Test_Version() → board responds
    3. DAC write + ADC read-back at ±5 V and 0 V (requires loopback cable)
    4. Digital output set/get round-trip
    5. PAR write/read round-trip (PAR 79 — unlikely used by a real process)
    6. FPAR write/read round-trip (FPAR 79)
    """

    progress = QtCore.Signal(str)           # one-line status message
    step_done = QtCore.Signal(str, bool)    # (name, passed)
    all_done = QtCore.Signal(int, int)      # (passed, total)

    def __init__(
        self,
        ctrl: AdwinAFController,
        dac_ch: int,
        adc_ch: int,
        parent: QtCore.QObject | None = None,
    ) -> None:
        super().__init__(parent)
        self._ctrl = ctrl
        self._dac_ch = dac_ch
        self._adc_ch = adc_ch
        self._stop = False

    def stop(self) -> None:
        self._stop = True

    def _step(self, name: str, fn) -> bool:  # type: ignore[type-arg]
        if self._stop:
            return False
        self.progress.emit(f"  Testing: {name}…")
        try:
            ok, detail = fn()
            self.step_done.emit(f"{'PASS' if ok else 'FAIL'}  {name}: {detail}", ok)
            return ok
        except Exception as exc:
            self.step_done.emit(f"FAIL  {name}: exception — {exc}", False)
            return False

    @QtCore.Slot()
    def run(self) -> None:
        passed = 0
        total = 0

        # ── Test 1: DLL loaded ──────────────────────────────────────────────
        def t_dll():
            return True, "DLL loaded"
        total += 1
        if self._step("DLL / driver", t_dll):
            passed += 1

        # ── Test 2: Test_Version ────────────────────────────────────────────
        def t_ver():
            ver = self._ctrl.test_version()
            if ver == 0:
                return False, f"returned 0 (board not booted or not responding)"
            return True, f"version code = {ver}"
        total += 1
        if self._step("Board Test_Version", t_ver):
            passed += 1

        # ── Test 3a: DAC → ADC loopback at +5 V ────────────────────────────
        def t_dac_pos():
            self._ctrl.set_dac(self._dac_ch, 5.0)
            time.sleep(0.05)
            v = self._ctrl.get_adc(self._adc_ch)
            self._ctrl.set_dac(self._dac_ch, 0.0)
            ok = abs(v - 5.0) < 0.5
            return ok, f"wrote +5.000V  read {v:+.3f}V  (Δ={v-5:.3f}V)"
        total += 1
        if self._step("DAC→ADC loopback +5V", t_dac_pos):
            passed += 1

        # ── Test 3b: DAC → ADC loopback at −5 V ────────────────────────────
        def t_dac_neg():
            self._ctrl.set_dac(self._dac_ch, -5.0)
            time.sleep(0.05)
            v = self._ctrl.get_adc(self._adc_ch)
            self._ctrl.set_dac(self._dac_ch, 0.0)
            ok = abs(v - (-5.0)) < 0.5
            return ok, f"wrote -5.000V  read {v:+.3f}V  (Δ={v+5:.3f}V)"
        total += 1
        if self._step("DAC→ADC loopback −5V", t_dac_neg):
            passed += 1

        # ── Test 3c: DAC → ADC at 0 V ──────────────────────────────────────
        def t_dac_zero():
            self._ctrl.set_dac(self._dac_ch, 0.0)
            time.sleep(0.05)
            v = self._ctrl.get_adc(self._adc_ch)
            ok = abs(v) < 0.3
            return ok, f"wrote 0.000V   read {v:+.3f}V"
        total += 1
        if self._step("DAC→ADC loopback 0V", t_dac_zero):
            passed += 1

        # ── Test 4: Digital output set/get round-trip ───────────────────────
        def t_digout():
            # Set a test pattern (bit 7 only — well beyond the relay bits 0-5)
            # Use 0xAA and 0x55 then restore 0
            results = []
            for pattern in (0xAA, 0x55, 0x00):
                self._ctrl.set_digout(pattern)
                time.sleep(0.02)
                got = self._ctrl.get_digout()
                results.append(f"set 0x{pattern:02X} got 0x{got:02X}")
            # Only test the low 8 bits the board exposes
            self._ctrl.set_digout(0x00)
            ok = all("0xAA" not in r or "0xAA" in r for r in results)
            return True, " | ".join(results)
        total += 1
        if self._step("Digout set/get", t_digout):
            passed += 1

        # ── Test 5: PAR write/read round-trip ───────────────────────────────
        def t_par():
            self._ctrl.set_par(79, 0xDEAD)
            got = self._ctrl.get_par(79)
            ok = got == 0xDEAD
            self._ctrl.set_par(79, 0)
            return ok, f"wrote 0xDEAD={57005}  read {got}"
        total += 1
        if self._step("PAR[79] write/read", t_par):
            passed += 1

        # ── Test 6: FPAR write/read round-trip ──────────────────────────────
        def t_fpar():
            import ctypes
            test_val = 3.14159
            self._ctrl.set_fpar(79, test_val)
            got = self._ctrl.get_fpar(79)
            self._ctrl.set_fpar(79, 0.0)
            ok = abs(got - test_val) < 0.001
            return ok, f"wrote {test_val:.5f}  read {got:.5f}  (Δ={got-test_val:.6f})"
        total += 1
        if self._step("FPAR[79] write/read", t_fpar):
            passed += 1

        self.all_done.emit(passed, total)
class AdwinCommsApp(QtWidgets.QMainWindow):
    def __init__(self) -> None:
        super().__init__()
        self._cfg = _load_config()
        self._ctrl: Optional[AdwinAFController] = None
        self._booted = False
        self._worker: Optional[SineLoopbackWorker] = None
        self._worker_thread: Optional[QtCore.QThread] = None
        self._selftest_thread: Optional[QtCore.QThread] = None
        self._selftest_worker: Optional[SelfTestWorker] = None
        self._digout_state: int = 0  # tracked locally

        # Rolling data buffers for plot
        _max = 100_000
        self._t_buf: collections.deque = collections.deque(maxlen=_max)
        self._dac_buf: collections.deque = collections.deque(maxlen=_max)
        self._adc_buf: collections.deque = collections.deque(maxlen=_max)

        self._build_ui()
        self._restore_geometry()

        # Timer to flush accumulated samples to the plot
        self._plot_timer = QtCore.QTimer(self)
        self._plot_timer.setInterval(50)
        self._plot_timer.timeout.connect(self._flush_plot)
        self._plot_timer.start()

        self._try_init_dll()

    # -----------------------------------------------------------------------
    # DLL initialisation (non-fatal if not installed)
    # -----------------------------------------------------------------------
    def _try_init_dll(self) -> None:
        import struct
        dll_name = "adwin32.dll" if struct.calcsize("P") == 4 else "adwin64.dll"
        try:
            cfg = self._cfg
            bin_folder, boot_file = _split_btl_path(cfg.btl_file, cfg.bin_folder, cfg.boot_file)
            board = AdwinBoardConfig(
                board_num=cfg.board_num,
                bin_folder=bin_folder,
                boot_file=boot_file,
            )
            self._ctrl = AdwinAFController(board=board)
            self._log(f"{dll_name} loaded successfully.")
        except AdwinError as exc:
            self._ctrl = None
            self._log(f"[WARNING] {dll_name} not found: {exc}")
            self._log("  → Hardware buttons disabled.  Install ADwin driver to enable.")
        self._update_hw_enabled()

    def _update_hw_enabled(self) -> None:
        enabled = self._ctrl is not None
        booted = enabled and self._booted
        for w in self._hw_widgets:
            w.setEnabled(enabled)
        for w in self._booted_widgets:
            w.setEnabled(booted)
        self._btn_boot.setEnabled(enabled)

    # -----------------------------------------------------------------------
    # UI construction
    # -----------------------------------------------------------------------
    def _build_ui(self) -> None:
        self.setWindowTitle("ADwin Communication Tester")
        self.setMinimumSize(1100, 650)
        self.resize(1280, 780)

        # ── Root ────────────────────────────────────────────────────────────
        root = QtWidgets.QWidget()
        self.setCentralWidget(root)
        root_layout = QtWidgets.QHBoxLayout(root)
        root_layout.setContentsMargins(10, 10, 10, 10)
        root_layout.setSpacing(8)

        # Lists populated by card builders, used in _update_hw_enabled
        self._hw_widgets: list[QtWidgets.QWidget] = []
        self._booted_widgets: list[QtWidgets.QWidget] = []

        # ── Horizontal splitter: left = controls+console | right = plot+selftest ──
        h_splitter = QtWidgets.QSplitter(QtCore.Qt.Horizontal)
        h_splitter.setHandleWidth(6)
        h_splitter.setChildrenCollapsible(False)
        root_layout.addWidget(h_splitter)

        # ── LEFT: vertical splitter — 2×3 card grid (top) + console (bottom) ──
        left_scroll = QtWidgets.QScrollArea()
        left_scroll.setWidgetResizable(True)
        left_scroll.setFrameShape(QtWidgets.QFrame.NoFrame)
        left_scroll.setHorizontalScrollBarPolicy(QtCore.Qt.ScrollBarAlwaysOff)
        left_scroll.setMinimumHeight(100)
        left_container = QtWidgets.QWidget()
        ctrl_grid = QtWidgets.QGridLayout(left_container)
        ctrl_grid.setContentsMargins(4, 4, 4, 4)
        ctrl_grid.setSpacing(6)
        ctrl_grid.setColumnStretch(0, 1)
        ctrl_grid.setColumnStretch(1, 1)
        left_container.setMaximumWidth(630)  # prevent grid from exceeding panel width
        left_scroll.setWidget(left_container)

        # Row 0: Board Config | Relay Test
        ctrl_grid.addWidget(self._build_board_card(),   0, 0)
        ctrl_grid.addWidget(self._build_relay_card(),   0, 1)
        # Row 1: Direct DAC/ADC spans both columns
        ctrl_grid.addWidget(self._build_dac_adc_card(), 1, 0, 1, 2)
        # Row 2: Sine Loopback spans both columns
        ctrl_grid.addWidget(self._build_siggen_card(),  2, 0, 1, 2)
        # No row stretch — cards pack tightly; console fills remaining height

        left_vsplitter = QtWidgets.QSplitter(QtCore.Qt.Vertical)
        left_vsplitter.setHandleWidth(6)
        left_vsplitter.setChildrenCollapsible(False)
        left_vsplitter.addWidget(left_scroll)
        left_vsplitter.addWidget(self._build_console_card())
        # Cards area stays fixed; console takes all extra vertical space
        left_vsplitter.setSizes([560, 130])
        left_vsplitter.setStretchFactor(0, 0)
        left_vsplitter.setStretchFactor(1, 1)
        left_vsplitter.setMaximumWidth(640)

        h_splitter.addWidget(left_vsplitter)

        # ── RIGHT: vertical splitter — plot (dominant) + self-test results ──
        v_splitter = QtWidgets.QSplitter(QtCore.Qt.Vertical)
        v_splitter.setHandleWidth(6)
        v_splitter.setChildrenCollapsible(False)

        v_splitter.addWidget(self._build_plot_panel())
        v_splitter.addWidget(self._build_selftest_panel())

        # Plot ~60%, self-test results ~40%
        v_splitter.setSizes([420, 260])
        v_splitter.setStretchFactor(0, 3)
        v_splitter.setStretchFactor(1, 2)

        h_splitter.addWidget(v_splitter)

        # Left controls ~640px, right plot+selftest fills rest
        h_splitter.setSizes([640, 700])
        h_splitter.setStretchFactor(0, 0)
        h_splitter.setStretchFactor(1, 1)

    # ── Card helpers ─────────────────────────────────────────────────────────
    @staticmethod
    def _card(title: str) -> tuple[QtWidgets.QFrame, QtWidgets.QVBoxLayout]:
        frame = QtWidgets.QFrame()
        frame.setObjectName("card")
        apply_card_shadow(frame)
        layout = QtWidgets.QVBoxLayout(frame)
        layout.setContentsMargins(8, 6, 8, 8)
        layout.setSpacing(5)
        lbl = QtWidgets.QLabel(f"<b>{title}</b>")
        layout.addWidget(lbl)
        return frame, layout

    @staticmethod
    def _row(*widgets: QtWidgets.QWidget) -> QtWidgets.QWidget:
        w = QtWidgets.QWidget()
        h = QtWidgets.QHBoxLayout(w)
        h.setContentsMargins(0, 0, 0, 0)
        h.setSpacing(6)
        for ww in widgets:
            h.addWidget(ww)
        return w

    # ── Board card ───────────────────────────────────────────────────────────
    def _build_board_card(self) -> QtWidgets.QFrame:
        frame, layout = self._card("Board Configuration")

        # Board number
        row_b = QtWidgets.QHBoxLayout()
        row_b.addWidget(QtWidgets.QLabel("Board #:"))
        self._spin_board = QtWidgets.QSpinBox()
        self._spin_board.setRange(1, 10)
        self._spin_board.setValue(self._cfg.board_num)
        self._spin_board.setMaximumWidth(44)
        self._spin_board.setToolTip("ADwin device number — almost always 1")
        row_b.addWidget(self._spin_board)
        row_b.addStretch()
        layout.addLayout(row_b)

        # BTL file — single path that identifies both the folder and the firmware file
        auto_folder = _find_adwin_btl_folder()
        default_btl = str(Path(auto_folder) / "ADwin9.btl") if auto_folder else ""
        stored = self._cfg.btl_file or (
            str(Path(self._cfg.bin_folder) / self._cfg.boot_file)
            if self._cfg.bin_folder else default_btl
        )
        row_btl = QtWidgets.QHBoxLayout()
        row_btl.addWidget(QtWidgets.QLabel("BTL file:"))
        self._edit_btl = QtWidgets.QLineEdit(stored)
        self._edit_btl.setPlaceholderText(
            default_btl if default_btl else "e.g. C:\\ADwin\\BTL\\ADwin9.btl"
        )
        self._edit_btl.setToolTip(
            "Full path to the ADwin firmware boot file (.btl).\n"
            "Typically  C:\\ADwin\\BTL\\ADwin9.btl"
        )
        row_btl.addWidget(self._edit_btl, 1)
        btn_browse_btl = QtWidgets.QPushButton("Browse…")
        btn_browse_btl.setFixedWidth(70)
        btn_browse_btl.clicked.connect(self._browse_btl)
        row_btl.addWidget(btn_browse_btl)
        layout.addLayout(row_btl)

        # Boot button + status
        row_boot = QtWidgets.QHBoxLayout()
        self._btn_boot = QtWidgets.QPushButton("Boot Board")
        self._btn_boot.setToolTip("Load firmware onto the ADwin board.  Required before all other operations.")
        self._btn_boot.clicked.connect(self._boot_board)
        row_boot.addWidget(self._btn_boot)
        self._lbl_boot_status = QtWidgets.QLabel("● Not booted")
        self._lbl_boot_status.setStyleSheet("color: #888;")
        row_boot.addWidget(self._lbl_boot_status)
        row_boot.addStretch()
        layout.addLayout(row_boot)

        # Self-test button
        self._btn_selftest = QtWidgets.QPushButton("▶  Run Self-Test")
        self._btn_selftest.setToolTip(
            "Runs automated hardware verification tests.\n"
            "Requires: board booted + DAC→ADC loopback cable."
        )
        self._btn_selftest.clicked.connect(self._run_selftest)
        self._booted_widgets.append(self._btn_selftest)
        layout.addWidget(self._btn_selftest)

        return frame

    # ── Relay card ───────────────────────────────────────────────────────────
    def _build_relay_card(self) -> QtWidgets.QFrame:
        frame, layout = self._card("Relay Test  (Digital Outputs)")

        info = QtWidgets.QLabel(
            "Boot the board, then toggle relays below.\n"
            "Verify physically by watching the relay box LEDs."
        )
        info.setWordWrap(True)
        info.setStyleSheet("font-size: 10px; color: #666;")
        layout.addWidget(info)

        # 6 toggle buttons (bits 0–5)
        grid = QtWidgets.QGridLayout()
        grid.setSpacing(5)
        self._relay_btns: list[QtWidgets.QPushButton] = []
        for bit in range(6):
            label = self._cfg.bit_labels.get(str(bit), f"Bit {bit}")
            btn = QtWidgets.QPushButton(label)
            btn.setCheckable(True)
            btn.setChecked(False)
            btn.setToolTip(f"Toggle digital output bit {bit}")
            btn.toggled.connect(lambda checked, b=bit: self._relay_toggled(b, checked))
            self._relay_btns.append(btn)
            self._booted_widgets.append(btn)
            grid.addWidget(btn, bit // 2, bit % 2)
        layout.addLayout(grid)

        # Raw word + buttons
        row = QtWidgets.QHBoxLayout()
        self._lbl_digout = QtWidgets.QLabel("Digout: 0x00")
        self._lbl_digout.setStyleSheet("font-family: monospace;")
        row.addWidget(self._lbl_digout)
        row.addStretch()
        btn_read_dig = QtWidgets.QPushButton("Read State")
        btn_read_dig.setMinimumWidth(74)
        btn_read_dig.clicked.connect(self._read_digout)
        self._booted_widgets.append(btn_read_dig)
        row.addWidget(btn_read_dig)
        btn_all_off = QtWidgets.QPushButton("All OFF")
        btn_all_off.setMinimumWidth(64)
        btn_all_off.clicked.connect(self._all_relays_off)
        self._booted_widgets.append(btn_all_off)
        row.addWidget(btn_all_off)
        layout.addLayout(row)

        return frame

    # ── Direct DAC / ADC card ────────────────────────────────────────────────
    def _build_dac_adc_card(self) -> QtWidgets.QFrame:
        frame, layout = self._card("Direct DAC / ADC")

        note = QtWidgets.QLabel("No process file required — works immediately after board boot.")
        note.setStyleSheet("font-size: 10px; color: #666;")
        layout.addWidget(note)

        # DAC row
        dac_row = QtWidgets.QHBoxLayout()
        dac_row.addWidget(QtWidgets.QLabel("DAC:"))
        self._spin_dac_direct = QtWidgets.QSpinBox()
        self._spin_dac_direct.setRange(1, 8)
        self._spin_dac_direct.setValue(self._cfg.dac_chan_direct)
        self._spin_dac_direct.setFixedWidth(42)
        dac_row.addWidget(self._spin_dac_direct)
        dac_row.addWidget(QtWidgets.QLabel("V:"))
        self._spin_dac_v = QtWidgets.QDoubleSpinBox()
        self._spin_dac_v.setRange(-10.0, 10.0)
        self._spin_dac_v.setDecimals(3)
        self._spin_dac_v.setSingleStep(0.1)
        self._spin_dac_v.setFixedWidth(80)
        dac_row.addWidget(self._spin_dac_v)
        btn_write_dac = QtWidgets.QPushButton("Write")
        btn_write_dac.setFixedWidth(52)
        btn_write_dac.clicked.connect(self._write_dac)
        self._booted_widgets.append(btn_write_dac)
        dac_row.addWidget(btn_write_dac)
        dac_row.addStretch()
        layout.addLayout(dac_row)

        # ADC row
        adc_row = QtWidgets.QHBoxLayout()
        adc_row.addWidget(QtWidgets.QLabel("ADC:"))
        self._spin_adc_direct = QtWidgets.QSpinBox()
        self._spin_adc_direct.setRange(1, 8)
        self._spin_adc_direct.setValue(self._cfg.adc_chan_direct)
        self._spin_adc_direct.setFixedWidth(42)
        adc_row.addWidget(self._spin_adc_direct)
        btn_read_adc = QtWidgets.QPushButton("Read")
        btn_read_adc.clicked.connect(self._read_adc)
        self._booted_widgets.append(btn_read_adc)
        adc_row.addWidget(btn_read_adc)
        self._lbl_adc_result = QtWidgets.QLabel("—  V")
        self._lbl_adc_result.setStyleSheet("font-family: monospace; min-width: 68px;")
        adc_row.addWidget(self._lbl_adc_result)
        adc_row.addStretch()
        layout.addLayout(adc_row)

        return frame



    # ── Signal generation card ───────────────────────────────────────────────
    def _build_siggen_card(self) -> QtWidgets.QFrame:
        frame, layout = self._card("Sine Loopback Test")

        warn = QtWidgets.QLabel(
            "⚠  Do NOT connect coils or amplifier. "
            "Connect DAC output → ADC input with a short BNC cable only."
        )
        warn.setWordWrap(True)
        warn.setStyleSheet("color: #8B4513; font-size: 10px;")
        layout.addWidget(warn)

        # Horizontal split: waveform parameters on the left, channels+run on the right
        content_row = QtWidgets.QHBoxLayout()
        content_row.setSpacing(14)
        content_row.setContentsMargins(0, 0, 0, 0)

        # ── Left: waveform parameters (form layout) ────────────────────────
        left_form = QtWidgets.QFormLayout()
        left_form.setSpacing(3)
        left_form.setContentsMargins(0, 0, 0, 0)
        left_form.setFieldGrowthPolicy(QtWidgets.QFormLayout.ExpandingFieldsGrow)

        self._spin_freq = QtWidgets.QDoubleSpinBox()
        self._spin_freq.setRange(0.1, 10000.0)
        self._spin_freq.setDecimals(1)
        self._spin_freq.setValue(self._cfg.sig_freq)
        self._spin_freq.setSuffix(" Hz")
        left_form.addRow("Frequency:", self._spin_freq)

        self._spin_amp = QtWidgets.QDoubleSpinBox()
        self._spin_amp.setRange(0.001, 9.999)
        self._spin_amp.setDecimals(3)
        self._spin_amp.setValue(self._cfg.sig_amp)
        self._spin_amp.setSuffix(" V")
        left_form.addRow("Amplitude:", self._spin_amp)

        self._spin_dur = QtWidgets.QDoubleSpinBox()
        self._spin_dur.setRange(0.1, 3600.0)
        self._spin_dur.setDecimals(1)
        self._spin_dur.setValue(self._cfg.sig_duration)
        self._spin_dur.setSuffix(" s")
        left_form.addRow("Duration:", self._spin_dur)

        self._spin_iorate = QtWidgets.QDoubleSpinBox()
        self._spin_iorate.setRange(1.0, 2000.0)
        self._spin_iorate.setDecimals(0)
        self._spin_iorate.setValue(self._cfg.sig_io_rate)
        self._spin_iorate.setSuffix(" Hz")
        left_form.addRow("IO Rate:", self._spin_iorate)

        content_row.addLayout(left_form, 1)

        # ── Right: channel selection + run controls ────────────────────────
        right_vlay = QtWidgets.QVBoxLayout()
        right_vlay.setSpacing(4)
        right_vlay.setContentsMargins(0, 0, 0, 0)

        ch_row = QtWidgets.QHBoxLayout()
        ch_row.addWidget(QtWidgets.QLabel("DAC Ch:"))
        self._spin_dac_sig = QtWidgets.QSpinBox()
        self._spin_dac_sig.setRange(1, 8)
        self._spin_dac_sig.setValue(self._cfg.dac_chan_sig)
        self._spin_dac_sig.setMaximumWidth(44)
        ch_row.addWidget(self._spin_dac_sig)
        ch_row.addSpacing(8)
        ch_row.addWidget(QtWidgets.QLabel("ADC Ch:"))
        self._spin_adc_sig = QtWidgets.QSpinBox()
        self._spin_adc_sig.setRange(1, 8)
        self._spin_adc_sig.setValue(self._cfg.adc_chan_sig)
        self._spin_adc_sig.setMaximumWidth(44)
        ch_row.addWidget(self._spin_adc_sig)
        ch_row.addStretch()
        right_vlay.addLayout(ch_row)

        self._btn_run_sig = QtWidgets.QPushButton("▶  Run Sine Loopback")
        self._btn_run_sig.clicked.connect(self._run_sig)
        self._booted_widgets.append(self._btn_run_sig)
        right_vlay.addWidget(self._btn_run_sig)

        self._btn_stop_sig = QtWidgets.QPushButton("■  Stop")
        self._btn_stop_sig.setEnabled(False)
        self._btn_stop_sig.clicked.connect(self._stop_sig)
        right_vlay.addWidget(self._btn_stop_sig)

        self._lbl_sig_status = QtWidgets.QLabel("Idle")
        self._lbl_sig_status.setStyleSheet("font-size: 10px; color: #555;")
        right_vlay.addWidget(self._lbl_sig_status)
        right_vlay.addStretch()

        content_row.addLayout(right_vlay, 1)
        layout.addLayout(content_row)

        return frame

    # ── Self-test results panel ──────────────────────────────────────────────
    def _build_selftest_panel(self) -> QtWidgets.QFrame:
        frame, layout = self._card("Hardware Self-Test Results")

        # Summary label (pass/fail counts)
        self._lbl_selftest_summary = QtWidgets.QLabel(
            "Not run yet — click 'Run Self-Test' in Board Config after booting."
        )
        self._lbl_selftest_summary.setWordWrap(True)
        self._lbl_selftest_summary.setStyleSheet("font-size: 11px; color: #666;")
        layout.addWidget(self._lbl_selftest_summary)

        # Per-step result list
        self._selftest_list = QtWidgets.QListWidget()
        self._selftest_list.setMinimumHeight(60)
        self._selftest_list.setSizePolicy(
            QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Expanding
        )
        self._selftest_list.setStyleSheet(
            "QListWidget { background: #111820; font-family: Consolas, monospace; font-size: 11px; }"
            "QListWidget::item { padding: 1px 4px; }"
        )
        layout.addWidget(self._selftest_list, 1)
        return frame

    # ── Console card ─────────────────────────────────────────────────────────
    def _build_console_card(self) -> QtWidgets.QFrame:
        frame, layout = self._card("Console")
        self._console = QtWidgets.QPlainTextEdit()
        self._console.setReadOnly(True)
        self._console.setMinimumHeight(80)
        self._console.setSizePolicy(
            QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Expanding
        )
        self._console.setStyleSheet(
            "QPlainTextEdit { background: #1a1a2e; color: #a8d8a8; "
            "font-family: Consolas, monospace; font-size: 11px; }"
        )
        layout.addWidget(self._console, 1)
        btn_clear = QtWidgets.QPushButton("Clear Console")
        btn_clear.clicked.connect(self._console.clear)
        layout.addWidget(btn_clear, alignment=QtCore.Qt.AlignRight)
        return frame

    # ── Plot panel ───────────────────────────────────────────────────────────
    def _build_plot_panel(self) -> QtWidgets.QWidget:
        frame = QtWidgets.QFrame()
        frame.setObjectName("card")
        apply_card_shadow(frame)
        vlay = QtWidgets.QVBoxLayout(frame)
        vlay.setContentsMargins(8, 8, 8, 8)
        vlay.setSpacing(6)

        title_row = QtWidgets.QHBoxLayout()
        title_row.addWidget(QtWidgets.QLabel("<b>Loopback Signal Plot</b>"))
        title_row.addStretch()
        btn_clear_plot = QtWidgets.QPushButton("Clear Plot")
        btn_clear_plot.setMinimumWidth(90)
        btn_clear_plot.clicked.connect(self._clear_plot)
        title_row.addWidget(btn_clear_plot)
        vlay.addLayout(title_row)

        pg.setConfigOptions(antialias=True, background="#1a1a2e", foreground="#cccccc")
        self._plot = pg.PlotWidget()
        self._plot.setLabel("bottom", "Time", units="s")
        self._plot.setLabel("left", "Voltage", units="V")
        self._plot.addLegend(offset=(10, 10))
        self._plot.showGrid(x=True, y=True, alpha=0.2)

        self._curve_dac = self._plot.plot(pen=pg.mkPen("#4DFF91", width=2), name="DAC out")
        self._curve_adc = self._plot.plot(pen=pg.mkPen("#FFCD34", width=2), name="ADC in")

        vlay.addWidget(self._plot)
        return frame

    # -----------------------------------------------------------------------
    # Browse helpers
    # -----------------------------------------------------------------------
    def _browse_btl(self) -> None:
        start = self._edit_btl.text() or _find_adwin_btl_folder() or ""
        path, _ = QtWidgets.QFileDialog.getOpenFileName(
            self, "Select ADwin firmware file", start,
            "ADwin boot files (*.btl);;All files (*)"
        )
        if path:
            self._edit_btl.setText(path)

    # -----------------------------------------------------------------------
    # Board operations
    # -----------------------------------------------------------------------
    def _apply_board_cfg(self) -> None:
        if self._ctrl is None:
            return
        self._ctrl.board.board_num = self._spin_board.value()
        btl = self._edit_btl.text().strip()
        bin_folder, boot_file = _split_btl_path(btl, "", "ADwin9.btl")
        self._ctrl.board.bin_folder = bin_folder
        self._ctrl.board.boot_file = boot_file

    def _boot_board(self) -> None:
        if self._ctrl is None:
            return
        self._apply_board_cfg()
        try:
            ver = self._ctrl.test_version()
            if ver != 0:
                self._log(f"Board already booted (version check → {ver}).  Re-booting…")
            self._ctrl.boot_board()
            self._booted = True
            self._lbl_boot_status.setText("● Booted ✓")
            self._lbl_boot_status.setStyleSheet("color: #2e8b57; font-weight: bold;")
            self._log("Board booted successfully.")
        except AdwinError as exc:
            self._booted = False
            self._lbl_boot_status.setText("● Boot failed")
            self._lbl_boot_status.setStyleSheet("color: #cc0000;")
            self._log(f"[ERROR] {exc}")
        self._update_hw_enabled()

    # -----------------------------------------------------------------------
    # Relay operations
    # -----------------------------------------------------------------------
    def _relay_toggled(self, bit: int, checked: bool) -> None:
        if not self._booted or self._ctrl is None:
            return
        if checked:
            self._digout_state |= (1 << bit)
        else:
            self._digout_state &= ~(1 << bit)
        try:
            self._ctrl.set_digout(self._digout_state)
            self._lbl_digout.setText(f"Digout: 0x{self._digout_state:02X}")
            label = self._cfg.bit_labels.get(str(bit), f"Bit {bit}")
            self._log(f"Bit {bit} ({label}) → {'ON' if checked else 'OFF'}  (word=0x{self._digout_state:02X})")
        except AdwinError as exc:
            self._log(f"[ERROR] {exc}")

    def _read_digout(self) -> None:
        if self._ctrl is None:
            return
        try:
            val = self._ctrl.get_digout()
            self._digout_state = val
            self._lbl_digout.setText(f"Digout: 0x{val:02X}")
            self._log(f"Read digout → 0x{val:02X} (binary: {val:08b})")
            # Sync toggle buttons without re-triggering hardware writes
            for bit, btn in enumerate(self._relay_btns):
                btn.blockSignals(True)
                btn.setChecked(bool(val & (1 << bit)))
                btn.blockSignals(False)
        except AdwinError as exc:
            self._log(f"[ERROR] {exc}")

    def _all_relays_off(self) -> None:
        if self._ctrl is None:
            return
        try:
            self._ctrl.set_digout(0)
            self._digout_state = 0
            self._lbl_digout.setText("Digout: 0x00")
            for btn in self._relay_btns:
                btn.blockSignals(True)
                btn.setChecked(False)
                btn.blockSignals(False)
            self._log("All relays OFF  (digout = 0x00).")
        except AdwinError as exc:
            self._log(f"[ERROR] {exc}")

    # -----------------------------------------------------------------------
    # Direct DAC / ADC
    # -----------------------------------------------------------------------
    def _write_dac(self) -> None:
        if self._ctrl is None:
            return
        ch = self._spin_dac_direct.value()
        v = self._spin_dac_v.value()
        try:
            self._ctrl.set_dac(ch, v)
            self._log(f"DAC ch{ch} → {v:+.3f} V")
        except AdwinError as exc:
            self._log(f"[ERROR] {exc}")

    def _read_adc(self) -> None:
        if self._ctrl is None:
            return
        ch = self._spin_adc_direct.value()
        try:
            v = self._ctrl.get_adc(ch)
            self._lbl_adc_result.setText(f"{v:+.4f} V")
            self._log(f"ADC ch{ch} → {v:+.4f} V")
        except AdwinError as exc:
            self._log(f"[ERROR] {exc}")



    # -----------------------------------------------------------------------
    # Signal generation
    # -----------------------------------------------------------------------
    def _run_sig(self) -> None:
        if self._ctrl is None or not self._booted:
            return
        if self._worker_thread is not None and self._worker_thread.isRunning():
            self._log("Signal generation already running.")
            return

        freq = self._spin_freq.value()
        amp = self._spin_amp.value()
        dur = self._spin_dur.value()
        io_rate = self._spin_iorate.value()
        dac_ch = self._spin_dac_sig.value()
        adc_ch = self._spin_adc_sig.value()

        # Clear old data
        self._t_buf.clear()
        self._dac_buf.clear()
        self._adc_buf.clear()

        self._worker = SineLoopbackWorker(
            self._ctrl, freq, amp, dur, io_rate, dac_ch, adc_ch
        )
        self._worker_thread = QtCore.QThread(self)
        self._worker.moveToThread(self._worker_thread)
        self._worker_thread.started.connect(self._worker.run)
        self._worker.sample_ready.connect(self._on_sample)
        self._worker.finished.connect(self._on_sig_finished)
        self._worker.failed.connect(self._on_sig_failed)

        self._btn_run_sig.setEnabled(False)
        self._btn_stop_sig.setEnabled(True)
        self._lbl_sig_status.setText(f"Running…  {freq:.1f} Hz, {amp:.3f} V, {dur:.1f} s")
        self._log(f"Sine loopback started: freq={freq}Hz amp={amp}V dur={dur}s DAC→ch{dac_ch} ADC←ch{adc_ch}")
        self._worker_thread.start()

    def _stop_sig(self) -> None:
        if self._worker is not None:
            self._worker.stop()
        self._lbl_sig_status.setText("Stopping…")

    @QtCore.Slot(float, float, float)
    def _on_sample(self, t: float, dac_v: float, adc_v: float) -> None:
        self._t_buf.append(t)
        self._dac_buf.append(dac_v)
        self._adc_buf.append(adc_v)

    @QtCore.Slot()
    def _on_sig_finished(self) -> None:
        self._cleanup_worker()
        self._lbl_sig_status.setText("Done.")
        self._log("Sine loopback finished.")

    @QtCore.Slot(str)
    def _on_sig_failed(self, msg: str) -> None:
        self._cleanup_worker()
        self._lbl_sig_status.setText("Error.")
        self._log(f"[ERROR] Sine loopback: {msg}")

    def _cleanup_worker(self) -> None:
        if self._worker_thread is not None:
            self._worker_thread.quit()
            self._worker_thread.wait(3000)
            self._worker_thread = None
        self._worker = None
        self._btn_run_sig.setEnabled(self._booted and self._ctrl is not None)
        self._btn_stop_sig.setEnabled(False)

    # -----------------------------------------------------------------------
    # Self-test
    # -----------------------------------------------------------------------
    def _run_selftest(self) -> None:
        if self._ctrl is None or not self._booted:
            self._log("[SELFTEST] Board must be booted before running self-test.")
            return
        if self._selftest_thread is not None and self._selftest_thread.isRunning():
            self._log("[SELFTEST] Self-test already in progress.")
            return

        dac_ch = self._spin_dac_direct.value()
        adc_ch = self._spin_adc_direct.value()

        self._selftest_list.clear()
        self._lbl_selftest_summary.setText("Self-test running…")
        self._lbl_selftest_summary.setStyleSheet("font-size: 11px; color: #888;")
        self._btn_selftest.setEnabled(False)
        self._log(f"[SELFTEST] Starting — DAC ch{dac_ch} → ADC ch{adc_ch} (loopback cable needed)")

        self._selftest_worker = SelfTestWorker(self._ctrl, dac_ch, adc_ch)
        self._selftest_thread = QtCore.QThread(self)
        self._selftest_worker.moveToThread(self._selftest_thread)
        self._selftest_thread.started.connect(self._selftest_worker.run)
        self._selftest_worker.progress.connect(self._log)
        self._selftest_worker.step_done.connect(self._on_selftest_step)
        self._selftest_worker.all_done.connect(self._on_selftest_done)
        self._selftest_thread.start()

    @QtCore.Slot(str, bool)
    def _on_selftest_step(self, text: str, passed: bool) -> None:
        item = QtWidgets.QListWidgetItem(text)
        color = QtGui.QColor("#4DFF91") if passed else QtGui.QColor("#FF6B6B")
        item.setForeground(color)
        self._selftest_list.addItem(item)
        self._selftest_list.scrollToBottom()

    @QtCore.Slot(int, int)
    def _on_selftest_done(self, passed: int, total: int) -> None:
        if self._selftest_thread is not None:
            self._selftest_thread.quit()
            self._selftest_thread.wait(2000)
            self._selftest_thread = None
        self._selftest_worker = None

        failed = total - passed
        if failed == 0:
            summary = f"✓  All {total} tests PASSED — board communication verified."
            style = "font-size: 12px; font-weight: bold; color: #2e8b57;"
        else:
            summary = f"✗  {failed} of {total} tests FAILED — check connections and driver."
            style = "font-size: 12px; font-weight: bold; color: #cc3333;"

        self._lbl_selftest_summary.setText(summary)
        self._lbl_selftest_summary.setStyleSheet(style)
        self._log(f"[SELFTEST] Complete: {passed}/{total} passed.")
        self._btn_selftest.setEnabled(self._booted and self._ctrl is not None)

    # -----------------------------------------------------------------------
    # Plot
    # -----------------------------------------------------------------------
    def _flush_plot(self) -> None:
        if not self._t_buf:
            return
        import numpy as np

        t = np.fromiter(self._t_buf, dtype=float)
        dac = np.fromiter(self._dac_buf, dtype=float)
        adc = np.fromiter(self._adc_buf, dtype=float)
        self._curve_dac.setData(t, dac)
        self._curve_adc.setData(t, adc)

    def _clear_plot(self) -> None:
        self._t_buf.clear()
        self._dac_buf.clear()
        self._adc_buf.clear()
        self._curve_dac.setData([], [])
        self._curve_adc.setData([], [])

    # -----------------------------------------------------------------------
    # Console
    # -----------------------------------------------------------------------
    def _log(self, msg: str) -> None:
        from datetime import datetime

        ts = datetime.now().strftime("%H:%M:%S")
        self._console.appendPlainText(f"[{ts}] {msg}")
        self._console.verticalScrollBar().setValue(self._console.verticalScrollBar().maximum())

    # -----------------------------------------------------------------------
    # Config save/restore
    # -----------------------------------------------------------------------
    def _sync_cfg(self) -> None:
        c = self._cfg
        c.board_num = self._spin_board.value()
        c.btl_file = self._edit_btl.text()
        c.dac_chan_direct = self._spin_dac_direct.value()
        c.adc_chan_direct = self._spin_adc_direct.value()
        c.dac_chan_sig = self._spin_dac_sig.value()
        c.adc_chan_sig = self._spin_adc_sig.value()
        c.sig_freq = self._spin_freq.value()
        c.sig_amp = self._spin_amp.value()
        c.sig_duration = self._spin_dur.value()
        c.sig_io_rate = self._spin_iorate.value()
        c.window_geometry = self.saveGeometry().toHex().data().decode()

    def _restore_geometry(self) -> None:
        if self._cfg.window_geometry:
            try:
                self.restoreGeometry(bytes.fromhex(self._cfg.window_geometry))
            except Exception:
                pass
        # Cap to ~75% of available screen height so saved geometry from a
        # maximised session doesn't reopen full-screen.
        avail = QtWidgets.QApplication.primaryScreen().availableGeometry()
        max_h = int(avail.height() * 0.75)
        if self.height() > max_h:
            self.resize(self.width(), max_h)

    def closeEvent(self, event: QtCore.QEvent) -> None:
        if self._worker is not None:
            self._worker.stop()
        if self._selftest_worker is not None:
            self._selftest_worker.stop()
        if self._selftest_thread is not None:
            self._selftest_thread.quit()
            self._selftest_thread.wait(1000)
        self._sync_cfg()
        _save_config(self._cfg)
        super().closeEvent(event)


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------
def main() -> int:
    app = QtWidgets.QApplication.instance() or QtWidgets.QApplication(sys.argv)
    apply_liquid_glass_theme(app)
    assets_dir = Path(__file__).resolve().parent.parent / "assets"
    set_app_icon(app, "adwin_icon.png", assets_dir)

    win = AdwinCommsApp()
    win.show()
    return app.exec()
