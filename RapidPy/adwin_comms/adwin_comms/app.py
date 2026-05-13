from __future__ import annotations

import collections
import csv
import ctypes
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

from rapidpy_common.adwin_af import (  # noqa: E402
    AdwinAFController,
    AdwinBoardConfig,
    AdwinDenseCaptureRequest,
    AdwinDenseCaptureResult,
    AdwinError,
    _find_adwin_btl_folder,
    _find_adwin_dll,
    find_btl_files,
)
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
    sig_freq: float = 5.0
    sig_amp: float = 1.0
    sig_duration: float = 5.0
    sig_io_rate: float = 1000.0
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
        # Migrate the original placeholder sine defaults to values that are
        # visually meaningful on the loopback plot. VB6 assumes IORate/Freq
        # sets points-per-cycle; 100 Hz @ 500 Hz was only 5 samples/cycle.
        if raw.get("sig_freq") == 100.0 and raw.get("sig_io_rate") == 500.0:
            cfg.sig_freq = 5.0
            cfg.sig_io_rate = 1000.0
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


def _default_process_file() -> str:
    candidate = Path(__file__).resolve().parents[3] / "VB6" / "ADwin" / "sineout.T91"
    return str(candidate) if candidate.is_file() else ""


# ---------------------------------------------------------------------------
# Sine loopback worker (background thread)
# ---------------------------------------------------------------------------
class SineLoopbackWorker(QtCore.QObject):
    """Runs a board-timed dense sine capture using the legacy ADwin process."""

    capture_ready = QtCore.Signal(object)
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
        self._io_rate = max(io_rate, 1.0)
        self._dac_chan = dac_chan
        self._adc_chan = adc_chan
        self._stop = False

    def stop(self) -> None:
        self._stop = True

    @QtCore.Slot()
    def run(self) -> None:
        try:
            result = self._ctrl.run_dense_loopback(
                AdwinDenseCaptureRequest(
                    sine_freq_hz=self._freq,
                    amplitude_v=self._amp,
                    io_rate_hz=self._io_rate,
                    duration_s=self._duration,
                    dac_chan=self._dac_chan,
                    adc_chan=self._adc_chan,
                ),
                timeout_s=max(30.0, self._duration + 10.0),
                should_stop=lambda: self._stop,
            )
            if not self._stop:
                self.capture_ready.emit(result)
            self.finished.emit()
        except Exception as exc:  # noqa: BLE001
            if self._stop:
                self.finished.emit()
                return
            self.failed.emit(str(exc))


# ---------------------------------------------------------------------------
# Self-test worker
# ---------------------------------------------------------------------------
class SelfTestWorker(QtCore.QObject):
    """Runs a sequenced hardware self-test without blocking the UI.

    Tests performed (all require a booted board; analog loopback tests require
    a short DAC→ADC patch cable):

    1. DLL / driver presence
    2. Board communication / digital-I/O readiness
    3. DAC write + ADC read-back at ±5 V and 0 V (requires loopback cable)
    4. Digital output set/get round-trip (relay bits 0..5 only)
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

        # ── Test 2: Board communication / digital-I/O readiness ───────────
        def t_ver():
            ver = self._ctrl.test_version()
            dig = self._ctrl.get_digout() & 0x3F
            if ver != 0:
                return True, f"Test_Version={ver}; Get_Digout=0x{dig:02X}"
            return True, f"Test_Version=0; digital I/O still responds (Get_Digout=0x{dig:02X})"
        total += 1
        if self._step("Board communication / I/O", t_ver):
            passed += 1

        # ── Test 3a: DAC → ADC loopback at +5 V ────────────────────────────
        def t_dac_pos():
            self._ctrl.set_dac(self._dac_ch, 5.0)
            time.sleep(0.05)
            v = self._ctrl.get_adc(self._adc_ch)
            self._ctrl.set_dac(self._dac_ch, 0.0)
            ok = abs(v - 5.0) < 0.5
            hint = ""
            if abs(v) < 0.05:
                hint = "  Likely no DAC→ADC loopback cable, wrong ADC channel, or open input."
            return ok, f"wrote +5.000V  read {v:+.3f}V  (Δ={v-5:.3f}V){hint}"
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
            hint = ""
            if abs(v) < 0.05:
                hint = "  Likely no DAC→ADC loopback cable, wrong ADC channel, or open input."
            return ok, f"wrote -5.000V  read {v:+.3f}V  (Δ={v+5:.3f}V){hint}"
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
            # VB6 only uses relay bits 0..5, so keep test patterns within 0x00..0x3F.
            initial = self._ctrl.get_digout() & 0x3F
            results = []
            ok = True
            for pattern in (0x2A, 0x15, 0x00):
                self._ctrl.set_digout(pattern)
                time.sleep(0.02)
                got = self._ctrl.get_digout() & 0x3F
                ok = ok and (got == pattern)
                results.append(f"set 0x{pattern:02X} got 0x{got:02X}")
            self._ctrl.set_digout(initial)
            return ok, " | ".join(results) + f" | restored 0x{initial:02X}"
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


# ---------------------------------------------------------------------------
# USB presence check (no DLL required)
# ---------------------------------------------------------------------------

def _check_adwin_usb_present() -> bool:
    """Return True if an ADwin USB device is listed in Windows Device Manager.

    Uses ``pnputil`` (built into Windows Vista+) — no extra packages needed.
    Falls back to True (assume present) if the check cannot be run so that
    the caller still attempts a connection.
    """
    try:
        import subprocess
        result = subprocess.run(
            ["pnputil", "/enum-devices"],
            capture_output=True, text=True, timeout=8,
            creationflags=subprocess.CREATE_NO_WINDOW,
        )
        return "adwin" in result.stdout.lower()
    except Exception:
        return True  # can't check → optimistically assume present


def _probe_adwin_device_numbers(max_device_num: int = 10) -> list[dict[str, object]]:
    """Probe candidate ADwin device numbers using the runtime DLL only.

    This is intentionally low-level diagnostic logic. It does not assume that
    a nonzero ``ADTest_Version`` return means success; it also inspects the
    ADwin error channel after each call.
    """
    dll_path = _find_adwin_dll()
    if not dll_path:
        raise AdwinError("ADwin runtime DLL not found. Install the ADwin driver first.")

    try:
        dll = ctypes.WinDLL(dll_path)
    except OSError as exc:
        raise AdwinError(f"Unable to load {dll_path}: {exc}") from exc

    test_version = dll.ADTest_Version
    test_version.argtypes = [ctypes.c_int, ctypes.c_int]
    test_version.restype = ctypes.c_int

    get_digout = dll.Get_Digout
    get_digout.argtypes = [ctypes.c_int]
    get_digout.restype = ctypes.c_long

    get_error_code = dll.ADGetErrorCode
    get_error_code.argtypes = []
    get_error_code.restype = ctypes.c_long

    get_error_text = dll.ADGetErrorText
    get_error_text.argtypes = [ctypes.c_long, ctypes.c_char_p, ctypes.c_long]
    get_error_text.restype = ctypes.c_long

    def read_error() -> tuple[int, str]:
        code = int(get_error_code())
        if code == 0:
            return 0, ""
        buf = ctypes.create_string_buffer(512)
        get_error_text(code, buf, len(buf))
        return code, buf.value.decode(errors="replace").strip()

    results: list[dict[str, object]] = []
    for device_num in range(1, max_device_num + 1):
        version_raw = int(test_version(device_num, 0))
        version_err_code, version_err_text = read_error()

        digout_raw = int(get_digout(device_num))
        digout_err_code, digout_err_text = read_error()

        errors: list[str] = []
        if version_err_code != 0:
            errors.append(f"Test_Version: {version_err_text or f'error {version_err_code}'}")
        if digout_err_code != 0:
            errors.append(f"Get_Digout: {digout_err_text or f'error {digout_err_code}'}")

        if not errors:
            status = "Booted" if version_raw != 0 else "Reachable, not booted"
            detail = "No ADwin error reported."
        elif version_err_code == 11 and digout_err_code == 11:
            status = "Unknown device"
            detail = "The device No. is not known."
        else:
            status = "Mixed / partial"
            detail = " | ".join(errors)

        results.append(
            {
                "device_num": device_num,
                "version_raw": version_raw,
                "digout_raw": digout_raw,
                "status": status,
                "detail": detail,
                "version_err_code": version_err_code,
                "digout_err_code": digout_err_code,
            }
        )

    return results


# ---------------------------------------------------------------------------
# Background connection worker
# ---------------------------------------------------------------------------

class AdwinConnectWorker(QtCore.QThread):
    """Attempts DLL load + board connection in a background thread.

    Emits
    -----
    log_msg(str)
        Progress messages for the console.
    connected(int, str)
        Board is responding: (test_version_code, btl_path_used).
        btl_path_used is "" when already booted (no reflash needed).
    failed(str, list)
        Could not connect: (human_readable_reason, list_of_suggestions).
    """

    log_msg   = QtCore.Signal(str)
    connected = QtCore.Signal(int, str)   # (version, btl_path)
    failed    = QtCore.Signal(str, list)  # (reason, suggestions)

    # Friendly hardware name hints shown in the error dialog
    _BTL_HINTS: dict[str, str] = {
        "adwin9.btl":   "ADwin Pro II  (T9 / T10)",
        "adwin10.btl":  "ADwin Pro II  (T10)",
        "adwin11.btl":  "ADwin Pro  /  Gold II  (T11)",
        "adwin12.btl":  "ADwin Pro II  (T12)",
        "adwin121.btl": "ADwin Pro II  (T12.1)",
        "adwins9.btl":  "ADwin Pro II Slot  (S9)",
    }

    def __init__(
        self,
        cfg: "AdwinCommsConfig",
        btl_override: str = "",
        force_reboot: bool = False,
        parent: QtCore.QObject | None = None,
    ) -> None:
        super().__init__(parent)
        self._cfg = cfg
        self._btl_override = btl_override.strip()
        self._force_reboot = force_reboot
        # ctrl is built inside run() to keep DLL load on the worker thread
        self.ctrl: AdwinAFController | None = None

    def run(self) -> None:
        import struct
        dll_name = "adwin32.dll" if struct.calcsize("P") == 4 else "adwin64.dll"

        # ── Step 1: Load DLL ───────────────────────────────────────────────
        self.log_msg.emit(f"Loading {dll_name}…")
        try:
            cfg = self._cfg
            bin_folder, boot_file = _split_btl_path(
                self._btl_override or cfg.btl_file, cfg.bin_folder, cfg.boot_file
            )
            board = AdwinBoardConfig(
                board_num=cfg.board_num,
                bin_folder=bin_folder,
                boot_file=boot_file,
            )
            ctrl = AdwinAFController(board=board)
            self.log_msg.emit(f"  {dll_name} loaded OK.")
        except AdwinError as exc:
            from rapidpy_common.adwin_af import _find_adwin_dll
            found = _find_adwin_dll()
            self.failed.emit(
                str(exc),
                [
                    f"Install ADwin software (includes {dll_name}):",
                    "  https://www.adwin.de/us/produkte/adbasic.html",
                    "Run the installer as Administrator.",
                    "Restart this application after installing.",
                    f"If already installed, verify {dll_name} exists in C:\\Windows\\",
                    f"  Currently found at: {found or '(not found)'}",
                ],
            )
            return

        # ── Step 2: Check if already booted ───────────────────────────────
        try:
            ver = ctrl.test_version()
        except Exception as exc:
            self.log_msg.emit(f"  Test_Version raised: {exc} — treating as 0")
            ver = 0

        if ver != 0 and not self._force_reboot:
            self.log_msg.emit(f"  Board already running (Test_Version → {ver}).")
            self.ctrl = ctrl
            self.connected.emit(ver, "")
            return

        if ver != 0 and self._force_reboot:
            self.log_msg.emit(f"  Board running (v{ver}) — force-reflash requested.")

        # ── Step 3: Check USB device presence before attempting ADboot ─────
        self.log_msg.emit("  Checking USB device presence…")
        if not _check_adwin_usb_present():
            self.failed.emit(
                "No ADwin device found in Windows Device Manager.",
                [
                    "Check the USB cable — unplug and re-plug the ADwin.",
                    "Make sure the ADwin power switch is ON (green LED lit).",
                    "Try a different USB port or cable.",
                    "Open Device Manager → look for 'ADWINDevice' class.",
                    "  Yellow warning ⚠ → right-click → Update driver.",
                    "Restart Windows, then retry.",
                ],
            )
            return

        # ── Step 4: Try each BTL file in priority order ────────────────────
        btl_files = find_btl_files(self._btl_override or ctrl.board.bin_folder)
        # Always try the currently configured file first (user may have browsed)
        current = self._btl_override or (
            str(Path(ctrl.board.bin_folder) / ctrl.board.boot_file)
            if ctrl.board.bin_folder else ""
        )
        if current and current not in btl_files:
            btl_files.insert(0, current)
        if not btl_files:
            self.failed.emit(
                "No .btl firmware files found.",
                [
                    "Use the Browse… button to locate your ADwin firmware file.",
                    "Typical location: C:\\ADwin\\ADwin9.btl",
                    "Install ADwin software if no .btl files are present.",
                ],
            )
            return

        tried: list[str] = []
        for btl_path in btl_files:
            btl_name = Path(btl_path).name
            hint = self._BTL_HINTS.get(btl_name.lower(), "")
            label = f"{btl_name}  ({hint})" if hint else btl_name
            self.log_msg.emit(f"  Trying {label}…")
            ctrl.board.bin_folder = str(Path(btl_path).parent)
            ctrl.board.boot_file = btl_name
            try:
                ctrl.boot_board()
                try:
                    ver2 = ctrl.test_version()
                except AdwinError as exc:
                    self.log_msg.emit(f"  Test_Version after boot raised: {exc}")
                    ver2 = 0

                # Legacy VB6 RAPID logic treats ADboot() returning success as the
                # key signal. On this machine, Test_Version remains 0 even though
                # digital I/O is immediately usable after boot, so verify readiness
                # using a benign read operation instead of requiring ver2 != 0.
                ready_word = ctrl.get_digout()
                if ver2 != 0:
                    self.log_msg.emit(f"  ✓ Booted with {btl_name}  (Test_Version → {ver2})")
                else:
                    self.log_msg.emit(
                        f"  ✓ Booted with {btl_name}; Test_Version stayed 0, "
                        f"but digital I/O is ready (Get_Digout → 0x{ready_word:02X})."
                    )
                self.ctrl = ctrl
                self.connected.emit(ver2, btl_path)
                return
            except AdwinError as exc:
                short = str(exc).split("\n")[0]
                tried.append(f"{label}: {short}")
                self.log_msg.emit(f"  ✗ {btl_name}: {short}")

        # ── All BTL files failed ───────────────────────────────────────────
        hints = "\n".join(
            f"  {n}  →  {h}" for n, h in self._BTL_HINTS.items() if not n.startswith("adwin1")
        )
        tried_str = "\n".join(f"  • {t}" for t in tried)
        self.failed.emit(
            f"USB device detected but firmware boot failed.\n\nAttempted:\n{tried_str}",
            [
                "Select the correct BTL file for your hardware using Browse…",
                f"Hardware → BTL file mapping:\n{hints}",
                "Check Device Manager → ADWINDevice — reinstall driver if yellow ⚠.",
                "Power-cycle the ADwin, wait 5 s, then retry.",
                "Tick 'Force reboot' and retry if the board may be in an odd state.",
            ],
        )


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
    # Connection helpers
    # -----------------------------------------------------------------------
    _connect_worker: "AdwinConnectWorker | None" = None

    def _try_init_dll(self) -> None:
        """Auto-connect at startup in a background thread — never blocks the UI."""
        self._lbl_boot_status.setText("\u25cf Connecting\u2026")
        self._lbl_boot_status.setStyleSheet("color: #b45309; font-weight: bold;")
        self._log("Auto-connecting to ADwin board\u2026")
        self._start_connect_worker(force_reboot=False, startup=True)

    def _start_connect_worker(
        self,
        force_reboot: bool = False,
        startup: bool = False,
    ) -> None:
        """Spin up an AdwinConnectWorker; wire its signals and start it."""
        # Cancel any still-running worker (guard against deleted C++ object)
        if self._connect_worker is not None:
            try:
                if self._connect_worker.isRunning():
                    self._connect_worker.quit()
                    self._connect_worker.wait(500)
            except RuntimeError:
                pass  # C++ object already deleted — nothing to stop
            self._connect_worker = None

        btl_override = "" if startup else self._edit_btl.text().strip()
        worker = AdwinConnectWorker(
            cfg=self._cfg,
            btl_override=btl_override,
            force_reboot=force_reboot,
            parent=self,
        )
        worker.log_msg.connect(self._log)
        worker.connected.connect(self._on_connected)
        worker.failed.connect(
            lambda reason, tips: self._on_connect_failed(reason, tips, startup=startup)
        )
        # NOTE: do NOT connect finished→deleteLater; parent=self keeps Qt ownership
        # safe, and deleteLater would invalidate the Python reference before we
        # access worker.ctrl in _on_connected.
        self._connect_worker = worker
        worker.start()

    def _on_connected(self, ver: int, btl_path: str) -> None:
        """Slot \u2014 called from worker when board is live."""
        if self._connect_worker is not None:
            self._ctrl = self._connect_worker.ctrl
        self._booted = True
        if ver != 0:
            self._lbl_boot_status.setText(f"\u25cf Connected  (v{ver})")
        else:
            self._lbl_boot_status.setText("\u25cf Connected  (I/O ready)")
        self._lbl_boot_status.setStyleSheet("color: #2e8b57; font-weight: bold;")
        if btl_path:
            self._edit_btl.setText(btl_path)
            self._cfg.btl_file = btl_path
            _save_config(self._cfg)
        self._update_hw_enabled()

    def _on_connect_failed(
        self, reason: str, suggestions: list, *, startup: bool
    ) -> None:
        """Slot \u2014 called from worker when all connection attempts failed."""
        self._ctrl = None
        self._booted = False
        self._lbl_boot_status.setText("\u25cf Not connected")
        self._lbl_boot_status.setStyleSheet("color: #cc0000;")
        first_line = reason.splitlines()[0] if reason else "Unknown error"
        self._log(f"[ERROR] {first_line}")
        self._update_hw_enabled()
        if "not found" in reason.lower() and ("dll" in reason.lower() or "adwin" in reason.lower()):
            title = "ADwin Driver Not Installed"
        elif "device manager" in reason.lower() or "usb" in reason.lower():
            title = "ADwin Board Not Detected on USB"
        elif "boot" in reason.lower():
            title = "ADwin Firmware Boot Failed"
        else:
            title = "ADwin Connection Failed"
        self._show_adwin_error_dialog(title, reason, suggestions)

    def _show_adwin_error_dialog(
        self,
        title: str,
        detail: str,
        suggestions: list,
    ) -> None:
        """Modal error dialog with structured fix instructions."""
        dlg = QtWidgets.QDialog(self)
        dlg.setWindowTitle(f"ADwin  \u2014  {title}")
        dlg.setMinimumWidth(540)
        dlg.setMaximumWidth(720)
        vl = QtWidgets.QVBoxLayout(dlg)
        vl.setSpacing(10)
        vl.setContentsMargins(18, 16, 18, 16)

        banner = QtWidgets.QLabel(f"\u26a0\u2009\u2009{title}")
        banner.setStyleSheet(
            "font-weight: bold; font-size: 14px; color: #b91c1c; padding-bottom: 4px;"
        )
        vl.addWidget(banner)

        if detail:
            detail_box = QtWidgets.QPlainTextEdit(detail)
            detail_box.setReadOnly(True)
            detail_box.setMaximumHeight(130)
            detail_box.setStyleSheet(
                "font-family: Consolas, monospace; font-size: 11px;"
                "background: #f8f4f4; border: 1px solid #e0c8c8; border-radius: 6px;"
            )
            vl.addWidget(detail_box)

        if suggestions:
            vl.addWidget(QtWidgets.QLabel("<b>How to fix:</b>"))
            for i, s in enumerate(suggestions, 1):
                lbl = QtWidgets.QLabel(f"  {i}.\u2002{s}")
                lbl.setWordWrap(True)
                lbl.setTextInteractionFlags(
                    QtCore.Qt.TextInteractionFlag.TextSelectableByMouse
                )
                vl.addWidget(lbl)

        vl.addSpacing(6)
        btns = QtWidgets.QDialogButtonBox(QtWidgets.QDialogButtonBox.StandardButton.Ok)
        btns.accepted.connect(dlg.accept)
        vl.addWidget(btns)
        dlg.exec()

    def _update_hw_enabled(self) -> None:
        enabled = self._ctrl is not None
        booted = enabled and self._booted
        for w in self._hw_widgets:
            w.setEnabled(enabled)
        for w in self._booted_widgets:
            w.setEnabled(booted)
        self._btn_boot.setEnabled(True)   # always enabled so user can retry

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
        # Row 1: Device-number diagnostics spans both columns
        ctrl_grid.addWidget(self._build_device_diag_card(), 1, 0, 1, 2)
        # Row 2: Direct DAC/ADC spans both columns
        ctrl_grid.addWidget(self._build_dac_adc_card(), 2, 0, 1, 2)
        # Row 3: Sine Loopback spans both columns
        ctrl_grid.addWidget(self._build_siggen_card(),  3, 0, 1, 2)
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
        self._edit_btl.setMinimumWidth(170)
        self._edit_btl.setMaximumWidth(245)
        self._edit_btl.setPlaceholderText(
            default_btl if default_btl else "e.g. C:\\ADwin\\BTL\\ADwin9.btl"
        )
        self._edit_btl.setToolTip(
            "Full path to the ADwin firmware boot file (.btl).\n"
            "Typically  C:\\ADwin\\BTL\\ADwin9.btl"
        )
        row_btl.addWidget(self._edit_btl, 1)
        btn_browse_btl = QtWidgets.QPushButton("Browse…")
        btn_browse_btl.setMinimumWidth(92)
        btn_browse_btl.setMaximumWidth(92)
        btn_browse_btl.clicked.connect(self._browse_btl)
        row_btl.addWidget(btn_browse_btl)
        layout.addLayout(row_btl)

        # Boot button + status
        row_boot = QtWidgets.QHBoxLayout()
        self._btn_boot = QtWidgets.QPushButton("Connect / Boot")
        self._btn_boot.setToolTip(
            "Connect to an already-running board, or flash firmware if not yet booted.\n"
            "Check \"Force reboot\" to always reflash (slow)."
        )
        self._btn_boot.clicked.connect(self._boot_board)
        row_boot.addWidget(self._btn_boot)
        self._lbl_boot_status = QtWidgets.QLabel("● Not connected")
        self._lbl_boot_status.setStyleSheet("color: #888;")
        row_boot.addWidget(self._lbl_boot_status)
        row_boot.addStretch()
        layout.addLayout(row_boot)

        self._chk_force_reboot = QtWidgets.QCheckBox("Force reboot (reflash firmware)")
        self._chk_force_reboot.setToolTip(
            "Always call ADboot even if the board is already running.\n"
            "Only needed after a firmware upgrade or a hardware reset."
        )
        self._chk_force_reboot.setStyleSheet("font-size: 10px; color: #666;")
        layout.addWidget(self._chk_force_reboot)

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
            btn = QtWidgets.QPushButton()
            btn.setCheckable(True)
            btn.setChecked(False)
            btn.setToolTip(f"Toggle digital output bit {bit}")
            btn.setMinimumHeight(42)
            btn.toggled.connect(lambda checked, b=bit: self._relay_toggled(b, checked))
            self._relay_btns.append(btn)
            self._booted_widgets.append(btn)
            grid.addWidget(btn, bit // 2, bit % 2)
            self._update_relay_button(bit)
        layout.addLayout(grid)

        # Raw word + buttons
        row = QtWidgets.QHBoxLayout()
        self._lbl_digout = QtWidgets.QLabel("Digout: 0x00")
        self._lbl_digout.setStyleSheet("font-family: monospace;")
        row.addWidget(self._lbl_digout)
        row.addStretch()
        btn_read_dig = QtWidgets.QPushButton("Read State")
        btn_read_dig.setMinimumWidth(82)
        btn_read_dig.setStyleSheet("font-size: 10px;")
        btn_read_dig.clicked.connect(self._read_digout)
        self._booted_widgets.append(btn_read_dig)
        row.addWidget(btn_read_dig)
        btn_all_off = QtWidgets.QPushButton("All OFF")
        btn_all_off.setMinimumWidth(70)
        btn_all_off.setStyleSheet("font-size: 10px;")
        btn_all_off.clicked.connect(self._all_relays_off)
        self._booted_widgets.append(btn_all_off)
        row.addWidget(btn_all_off)
        layout.addLayout(row)

        return frame

    # ── Device-number diagnostics card ─────────────────────────────────────
    def _build_device_diag_card(self) -> QtWidgets.QFrame:
        frame, layout = self._card("Device # Diagnostics")

        info = QtWidgets.QLabel(
            "Scans candidate ADwin device numbers using the DLL only. "
            "A valid device number should report no ADwin error.\n"
            "Important: ADTest_Version can return a nonzero raw value even when the DLL "
            "still says the device number is unknown."
        )
        info.setWordWrap(True)
        info.setStyleSheet("font-size: 10px; color: #666;")
        layout.addWidget(info)

        row = QtWidgets.QHBoxLayout()
        self._btn_diag_scan = QtWidgets.QPushButton("Scan 1–10")
        self._btn_diag_scan.clicked.connect(self._scan_device_numbers)
        row.addWidget(self._btn_diag_scan)

        self._btn_diag_apply = QtWidgets.QPushButton("Use Selected #")
        self._btn_diag_apply.setEnabled(False)
        self._btn_diag_apply.clicked.connect(self._apply_selected_device_number)
        row.addWidget(self._btn_diag_apply)
        row.addStretch()

        self._lbl_diag_summary = QtWidgets.QLabel("No scan run yet.")
        self._lbl_diag_summary.setStyleSheet("font-size: 10px; color: #666;")
        row.addWidget(self._lbl_diag_summary)
        layout.addLayout(row)

        self._tbl_diag = QtWidgets.QTableWidget(0, 5)
        self._tbl_diag.setHorizontalHeaderLabels(["Dev #", "Test", "DigOut", "Status", "Detail"])
        self._tbl_diag.setSelectionBehavior(QtWidgets.QAbstractItemView.SelectRows)
        self._tbl_diag.setSelectionMode(QtWidgets.QAbstractItemView.SingleSelection)
        self._tbl_diag.setEditTriggers(QtWidgets.QAbstractItemView.NoEditTriggers)
        self._tbl_diag.setAlternatingRowColors(True)
        self._tbl_diag.setMinimumHeight(150)
        self._tbl_diag.verticalHeader().setVisible(False)
        self._tbl_diag.horizontalHeader().setStretchLastSection(True)
        self._tbl_diag.horizontalHeader().setSectionResizeMode(0, QtWidgets.QHeaderView.ResizeToContents)
        self._tbl_diag.horizontalHeader().setSectionResizeMode(1, QtWidgets.QHeaderView.ResizeToContents)
        self._tbl_diag.horizontalHeader().setSectionResizeMode(2, QtWidgets.QHeaderView.ResizeToContents)
        self._tbl_diag.horizontalHeader().setSectionResizeMode(3, QtWidgets.QHeaderView.ResizeToContents)
        self._tbl_diag.itemSelectionChanged.connect(self._update_device_diag_buttons)
        self._tbl_diag.itemDoubleClicked.connect(lambda _item: self._apply_selected_device_number())
        layout.addWidget(self._tbl_diag)

        return frame

    # ── Direct DAC / ADC card ────────────────────────────────────────────────
    def _build_dac_adc_card(self) -> QtWidgets.QFrame:
        frame, layout = self._card("Direct DAC / ADC")

        note = QtWidgets.QLabel(
            "No process file required — works immediately after board boot.\n"
            "DAC = digital-to-analog output. Write a voltage on channel 1–8 in the range -10 V to +10 V.\n"
            "ADC = analog-to-digital input. Read back the measured voltage on channel 1–8, also in roughly -10 V to +10 V."
        )
        note.setStyleSheet("font-size: 10px; color: #666;")
        note.setWordWrap(True)
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
        btn_write_dac = QtWidgets.QPushButton("Write DAC")
        btn_write_dac.setMinimumWidth(82)
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
        self._spin_freq.setSingleStep(0.5)
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
        self._spin_iorate.setRange(1.0, 50000.0)
        self._spin_iorate.setDecimals(0)
        self._spin_iorate.setSingleStep(100.0)
        self._spin_iorate.setValue(self._cfg.sig_io_rate)
        self._spin_iorate.setSuffix(" Hz")
        left_form.addRow("IO Rate:", self._spin_iorate)

        self._lbl_sig_hint = QtWidgets.QLabel()
        self._lbl_sig_hint.setWordWrap(True)
        self._lbl_sig_hint.setStyleSheet("font-size: 10px; color: #666;")
        left_form.addRow("Sampling:", self._lbl_sig_hint)

        self._spin_freq.valueChanged.connect(self._update_sig_hint)
        self._spin_iorate.valueChanged.connect(self._update_sig_hint)
        self._update_sig_hint()

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
        btn_save_plot = QtWidgets.QPushButton("Save Plot")
        btn_save_plot.setMinimumWidth(90)
        btn_save_plot.clicked.connect(self._save_plot_image)
        title_row.addWidget(btn_save_plot)
        btn_save_csv = QtWidgets.QPushButton("Save CSV")
        btn_save_csv.setMinimumWidth(90)
        btn_save_csv.clicked.connect(self._save_plot_csv)
        title_row.addWidget(btn_save_csv)
        btn_clear_plot = QtWidgets.QPushButton("Clear Plot")
        btn_clear_plot.setMinimumWidth(90)
        btn_clear_plot.clicked.connect(self._clear_plot)
        title_row.addWidget(btn_clear_plot)
        vlay.addLayout(title_row)

        help_lbl = QtWidgets.QLabel(
            "Left-drag to draw a zoom box. Double-click anywhere in the plot to auto-rescale. "
            "Save Plot writes the current chart image, Save CSV writes time/DAC/ADC samples, and "
            "Clear Plot resets the view to the default time/voltage window."
        )
        help_lbl.setWordWrap(True)
        help_lbl.setStyleSheet("font-size: 10px; color: #666;")
        vlay.addWidget(help_lbl)

        pg.setConfigOptions(antialias=True, background="#1a1a2e", foreground="#cccccc")
        self._plot = pg.PlotWidget()
        self._plot.setLabel("bottom", "Time", units="s")
        self._plot.setLabel("left", "Voltage", units="V")
        self._plot.addLegend(offset=(10, 10))
        self._plot.showGrid(x=True, y=True, alpha=0.2)
        self._plot.setMouseEnabled(x=True, y=True)
        self._plot.getViewBox().setMouseMode(pg.ViewBox.RectMode)
        self._plot.scene().sigMouseClicked.connect(self._on_plot_mouse_clicked)

        self._curve_dac = self._plot.plot(pen=pg.mkPen("#4DFF91", width=2), name="DAC out")
        self._curve_adc = self._plot.plot(pen=pg.mkPen("#FFCD34", width=2), name="ADC in")

        vlay.addWidget(self._plot)
        self._reset_plot_view()
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

    def _update_device_diag_buttons(self) -> None:
        self._btn_diag_apply.setEnabled(self._tbl_diag.currentRow() >= 0)

    def _scan_device_numbers(self) -> None:
        self._log("Scanning ADwin device numbers 1..10…")
        self._btn_diag_scan.setEnabled(False)
        self._btn_diag_apply.setEnabled(False)
        self._lbl_diag_summary.setText("Scanning…")
        self._tbl_diag.setRowCount(0)
        QtWidgets.QApplication.setOverrideCursor(QtCore.Qt.WaitCursor)
        try:
            results = _probe_adwin_device_numbers(10)
        except Exception as exc:
            self._lbl_diag_summary.setText("Scan failed")
            self._log(f"[ERROR] Device-number scan failed: {exc}")
            return
        finally:
            QtWidgets.QApplication.restoreOverrideCursor()
            self._btn_diag_scan.setEnabled(True)

        current_board = self._spin_board.value()
        selectable_rows: list[int] = []
        for row_idx, result in enumerate(results):
            self._tbl_diag.insertRow(row_idx)
            items = [
                QtWidgets.QTableWidgetItem(str(result["device_num"])),
                QtWidgets.QTableWidgetItem(str(result["version_raw"])),
                QtWidgets.QTableWidgetItem(f"0x{int(result['digout_raw']) & 0xFF:02X}"),
                QtWidgets.QTableWidgetItem(str(result["status"])),
                QtWidgets.QTableWidgetItem(str(result["detail"])),
            ]
            for col_idx, item in enumerate(items):
                self._tbl_diag.setItem(row_idx, col_idx, item)

            if result["status"] != "Unknown device":
                selectable_rows.append(row_idx)
            if int(result["device_num"]) == current_board:
                for item in items:
                    item.setBackground(QtGui.QColor("#fff7cc"))

        if selectable_rows:
            first_row = selectable_rows[0]
            self._tbl_diag.selectRow(first_row)
            candidates = ", ".join(
                str(results[idx]["device_num"]) for idx in selectable_rows
            )
            self._lbl_diag_summary.setText(f"Reachable candidates: {candidates}")
            self._log(f"Device scan complete. Reachable candidates: {candidates}")
        else:
            self._lbl_diag_summary.setText("No reachable device number found")
            self._log(
                "Device scan complete. ADwin reported 'device No. is not known' "
                "for every probed device number."
            )

        self._update_device_diag_buttons()

    def _apply_selected_device_number(self) -> None:
        row = self._tbl_diag.currentRow()
        if row < 0:
            return
        item = self._tbl_diag.item(row, 0)
        if item is None:
            return
        device_num = int(item.text())
        self._spin_board.setValue(device_num)
        self._sync_cfg()
        _save_config(self._cfg)
        self._log(f"Board # set to {device_num} from diagnostics panel.")

    def _update_relay_button(self, bit: int) -> None:
        btn = self._relay_btns[bit]
        label = self._cfg.bit_labels.get(str(bit), f"Bit {bit}")
        is_on = btn.isChecked()
        btn.setText(f"{label}\n{'ON' if is_on else 'OFF'}")
        if is_on:
            btn.setStyleSheet(
                "QPushButton { background: #2e8b57; color: white; font-weight: bold; "
                "border: 1px solid #245f44; border-radius: 8px; padding: 6px; }"
            )
        else:
            btn.setStyleSheet(
                "QPushButton { background: #f3f4f6; color: #374151; font-weight: bold; "
                "border: 1px solid #c7ccd4; border-radius: 8px; padding: 6px; }"
            )

    def _refresh_relay_buttons(self) -> None:
        for bit in range(len(self._relay_btns)):
            self._update_relay_button(bit)

    def _update_sig_hint(self) -> None:
        freq = max(self._spin_freq.value(), 0.1)
        io_rate = max(self._spin_iorate.value(), 1.0)
        points_per_cycle = io_rate / freq
        if points_per_cycle >= 40:
            quality = "smooth"
            color = "#2e8b57"
        elif points_per_cycle >= 20:
            quality = "usable"
            color = "#9a6700"
        else:
            quality = "jagged"
            color = "#b91c1c"
        self._lbl_sig_hint.setStyleSheet(f"font-size: 10px; color: {color};")
        self._lbl_sig_hint.setText(
            f"{points_per_cycle:.1f} samples/cycle. This uses the ADwin-side process, "
            f"so the requested IO rate is board-timed rather than host-polled. Aim for at least 20, "
            f"preferably 40+, for a smooth plotted sine."
        )

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
        """Connect / Boot button handler — launches background worker."""
        self._apply_board_cfg()
        force = self._chk_force_reboot.isChecked()
        self._lbl_boot_status.setText("● Connecting…")
        self._lbl_boot_status.setStyleSheet("color: #b45309; font-weight: bold;")
        self._log(f"{'Force-rebooting' if force else 'Connecting to'} ADwin board…")
        self._start_connect_worker(force_reboot=force, startup=False)

    # -----------------------------------------------------------------------
    # Relay operations
    # -----------------------------------------------------------------------
    def _relay_toggled(self, bit: int, checked: bool) -> None:
        if not self._booted or self._ctrl is None:
            return
        prev_state = self._digout_state
        if checked:
            self._digout_state |= (1 << bit)
        else:
            self._digout_state &= ~(1 << bit)
        self._digout_state &= 0x3F
        try:
            self._ctrl.set_digout(self._digout_state)
            self._lbl_digout.setText(f"Digout: 0x{self._digout_state:02X}")
            label = self._cfg.bit_labels.get(str(bit), f"Bit {bit}")
            self._log(f"Bit {bit} ({label}) → {'ON' if checked else 'OFF'}  (word=0x{self._digout_state:02X})")
            self._refresh_relay_buttons()
        except AdwinError as exc:
            self._digout_state = prev_state
            btn = self._relay_btns[bit]
            btn.blockSignals(True)
            btn.setChecked(bool(prev_state & (1 << bit)))
            btn.blockSignals(False)
            self._refresh_relay_buttons()
            self._log(f"[ERROR] {exc}")

    def _read_digout(self) -> None:
        if self._ctrl is None:
            return
        try:
            raw_val = self._ctrl.get_digout()
            val = raw_val & 0x3F
            self._digout_state = val
            self._lbl_digout.setText(f"Digout: 0x{val:02X}")
            if raw_val != val:
                self._log(
                    f"Read digout → raw 0x{raw_val:02X}; using relay bits 0..5 only → 0x{val:02X}"
                )
            else:
                self._log(f"Read digout → 0x{val:02X} (binary: {val:08b})")
            # Sync toggle buttons without re-triggering hardware writes
            for bit, btn in enumerate(self._relay_btns):
                btn.blockSignals(True)
                btn.setChecked(bool(val & (1 << bit)))
                btn.blockSignals(False)
            self._refresh_relay_buttons()
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
            self._refresh_relay_buttons()
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
        points_per_cycle = io_rate / max(freq, 0.1)
        process_file = _default_process_file()
        if not process_file:
            self._lbl_sig_status.setText("Process file missing.")
            self._log("[ERROR] Dense ADwin process file not found: VB6/ADwin/sineout.T91")
            return
        self._ctrl.board.process_file = process_file

        # Clear old data
        self._t_buf.clear()
        self._dac_buf.clear()
        self._adc_buf.clear()
        self._reset_plot_view()

        self._worker = SineLoopbackWorker(
            self._ctrl, freq, amp, dur, io_rate, dac_ch, adc_ch
        )
        self._worker_thread = QtCore.QThread(self)
        self._worker.moveToThread(self._worker_thread)
        self._worker_thread.started.connect(self._worker.run)
        self._worker.capture_ready.connect(self._on_capture_ready)
        self._worker.finished.connect(self._on_sig_finished)
        self._worker.failed.connect(self._on_sig_failed)

        self._btn_run_sig.setEnabled(False)
        self._btn_stop_sig.setEnabled(True)
        self._lbl_sig_status.setText(
            f"Running on ADwin…  {freq:.1f} Hz, {amp:.3f} V, {dur:.1f} s, {points_per_cycle:.1f} samples/cycle"
        )
        if points_per_cycle < 20:
            self._log(
                f"[WARNING] Dense capture sampling is sparse: {points_per_cycle:.1f} samples/cycle. "
                "The plot will look jagged. Increase IO Rate or lower Frequency."
            )
        self._log(
            f"Dense ADwin capture started: freq={freq}Hz amp={amp}V dur={dur}s "
            f"DAC→ch{dac_ch} ADC←ch{adc_ch} ({points_per_cycle:.1f} samples/cycle)"
        )
        self._worker_thread.start()

    def _stop_sig(self) -> None:
        if self._worker is not None:
            self._worker.stop()
        self._lbl_sig_status.setText("Stopping ADwin capture…")

    @QtCore.Slot(object)
    def _on_capture_ready(self, result: object) -> None:
        capture = result if isinstance(result, AdwinDenseCaptureResult) else None
        if capture is None:
            return
        self._t_buf.clear()
        self._dac_buf.clear()
        self._adc_buf.clear()
        self._t_buf.extend(capture.time_s)
        self._dac_buf.extend(capture.dac_v)
        self._adc_buf.extend(capture.adc_v)
        self._flush_plot()
        self._log(
            "Dense ADwin capture received: "
            f"{len(capture.time_s)} points, dt={capture.timestep_s * 1000.0:.3f} ms, "
            f"{capture.points_per_period:.1f} samples/cycle"
        )

    @QtCore.Slot()
    def _on_sig_finished(self) -> None:
        self._cleanup_worker()
        self._lbl_sig_status.setText("Done.")
        self._log("Dense ADwin capture finished.")

    @QtCore.Slot(str)
    def _on_sig_failed(self, msg: str) -> None:
        self._cleanup_worker()
        self._lbl_sig_status.setText("Error.")
        self._log(f"[ERROR] Dense ADwin capture: {msg}")

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
            self._selftest_thread.wait(3000)
            self._selftest_thread = None
        self._selftest_worker = None
        self._btn_selftest.setEnabled(self._booted and self._ctrl is not None)
        summary_color = "#2e8b57" if passed == total else "#b91c1c"
        self._lbl_selftest_summary.setStyleSheet(f"font-size: 11px; color: {summary_color};")
        self._lbl_selftest_summary.setText(f"Self-test complete: {passed}/{total} passed.")
        self._log(f"[SELFTEST] Complete — {passed}/{total} passed.")

    # -----------------------------------------------------------------------
    def _flush_plot(self) -> None:
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
        self._reset_plot_view()

    def _capture_basename(self) -> str:
        from datetime import datetime

        stamp = datetime.now().strftime("%Y%m%d-%H%M%S")
        return f"adwin-loopback-{stamp}"

    def _save_plot_image(self) -> None:
        if not hasattr(self, "_plot"):
            return
        if len(self._t_buf) == 0:
            self._log("[WARNING] No loopback capture is loaded; run a test before saving the plot.")
            return

        default_path = str(Path.home() / f"{self._capture_basename()}.png")
        path, _ = QtWidgets.QFileDialog.getSaveFileName(
            self,
            "Save loopback plot",
            default_path,
            "PNG image (*.png);;JPEG image (*.jpg *.jpeg);;All files (*)",
        )
        if not path:
            return

        pixmap = self._plot.grab()
        if not pixmap.save(path):
            self._log(f"[ERROR] Failed to save plot image: {path}")
            return
        self._log(f"Saved loopback plot image: {path}")

    def _save_plot_csv(self) -> None:
        if len(self._t_buf) == 0:
            self._log("[WARNING] No loopback capture is loaded; run a test before saving CSV data.")
            return

        default_path = str(Path.home() / f"{self._capture_basename()}.csv")
        path, _ = QtWidgets.QFileDialog.getSaveFileName(
            self,
            "Save loopback CSV",
            default_path,
            "CSV files (*.csv);;All files (*)",
        )
        if not path:
            return

        rows = zip(self._t_buf, self._dac_buf, self._adc_buf)
        with open(path, "w", newline="", encoding="utf-8") as fh:
            writer = csv.writer(fh)
            writer.writerow(["time_s", "dac_v", "adc_v"])
            writer.writerows(rows)
        self._log(f"Saved loopback CSV: {path}")

    def _default_plot_ranges(self) -> tuple[tuple[float, float], tuple[float, float]]:
        duration = max(float(self._spin_dur.value()), 1.0) if hasattr(self, "_spin_dur") else 5.0
        amplitude = max(float(self._spin_amp.value()), 0.5) if hasattr(self, "_spin_amp") else 1.0
        x_range = (0.0, duration)
        y_pad = max(0.25, amplitude * 1.25)
        y_range = (-y_pad, y_pad)
        return x_range, y_range

    def _reset_plot_view(self) -> None:
        if not hasattr(self, "_plot"):
            return
        x_range, y_range = self._default_plot_ranges()
        self._plot.setXRange(*x_range, padding=0.0)
        self._plot.setYRange(*y_range, padding=0.0)

    def _on_plot_mouse_clicked(self, event) -> None:
        if not hasattr(self, "_plot") or not event.double():
            return
        if self._plot.getViewBox().sceneBoundingRect().contains(event.scenePos()):
            self._plot.enableAutoRange(axis=pg.ViewBox.XYAxes, enable=True)
            self._plot.getViewBox().autoRange()
            self._plot.enableAutoRange(axis=pg.ViewBox.XYAxes, enable=False)

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
