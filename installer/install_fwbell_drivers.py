"""
RapidPy FW Bell Driver Installer
=================================
Standalone installer that copies usb5100.dll and libusb0.dll into the
system driver location used by the RapidPy gaussmeter app:

    C:\\Program Files\\RapidPy\\Gaussmeter\\drivers\\

When run as a PyInstaller frozen exe the DLLs are expected to be bundled
alongside this exe in a "drivers" subfolder. The installer copies them
to the destination and optionally adds the destination directory to the
system PATH so Windows can resolve the DLLs.

Usage:
    install_fwbell_drivers.exe               -- interactive GUI installer
    install_fwbell_drivers.exe --silent      -- silent, returns 0 on success

Requires elevation (Administrator) to write to Program Files.
"""
from __future__ import annotations

import argparse
import ctypes
import os
import shutil
import subprocess
import sys
import winreg
from pathlib import Path
from tkinter import filedialog, messagebox
import tkinter as tk


# ── Constants ─────────────────────────────────────────────────────────────────

DRIVER_FILES = ("usb5100.dll", "libusb0.dll")

DEST_DIR = Path(r"C:\Program Files\RapidPy\Gaussmeter\drivers")

KNOWN_SRC_DIRS = [
    Path(r"C:\Program Files (x86)\FW Bell\PC5180"),
    Path(r"C:\Program Files\FW Bell\PC5180"),
    Path(r"C:\FWBell\PC5180"),
]


# ── Helpers ───────────────────────────────────────────────────────────────────

def _is_admin() -> bool:
    try:
        return bool(ctypes.windll.shell32.IsUserAnAdmin())
    except Exception:
        return False


def _relaunch_as_admin() -> None:
    """Re-launch this exe with UAC elevation."""
    script = sys.executable
    params = " ".join(f'"{a}"' for a in sys.argv)
    ctypes.windll.shell32.ShellExecuteW(None, "runas", script, params, None, 1)
    sys.exit(0)


def _find_src_dir() -> Path | None:
    """Return the first directory that contains both required DLLs."""
    # When frozen by PyInstaller, data files live in sys._MEIPASS / "drivers".
    if getattr(sys, "frozen", False) and hasattr(sys, "_MEIPASS"):
        meipass_drivers = Path(sys._MEIPASS) / "drivers"
        if all((meipass_drivers / f).is_file() for f in DRIVER_FILES):
            return meipass_drivers

    # Also check a drivers/ subfolder next to the exe (for manual/dev layout).
    exe_drivers = Path(sys.executable).parent / "drivers"
    if all((exe_drivers / f).is_file() for f in DRIVER_FILES):
        return exe_drivers

    for d in KNOWN_SRC_DIRS:
        if all((d / f).is_file() for f in DRIVER_FILES):
            return d

    return None


def _add_to_system_path(directory: str) -> bool:
    """Add *directory* to the System PATH in the registry. Returns True on success."""
    try:
        key = winreg.OpenKey(
            winreg.HKEY_LOCAL_MACHINE,
            r"SYSTEM\CurrentControlSet\Control\Session Manager\Environment",
            0,
            winreg.KEY_READ | winreg.KEY_WRITE,
        )
        current, _ = winreg.QueryValueEx(key, "Path")
        entries = [e for e in current.split(";") if e]
        if directory.lower() not in (e.lower() for e in entries):
            entries.append(directory)
            new_path = ";".join(entries)
            winreg.SetValueEx(key, "Path", 0, winreg.REG_EXPAND_SZ, new_path)
            # Broadcast WM_SETTINGCHANGE so open processes pick up the new PATH.
            ctypes.windll.user32.SendMessageTimeoutW(0xFFFF, 0x001A, 0, "Environment", 2, 5000, None)
        winreg.CloseKey(key)
        return True
    except OSError:
        return False


def do_install(src_dir: Path, dest_dir: Path, add_to_path: bool = True) -> list[str]:
    """Copy DLLs from *src_dir* to *dest_dir*. Returns list of error strings."""
    errors: list[str] = []
    dest_dir.mkdir(parents=True, exist_ok=True)
    for fname in DRIVER_FILES:
        src = src_dir / fname
        dst = dest_dir / fname
        if not src.is_file():
            errors.append(f"Source not found: {src}")
            continue
        try:
            shutil.copy2(src, dst)
        except OSError as exc:
            errors.append(f"Failed to copy {fname}: {exc}")
    if not errors and add_to_path:
        if not _add_to_system_path(str(dest_dir)):
            errors.append(f"Could not add {dest_dir} to system PATH (non-fatal).")
    return errors


# ── GUI installer ─────────────────────────────────────────────────────────────

def run_gui() -> int:
    root = tk.Tk()
    root.withdraw()

    src_dir = _find_src_dir()

    if src_dir is None:
        answer = messagebox.askyesno(
            "RapidPy FW Bell Driver Installer",
            "Could not auto-detect the FW Bell driver files.\n\n"
            "Would you like to browse for the folder containing\n"
            "usb5100.dll and libusb0.dll?",
        )
        if not answer:
            messagebox.showinfo("Cancelled", "Installation cancelled.")
            return 1
        chosen = filedialog.askdirectory(title="Select folder containing usb5100.dll and libusb0.dll")
        if not chosen:
            return 1
        src_dir = Path(chosen)
        missing = [f for f in DRIVER_FILES if not (src_dir / f).is_file()]
        if missing:
            messagebox.showerror(
                "Files Not Found",
                f"The following files were not found in the selected folder:\n"
                + "\n".join(f"  {f}" for f in missing),
            )
            return 1
    else:
        answer = messagebox.askyesno(
            "RapidPy FW Bell Driver Installer",
            f"Driver files found at:\n  {src_dir}\n\n"
            f"Install to:\n  {DEST_DIR}\n\n"
            "Proceed?",
        )
        if not answer:
            messagebox.showinfo("Cancelled", "Installation cancelled.")
            return 1

    errors = do_install(src_dir, DEST_DIR, add_to_path=True)

    if errors:
        # Non-fatal PATH error is shown as a warning, not failure.
        fatal = [e for e in errors if "PATH" not in e]
        warn  = [e for e in errors if "PATH" in e]
        if fatal:
            messagebox.showerror("Installation Failed", "\n".join(fatal))
            return 1
        if warn:
            messagebox.showwarning("Partial Success", "\n".join(warn))
    else:
        messagebox.showinfo(
            "Installation Complete",
            f"Driver files installed to:\n  {DEST_DIR}\n\n"
            "The RapidPy Gaussmeter app will find them automatically on next launch.",
        )
    return 0


# ── Silent installer ──────────────────────────────────────────────────────────

def run_silent() -> int:
    src_dir = _find_src_dir()
    if src_dir is None:
        print("ERROR: Could not find usb5100.dll and libusb0.dll.", file=sys.stderr)
        print("Pass --src <dir> or place the DLLs in a drivers\\ subfolder next to this exe.", file=sys.stderr)
        return 1
    errors = do_install(src_dir, DEST_DIR, add_to_path=True)
    for e in errors:
        print(f"ERROR: {e}", file=sys.stderr)
    fatal = [e for e in errors if "PATH" not in e]
    if fatal:
        return 1
    print(f"OK: Driver files installed to {DEST_DIR}")
    return 0


# ── Entry point ───────────────────────────────────────────────────────────────

def main() -> int:
    parser = argparse.ArgumentParser(description="RapidPy FW Bell driver installer")
    parser.add_argument("--silent", action="store_true", help="Run without GUI")
    parser.add_argument("--src", metavar="DIR", help="Source folder containing the DLL files")
    args = parser.parse_args()

    # Honour explicit source override.
    if args.src:
        KNOWN_SRC_DIRS.insert(0, Path(args.src))

    if not _is_admin():
        if args.silent:
            print("ERROR: Administrator privileges required.", file=sys.stderr)
            return 1
        _relaunch_as_admin()
        return 0  # The relaunched process takes over.

    if args.silent:
        return run_silent()
    return run_gui()


if __name__ == "__main__":
    raise SystemExit(main())
