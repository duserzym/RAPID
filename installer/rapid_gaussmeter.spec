# PyInstaller spec for RapidPy Gaussmeter Control
# Run from repo root:
#   pyinstaller installer\rapid_gaussmeter.spec
#
# Requires PyInstaller:  pip install pyinstaller
# Output:  dist\RapidPy_Gaussmeter\  (one-folder bundle)

from pathlib import Path
import sys

REPO_ROOT   = Path(SPECPATH).parent          # e.g. E:\Github\RAPID
APP_DIR     = REPO_ROOT / "RapidPy" / "gaussmeter_control"
ICON_PATH   = APP_DIR / "assets" / "gaussmeter_icon.ico"
COMMON_DIR = REPO_ROOT / "RapidPy" / "rapidpy_common"
TOOLS_DIR  = REPO_ROOT / "tools"
LIB_DIR    = REPO_ROOT / "lib"

# Collect FW Bell DLLs to bundle alongside usb5100_probe.exe.
# Both DLLs must be present in lib\ (copy them there before building).
# usb5100.dll and libusb0.dll go into tools\ so the probe subprocess finds
# them in its own directory (standard Windows DLL search path).
_fw_bell_dlls = []
for _fname in ("usb5100.dll", "libusb0.dll"):
    _p = LIB_DIR / _fname
    if _p.exists():
        _fw_bell_dlls.append((str(_p), "tools"))

a = Analysis(
    [str(APP_DIR / "main.py")],
    pathex=[
        str(APP_DIR),
        str(REPO_ROOT / "RapidPy"),
        str(REPO_ROOT),
    ],
    binaries=[
        # Bundle the x86 helper so the frozen app can call it at runtime.
        (str(TOOLS_DIR / "usb5100_probe.exe"), "tools"),
    ],
    datas=[
        # Include the rapidpy_common package data if any exists.
        (str(COMMON_DIR), "rapidpy_common"),
        # FW Bell DLLs alongside the probe so it can find them at subprocess launch.
        *_fw_bell_dlls,
    ],
    hiddenimports=[
        "pyqtgraph",
        "pyqtgraph.exporters",
        "serial",
        "serial.tools",
        "serial.tools.list_ports",
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
)

pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name="RapidPy_Gaussmeter",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,              # No terminal window
    disable_windowed_traceback=False,
    argv_emulation=False,
    icon=str(ICON_PATH) if ICON_PATH.exists() else None,
)

coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name="RapidPy_Gaussmeter",
)
