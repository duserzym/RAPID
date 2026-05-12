# PyInstaller spec for RapidPy Gaussmeter Control
# Run from repo root:
#   pyinstaller installer\rapid_gaussmeter.spec
#
# Output: dist\RapidPy_Gaussmeter.exe  (one-file bundle)
#
# FW Bell DLLs (usb5100.dll, libusb0.dll) and the x86 probe helper are
# bundled into the exe and extracted to sys._MEIPASS/tools/ at runtime.
# The probe subprocess finds its DLLs there via the standard Windows DLL
# search path (directory of the launched executable).

from pathlib import Path

REPO_ROOT  = Path(SPECPATH).parent          # E:\Github\RAPID
APP_DIR    = REPO_ROOT / "RapidPy" / "gaussmeter_control"
ICON_PATH  = APP_DIR / "assets" / "gaussmeter_icon.ico"
COMMON_DIR = REPO_ROOT / "RapidPy" / "rapidpy_common"
TOOLS_DIR  = REPO_ROOT / "tools"
LIB_DIR    = REPO_ROOT / "lib"

# Bundle FW Bell DLLs into tools\ so the probe subprocess finds them.
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
        (str(TOOLS_DIR / "usb5100_probe.exe"), "tools"),
    ],
    datas=[
        (str(COMMON_DIR / "assets"), "rapidpy_common/assets"),
        (str(APP_DIR / "assets"),    "assets"),
        *_fw_bell_dlls,
    ],
    hiddenimports=[
        "pyqtgraph",
        "pyqtgraph.exporters",
        "serial",
        "serial.tools",
        "serial.tools.list_ports",
        "rapidpy_common",
        "rapidpy_common.gaussmeter",
        "rapidpy_common.ui",
        "rapidpy_common.palette",
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
    a.binaries,
    a.datas,
    [],
    name="RapidPy_Gaussmeter",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    icon=str(ICON_PATH) if ICON_PATH.exists() else None,
)
