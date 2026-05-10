# PyInstaller spec for RapidPy ADwin Communication Tester
# Run from repo root:
#   pyinstaller installer\rapid_adwin_comms.spec
#
# Output: dist\RapidPyADWin.exe  (one-file bundle)
#
# NOTE: adwin32.dll is NOT bundled — it must be installed on the target PC
# via the ADwin driver package from Jager Messtechnik.

from pathlib import Path

REPO_ROOT  = Path(SPECPATH).parent          # E:\Github\RAPID
APP_DIR    = REPO_ROOT / "RapidPy" / "adwin_comms"
COMMON_DIR = REPO_ROOT / "RapidPy" / "rapidpy_common"
ICON_PATH  = APP_DIR / "assets" / "adwin_icon.ico"

a = Analysis(
    [str(APP_DIR / "main.py")],
    pathex=[
        str(APP_DIR),
        str(REPO_ROOT / "RapidPy"),
        str(REPO_ROOT),
    ],
    binaries=[],
    datas=[
        (str(COMMON_DIR / "assets"), "rapidpy_common/assets"),
    ],
    hiddenimports=[
        "pyqtgraph",
        "pyqtgraph.exporters",
        "rapidpy_common",
        "rapidpy_common.adwin_af",
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
    name="RapidPyADWin",
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
