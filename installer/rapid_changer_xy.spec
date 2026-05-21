# PyInstaller spec for RapidPy Changer XY Control
# Run from repo root:
#   pyinstaller installer\rapid_changer_xy.spec
#
# Output: dist\RapidPyChangerXY.exe  (one-file bundle)

from pathlib import Path

REPO_ROOT = Path(SPECPATH).parent
APP_DIR = REPO_ROOT / "RapidPy" / "changer_xy_control"
COMMON_DIR = REPO_ROOT / "RapidPy" / "rapidpy_common"
ICON_PATH = APP_DIR / "assets" / "changer_xy_control_icon.ico"


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
        (str(APP_DIR / "assets"), "assets"),
    ],
    hiddenimports=[],
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
    name="RapidPyChangerXY",
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
