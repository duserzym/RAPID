# PyInstaller spec for RapidPy AF Clip Test
# Run from repo root:
#   pyinstaller installer\rapid_af_clip_test.spec
#
# Output: dist\RapidPyAFClipTest.exe  (one-file bundle)

from pathlib import Path

REPO_ROOT = Path(SPECPATH).parent
APP_DIR = REPO_ROOT / "RapidPy" / "af_clip_test"
COMMON_DIR = REPO_ROOT / "RapidPy" / "rapidpy_common"
ICON_PATH = APP_DIR / "assets" / "af_clip_test_icon.ico"


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
    name="RapidPyAFClipTest",
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
