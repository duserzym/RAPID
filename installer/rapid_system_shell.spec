# PyInstaller spec for RapidPy System Shell
# Run from repo root:
#   pyinstaller installer\rapid_system_shell.spec
#
# Output: dist\RapidPySystemShell.exe  (one-file bundle)

from pathlib import Path

REPO_ROOT = Path(SPECPATH).parent
APP_DIR = REPO_ROOT / "RapidPy" / "system_shell"
COMMON_DIR = REPO_ROOT / "RapidPy" / "rapidpy_common"


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
    name="RapidPySystemShell",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
)
