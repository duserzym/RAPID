# PyInstaller spec for RapidPy COM Port Mapper
# Run from repo root:
#   pyinstaller --noconfirm --clean installer\rapid_com_port_mapper.spec
#
# Output: dist\RapidPyCOMMapper.exe  (one-file bundle)

from pathlib import Path

REPO_ROOT  = Path(SPECPATH).parent          # E:\Github\RAPID
APP_DIR    = REPO_ROOT / "RapidPy" / "com_port_mapper"
ICON_PATH  = APP_DIR / "assets" / "com_port_mapper_icon.ico"
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
        (str(APP_DIR / "assets"),    "assets"),
    ],
    hiddenimports=[
        "serial",
        "serial.tools",
        "serial.tools.list_ports",
        "rapidpy_common",
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
    name="RapidPyCOMMapper",
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
