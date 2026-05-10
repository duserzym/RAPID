# PyInstaller spec for the standalone FW Bell driver installer
# Run from repo root:
#   pyinstaller installer\install_fwbell_drivers.spec
#
# Output: dist\install_fwbell_drivers\install_fwbell_drivers.exe
#
# The spec bundles usb5100.dll + libusb0.dll from the repo lib\ folder into a
# drivers\ subfolder inside the exe bundle so the installer can find and copy
# them without the user needing to locate them manually.

from pathlib import Path

REPO_ROOT = Path(SPECPATH).parent          # e.g. E:\Github\RAPID
LIB_DIR   = REPO_ROOT / "lib"

# Collect DLLs to bundle.  Both files must be present in lib\.
_dlls_to_bundle = []
for _fname in ("usb5100.dll", "libusb0.dll"):
    _p = LIB_DIR / _fname
    if _p.exists():
        _dlls_to_bundle.append((str(_p), "drivers"))

a = Analysis(
    [str(REPO_ROOT / "installer" / "install_fwbell_drivers.py")],
    pathex=[str(REPO_ROOT)],
    binaries=[],
    datas=_dlls_to_bundle,
    hiddenimports=[
        "tkinter",
        "tkinter.filedialog",
        "tkinter.messagebox",
    ],
    hookspath=[],
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
    name="install_fwbell_drivers",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,          # No console window for the GUI path
    uac_admin=True,         # Request elevation via manifest
    icon=None,
)

coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name="install_fwbell_drivers",
)
