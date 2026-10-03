# Build with installer/build_rapid_main.ps1. Both entry points share one payload.
from pathlib import Path
import sys
from PyInstaller.utils.hooks import collect_data_files, collect_submodules, copy_metadata

ROOT = Path(SPECPATH).parent
PACKAGES = ROOT / ".tmp" / "portable-packages"
if not (PACKAGES / "rapid_main").is_dir():
    raise RuntimeError("Run installer/build_rapid_main.ps1 to stage the release wheel first.")
sys.path.insert(0, str(PACKAGES))
modules = ["rapid_main", "rapidpy_common", "af_tuner", "data_viewer",
           "gaussmeter_control", "updown_control", "vrm_logger", "webcam_viewer"]
assets = ["rapid_main_assets", "af_tuner_assets", "data_viewer_assets",
          "gaussmeter_control_assets", "updown_control_assets", "vrm_logger_assets", "webcam_viewer_assets"]
hiddenimports = [sub for module in modules for sub in collect_submodules(module)] + assets
datas = [(str(ROOT / "LICENSE"), ".")]
for module in modules + assets:
    datas += collect_data_files(module)
datas += copy_metadata("berkeley-rapidpy")

a = Analysis([str(ROOT / "installer" / "rapid_main_entry.py")],
             pathex=[str(PACKAGES)], binaries=[], datas=datas,
             hiddenimports=hiddenimports, hookspath=[], hooksconfig={},
             runtime_hooks=[], excludes=["pytest", "IPython", "notebook"], noarchive=False)
# Qt uses Windows' unversioned ICU API. Build-host PATH entries for Poppler
# supplied an incompatible ICU 78 DLL; libheif also supplied private CRT/API
# shims. Do not let those tooling runtimes shadow Windows 10 system libraries.
def is_host_system_shadow(entry):
    name = Path(entry[0]).name.lower()
    return (name in {"icuuc.dll", "icudt78.dll", "ucrtbase.dll"}
            or name.startswith("api-ms-win-"))
a.binaries = [entry for entry in a.binaries if not is_host_system_shadow(entry)]
pyz = PYZ(a.pure)
icon = str(PACKAGES / "rapid_main_assets" / "rapid_main_window_icon.ico")
gui = EXE(pyz, a.scripts, [], exclude_binaries=True, name="RapidPyMain",
          console=False, icon=icon, upx=False)
cli = EXE(pyz, a.scripts, [], exclude_binaries=True, name="RapidPyMainConsole",
          console=True, icon=icon, upx=False)
coll = COLLECT(gui, cli, a.binaries, a.datas, name="RapidPyMain", strip=False, upx=False)
