# RAPID — RapidPy Paleomagnetic System

> **VB6 → Python 64-bit modernisation** of the UC Berkeley RAPID paleomagnetic magnetometer controller.

The original VB6 source is archived at [sourceforge.net/projects/paleomag](https://sourceforge.net/projects/paleomag/) (GPLv3). The `VB6/` folder in this repo contains the Berkeley-specific variant operated on the Hargraves magnetometer, upgraded 2024-2025 and deployed at the Institute for Rock Magnetism (University of Minnesota, 2025).

Modern replacements live under `RapidPy/` — self-contained Python apps with a shared hardware abstraction layer and a unified UMN maroon-and-gold UI style.

---

## Current Progress

RapidPy is now beyond the initial three-app milestone. The repository currently contains seven released modules, additional in-progress migration panels, and bench or support utilities, with public app pages under the GitHub Pages site and recent one-file Windows builds landing in the repo-root `dist/` folder.

- **Website overview:** https://duserzym.github.io/RAPID/
- **App pages:** https://duserzym.github.io/RAPID/apps/index.html
- **Manual index:** [docs/compiled-apps-manual-index.md](docs/compiled-apps-manual-index.md)

### Module Snapshot

| App | Status | Module | Current packaging | Documentation |
|---|---|---|---|---|
| Gaussmeter Control | Released | `RapidPy/gaussmeter_control` | `dist/RapidPy_Gaussmeter.exe` | [Manual](docs/gaussmeter-user-manual.md) · [App page](https://duserzym.github.io/RAPID/apps/index.html#gaussmeter-control) |
| VRM Decay Logger | Released | `RapidPy/vrm_logger` | `dist/RapidPyVRM.exe` | [Manual](docs/vrm-logger-user-manual.md) · [App page](https://duserzym.github.io/RAPID/apps/index.html#vrm-decay-logger) |
| ADwin Communications Tester | Released | `RapidPy/adwin_comms` | `dist/RapidPyADWin.exe` | [Manual](docs/adwin-comms-user-manual.md) · [App page](https://duserzym.github.io/RAPID/apps/index.html#adwin-comms) |
| COM Port Mapper | Released | `RapidPy/com_port_mapper` | `dist/RapidPyCOMMapper.exe` | [README](RapidPy/com_port_mapper/README.md) · [App page](https://duserzym.github.io/RAPID/apps/index.html#com-port-mapper) |
| XY Sample Changer | Released | `RapidPy/changer_xy_control` | `dist/RapidPyChangerXY.exe` | [Manual](docs/changer-xy-control-user-manual.md) · [App page](https://duserzym.github.io/RAPID/apps/index.html#changer-xy-control) |
| Up/Down Control | Released | `RapidPy/updown_control` | `dist/RapidPyUpDown.exe` | [Manual](docs/updown-control-user-manual.md) · [App page](https://duserzym.github.io/RAPID/apps/index.html#updown-control) |
| RapidPy DataViewer | Released utility | `RapidPy/data_viewer` | Source launcher today | [App page](https://duserzym.github.io/RAPID/apps/index.html#data-viewer) |
| AF Tuner | In progress | `RapidPy/af_tuner` | `dist/RapidPyAFTuner.exe` | [README](RapidPy/af_tuner/README.md) · [App page](https://duserzym.github.io/RAPID/apps/index.html#af-tuner) |
| RAPID v4 System Shell | In progress | `RapidPy/system_shell` | `dist/RapidPySystemShell.exe` | [README](RapidPy/system_shell/README.md) |
| DC Motor Control | Migration / bench | `RapidPy/dc_motor_control` | `dist/RapidPyDCMotors.exe` | [README](RapidPy/dc_motor_control/README.md) |
| AF Clip Test | Bench utility | `RapidPy/af_clip_test` | `dist/RapidPyAFClipTest.exe` | [README](RapidPy/af_clip_test/README.md) |

### Recent Highlights

- **COM Port Mapper:** released Windows utility for serial-role discovery and saved machine-local COM assignments before launching hardware apps.
- **XY Sample Changer:** released stage-centric operator panel with XY-only mode, in-app cup calibration editing, and VB6-compatible `HomeToCenter` / `MoveToCorner` behavior.
- **Up/Down Control:** released Z-axis panel combining raw motion, top-switch readback, vacuum control, and SQUID-guided MeasPos optimisation in one window.
- **RapidPy DataViewer:** released lightweight palaeomagnetic viewer for CIT `.sam`, MagIC `measurements.txt`, and legacy CSV data, with 2D/3D directional views, stereonet, intensity and paleointensity quicklooks, rule-based next-step hints, automatic file watching, and a generated app icon.
- **Docs and site:** GitHub Pages homepage, hash-routed app pages, refreshed screenshots, module icons, and a compiled-app manual index now track the public migration surface.

---

## Architecture

```
RapidPy/
├── rapidpy_common/          # Shared hardware layer
│   ├── ui.py                # UMN maroon/gold liquid-glass style sheet
│   ├── hardware.py          # Quicksilver motor protocol + motion helpers
│   ├── adwin_af.py          # ADwin AF ramp backend
│   └── gaussmeter.py        # gm0.dll / FW Bell wrapper + reading helpers
├── gaussmeter_control/      # Released FW Bell 908A operator panel
├── vrm_logger/              # Released three-axis SQUID VRM decay logger
├── adwin_comms/             # Released ADwin I/O diagnostic tester
├── com_port_mapper/         # Released serial-role discovery utility
├── af_tuner/                # In-progress AF coil tuning panel
├── af_clip_test/            # Bench utility for AF clipping workflows
├── changer_xy_control/      # Released stage-centric XY changer panel
├── updown_control/          # Released Z-axis / vacuum / MeasPos panel
├── dc_motor_control/        # Broader motor migration surface
├── system_shell/            # In-progress subsystem launcher
└── data_viewer/             # Standalone demagnetisation data viewer
```

---

## Build & Distribution

Recent Windows builds are one-file PyInstaller bundles written to the repo-root `dist/` folder:

| Executable | Description |
|---|---|
| `RapidPy_Gaussmeter.exe` | Gaussmeter Control |
| `RapidPyVRM.exe` | VRM Decay Logger |
| `RapidPyADWin.exe` | ADwin Communications Tester |
| `RapidPyCOMMapper.exe` | COM Port Mapper |
| `RapidPyChangerXY.exe` | XY Sample Changer |
| `RapidPyUpDown.exe` | Up/Down Control |
| `RapidPyAFTuner.exe` | AF Tuner |
| `RapidPySystemShell.exe` | RAPID v4 System Shell |
| `RapidPyDCMotors.exe` | DC Motor Control |
| `RapidPyAFClipTest.exe` | AF Clip Test |
| `install_fwbell_drivers.exe` | FW Bell USB driver installer (bundled with Gaussmeter Setup) |

`RapidPy/data_viewer` is currently maintained as a source-launched utility rather than a packaged executable; its app assets are generated from `RapidPy/data_viewer/tools/generate_icon.py`.

PyInstaller specs live under `installer/`. Most app folders provide a `build_windows.bat` that regenerates icons and invokes the matching spec from the repo root.

Build from an app folder with:

```powershell
Set-Location RapidPy\<module_folder>
conda run -n paleomag cmd /c build_windows.bat
```

Examples:

```powershell
Set-Location RapidPy\gaussmeter_control
conda run -n paleomag cmd /c build_windows.bat

Set-Location RapidPy\changer_xy_control
conda run -n paleomag cmd /c build_windows.bat

Set-Location RapidPy\updown_control
conda run -n paleomag cmd /c build_windows.bat
```

**Requirements:** `conda env create -f environment.yml` creates the `paleomag` environment (Python 3.13, PySide6, pyqtgraph, PyInstaller).

---

## Additional Documentation

| Document | Description |
|---|---|
| [compiled-apps-manual-index.md](docs/compiled-apps-manual-index.md) | Current compiled-app status, manual coverage, and VB6 transition-sheet tracking |
| [docs/site/index.html](docs/site/index.html) | Static source for the GitHub Pages homepage |
| [docs/site/apps/index.html](docs/site/apps/index.html) | Static source for the GitHub Pages app-detail pages |
| [fw-bell-gaussmeter-user-guide.md](docs/fw-bell-gaussmeter-user-guide.md) | Operator setup, driver installation, DLL placement, and GUI usage |
| [fw-bell-gaussmeter-developer-guide.md](docs/fw-bell-gaussmeter-developer-guide.md) | Direct driver work, helper usage, SCPI communication |
| [gaussmeter-user-manual.md](docs/gaussmeter-user-manual.md) | Full Gaussmeter Control app manual |
| [vrm-logger-user-manual.md](docs/vrm-logger-user-manual.md) | VRM Decay Logger full manual |
| [adwin-comms-user-manual.md](docs/adwin-comms-user-manual.md) | ADwin Communications Tester manual |
| [changer-xy-control-user-manual.md](docs/changer-xy-control-user-manual.md) | XY Sample Changer full manual |
| [updown-control-user-manual.md](docs/updown-control-user-manual.md) | Up/Down Control full manual |

---

## Licence

GPLv3 — see [LICENSE](LICENSE).

