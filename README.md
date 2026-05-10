# RAPID — RapidPy Paleomagnetic System

> **VB6 → Python 64-bit modernisation** of the UC Berkeley RAPID paleomagnetic magnetometer controller.

The original VB6 source is archived at [sourceforge.net/projects/paleomag](https://sourceforge.net/projects/paleomag/) (GPLv3). The `VB6/` folder in this repo contains the Berkeley-specific variant operated on the Hargraves magnetometer, upgraded 2024-2025 and deployed at the Institute for Rock Magnetism (University of Minnesota, 2025).

Modern replacements live under `RapidPy/` — self-contained Python apps with a shared hardware abstraction layer and a unified UMN maroon-and-gold UI style.

---

## RapidPy Modules

### Gaussmeter Control

![Gaussmeter Control](docs/images/gaussmeter_app.png)

A driver-backed operator panel for the FW Bell 908A USB gaussmeter.  
Supports live field readings, configurable sampling sessions with alarm thresholds, and a snap-to-point hover tooltip on the session plot.  
Ships with a bundled FW Bell USB driver installer (`install_fwbell_drivers.exe`).

- **Backend:** `gm0.dll` (legacy 32-bit via ctypes) **or** FW Bell `usb5100.dll` (64-bit)
- **Connection:** RS-232 / COM port or USB auto
- **Output:** timestamped CSV session export
- **Docs:** [Gaussmeter User Manual](docs/gaussmeter-user-manual.md) · [Developer Guide](docs/fw-bell-gaussmeter-developer-guide.md)

---

### VRM Decay Logger

![VRM Decay Logger](docs/images/vrm_logger_app.png)

Three-axis SQUID live-view logger for viscous remanent magnetisation decay experiments.  
Reads X/Y/Z moment via serial, applies per-axis calibration, and streams live curves with configurable logarithmic or linear time steps.  
Calibration constants can be loaded directly from the legacy `.INI` file used by the VB6 system.

- **Backend:** RS-232 SQUID serial stream
- **Calibration:** load from legacy `Paleomag_v3.INI` or enter manually
- **Output:** timestamped CSV with X/Y/Z columns
- **Docs:** [VRM Logger User Manual](docs/vrm-logger-user-manual.md)

---

### ADwin Communications Tester

A diagnostic panel for testing ADwin digital/analog I/O before running AF or DC motor routines.  
Provides channel-by-channel read/write controls and a live square-wave output visualisation.

- **Backend:** ADwin Gold II via `ADwin.bas`-compatible Python bindings
- **Purpose:** bench-test ADwin comms without launching the full RAPID orchestrator
- **Docs:** [ADwin Comms Manual](docs/adwin-comms-user-manual.md)

---

## Architecture

```
RapidPy/
├── rapidpy_common/          # Shared hardware layer
│   ├── ui.py                # UMN maroon/gold liquid-glass style sheet
│   ├── hardware.py          # Quicksilver motor protocol + motion helpers
│   ├── adwin_af.py          # ADwin AF ramp backend
│   └── gaussmeter.py        # gm0.dll / FW Bell wrapper + reading helpers
├── gaussmeter_control/      # FW Bell 908A operator panel
├── vrm_logger/              # Three-axis SQUID VRM decay logger
├── adwin_comms/             # ADwin I/O diagnostic tester
├── af_tuner/                # AF coil tuning panel
├── changer_xy_control/      # XY sample changer queue and controls
├── updown_control/          # Vertical axis (up/down) motor panel
├── dc_motor_control/        # General motor panel
└── system_shell/            # Operator launcher for all subsystem panels
```

---

## Build & Distribution

All Windows executables are one-file PyInstaller bundles output to `dist/`:

| Executable | Description |
|---|---|
| `RapidPy_Gaussmeter.exe` | Gaussmeter Control |
| `RapidPyVRM.exe` | VRM Decay Logger |
| `RapidPyADWin.exe` | ADwin Communications Tester |
| `install_fwbell_drivers.exe` | FW Bell USB driver installer (bundled with Gaussmeter Setup) |

Build from the repo root with:

```powershell
# Gaussmeter
Set-Location RapidPy\gaussmeter_control ; conda run -n paleomag cmd /c build_windows.bat

# VRM Logger
Set-Location RapidPy\vrm_logger ; conda run -n paleomag cmd /c build_windows.bat

# ADwin Comms
Set-Location RapidPy\adwin_comms ; conda run -n paleomag cmd /c build_windows.bat
```

**Requirements:** `conda env create -f environment.yml` → activates `paleomag` environment (Python 3.13, PySide6, pyqtgraph, PyInstaller).

---

## Additional Documentation

| Document | Description |
|---|---|
| [fw-bell-gaussmeter-user-guide.md](docs/fw-bell-gaussmeter-user-guide.md) | Operator setup, driver installation, DLL placement, and GUI usage |
| [fw-bell-gaussmeter-developer-guide.md](docs/fw-bell-gaussmeter-developer-guide.md) | Direct driver work, helper usage, SCPI communication |
| [gaussmeter-user-manual.md](docs/gaussmeter-user-manual.md) | Full Gaussmeter Control app manual |
| [vrm-logger-user-manual.md](docs/vrm-logger-user-manual.md) | VRM Decay Logger full manual |
| [adwin-comms-user-manual.md](docs/adwin-comms-user-manual.md) | ADwin Communications Tester manual |

---

## Licence

GPLv3 — see [LICENSE](LICENSE).

