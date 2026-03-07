# RapidPy VRM Logger

Lightweight VRM decay logger and live viewer for 3-axis SQUID output, based on VB6 behavior from `frmSquid.frm` and `frmVRM.frm`.

## Features

- Cross-platform GUI (macOS and Windows) using `PySide6`
- Fast live plotting via `pyqtgraph`
- SQUID serial protocol at `1200,N,8,1`
- User-defined sampling interval in seconds
- `Linear` spacing (fixed interval) or `Log` spacing (multiplies wait each step)
- Display units as `Volts` or calibrated `Moment`
- CSV logging with time, raw volts, displayed values, and unit
- Lower-left console panel with timestamped runtime status messages
- Auto-scaling live graph on both axes during streaming
- Config persistence in `~/.rapidpy_vrm_config.json`
- Custom `VRM` icon for app window and compiled executables

## Visual Theme

Applied palette (UMN-inspired maroon and gold):

- `Maroon: #7A0219`
- `Gold: #FFCD34`

UI styling also prioritizes macOS-like typography and geometry:

- `SF Pro` family first, with graceful fallback to `Avenir Next`
- Rounded cards/controls and soft spacing for a native macOS feel
- Glass-like translucent cards, layered shadows, and glossy accents inspired by current Apple liquid-glass styling

## Serial Read Logic (matching VB6)

Per sample:

1. Latch all axes count: `ALC`
2. Latch all axes data: `ALD`
3. For each axis (`X`, `Y`, `Z`):
   - Read count: `<axis>SC`
   - Read data: `<axis>SD`
4. Compute axis value as `-(data + count)`

This mirrors `frmSQUID.getVal` with flux count mode active.

## Install

```bash
cd RapidPy/vrm_logger
python -m pip install -r requirements.txt
```

## Run

```bash
cd RapidPy/vrm_logger
python main.py
```

## Build Executables

### macOS

```bash
cd RapidPy/vrm_logger
chmod +x build_macos.sh
./build_macos.sh
```

The build script auto-generates icon assets and uses `assets/vrm_icon.icns` when available.

### Windows

```bat
cd RapidPy\vrm_logger
build_windows.bat
```

The build script auto-generates icon assets and packages with `assets/vrm_icon.ico`.

## Icon Assets

You can regenerate icon files any time:

```bash
cd RapidPy/vrm_logger
python tools/generate_icon.py
```

This creates:

- `assets/vrm_icon.svg`
- `assets/vrm_icon.png`
- `assets/vrm_icon.ico`
- `assets/vrm_icon.icns` (on macOS when `sips` and `iconutil` are available)

## Calibration Notes

Default factors come from VB6 defaults in `modConfig.bas`:

- `X = -3.410`
- `Y = -3.470`
- `Z = -2.516`

Set display units to `Moment` to apply these factors. Keep `Volts` for raw outputs.
