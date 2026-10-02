# RapidPy Subsystem Apps

RapidPy now contains multiple subsystem control apps to support a staged VB6 -> Python transition.

## Apps

- `com_port_mapper`: COM sweep and RAPID device identification helper for Windows serial adapters
- `gaussmeter_control`: 908A gaussmeter operator panel backed by the legacy `gm0.dll` driver
- `vrm_logger`: VRM logging and SQUID live plotting
- `af_tuner`: AF coil tuning panel based on `frmAFTuner`
- `af_clip_test`: AF clipping-test panel based on the clipping workflow inside `frmAFTuner`
- `changer_xy_control`: hole/sample list and queue prep based on `frmChanger`
- `updown_control`: vertical axis controls based on `frmDCMotors` up/down panel sections
- `dc_motor_control`: general motor panel based on `frmDCMotors`
- `system_shell`: operator launcher for all converted subsystem panels

## Shared Layer

- `rapidpy_common/ui.py`: shared UMN maroon/gold liquid-glass style
- `rapidpy_common/hardware.py`: VB6-aligned Quicksilver motor protocol, movement, and conversion utilities
- `rapidpy_common/adwin_af.py`: ADWIN AF ramp backend (`boot/load/set params/start/readback`)
- `rapidpy_common/gaussmeter.py`: shared `gm0.dll` wrapper, reading conversion helpers, and gaussmeter probe entry points

Implemented parity highlights:

- Up/down `HomeToTop` switch-stop behavior and sample pickup/dropoff torque choreography
- XY table `HomeToCenter` and `MoveToCorner` limit-switch guided routines
- AF relay selection through ADWIN digital output bit setting before AF ramps

## Transition Plan

1. Keep each subsystem app familiar and standalone for operators.
2. Validate protocol-accurate hardware behavior on Windows against machine limits and safety interlocks.
3. Merge app backends into a single orchestrated control app once workflows are validated.

## Install and verify the integrated application

Python 3.11 or newer is required. From the `RapidPy` directory, install the
distribution and its declared runtime dependencies:

```powershell
python -m pip install .
python -m rapid_main --check-startup
rapid-main
```

`--check-startup` is read-only and does not open the UI or connect to hardware.
It reports required and optional modules, the packaged icon path, and the exact
configuration file selected by `RAPID_CONFIG` (or
`%USERPROFILE%\.rapid\config.json`). A malformed existing configuration is a
startup blocker instead of being silently treated as defaults.

The wheel also installs console entry points for AF Tuner, Data Viewer,
Gaussmeter Control, Up/Down Control, VRM Logger, and Webcam Viewer. The main app
prefers checkout scripts during development and automatically uses these
installed modules in a packaged environment.

To build and inspect the same wheel used by clean-environment verification:

```powershell
python -m pip wheel --no-deps --no-build-isolation --wheel-dir dist .
python -m pip install --no-deps --target clean-install dist\berkeley_rapidpy-*.whl
```

Install normally (without `--no-deps`) for an operator machine. The
`--no-deps` form is intended only for verification in an environment where the
declared dependencies are already present.
