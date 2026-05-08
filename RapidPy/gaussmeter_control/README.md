# RapidPy 908A Gaussmeter Control

Standalone PySide6 panel for the VB6 `frm908AGaussmeter` workflow.

## What It Does

- Connects to the legacy gaussmeter driver `gm0.dll`
- Supports USB auto mode and manual COM-port mode
- Reads live gaussmeter values with polling
- Exposes VB6-aligned controls for mode, units, range, null, auto-zero, peak reset, and time sync
- Reuses `rapidpy_common/gaussmeter.py` so the same source functions can be shared by the COM mapper and later subsystem ports

## Driver Requirement

This app does not talk directly to the instrument protocol from Python. It mirrors the VB6 architecture and requires `gm0.dll` to be available either:

- on `PATH`, or
- at the path stored in `RAPID_GM0_DLL`, or
- selected manually in the GUI

Without that DLL, the app will start but live communications remain unavailable.

## Run

```bash
cd RapidPy/gaussmeter_control
python -m pip install -r ..\..\requirements.txt
python main.py
```