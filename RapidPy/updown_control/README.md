# RapidPy Up/Down Control

Python subsystem app modeled after Up/Down controls in `VB6/frmDCMotors.frm`.

## Focus

- COM port dropdown with manual refresh
- Direct target entry with selectable units (`mm`, `cm`, `raw count`)
- Editable min/max raw-count bounds with out-of-range validation
- Preset actions with editable raw-count values:
	- `Sample Pickup`
	- `Sample Dropoff`
	- `Susceptibility Meter` (VB6 `SCoilPos`-style preset)
- Settings load/save to JSON and automatic initialization from local settings file
- Home-to-top switch-based homing

## Run

```bash
cd RapidPy/updown_control
python -m pip install -r requirements.txt
python main.py
```

## Windows Build

```bat
cd RapidPy\updown_control
build_windows.bat
```

Notes:

- Uses shared hardware wrapper from `RapidPy/rapidpy_common/hardware.py`
- Default local settings path: `~/.rapidpy_updown_settings.json`
