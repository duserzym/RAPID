# RapidPy DC Motor Control

Python subsystem app modeled after `VB6/frmDCMotors.frm`.

## Focus

- Multi-axis DC motor panel (`Changer X`, `Turning`, `Up/Down`, `Changer Y`)
- Position/speed moves with active-axis selection
- Turning motor spin command workflow
- Hole-to-position and position-to-hole conversion helpers
- Background telemetry for target position, actual encoder position, position error, and both controller velocity filters
- Live actual-torque readback from the low word of Quicksilver register 9, displayed in native controller units and percent full scale
- Efficient rolling pyqtgraph traces that remain responsive during blocking move and homing workflows
- Scrollable controls and active-monitor size clamping for smaller or mixed-resolution Windows displays

## Run

```bash
cd RapidPy/dc_motor_control
python -m pip install -r requirements.txt
python main.py
```

## Windows Build

```bat
cd RapidPy\dc_motor_control
build_windows.bat
```

Notes:

- Uses shared hardware wrapper from `RapidPy/rapidpy_common/hardware.py`
- Conversion utilities are aligned with VB6 concepts (`ConvertHoletoPos`, `ConvertPosToHole`).
- Telemetry follows the SilverLode dedicated-register map. Torque is an estimated controller value, not a calibrated precision torque measurement.
