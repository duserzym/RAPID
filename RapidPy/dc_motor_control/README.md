# RapidPy DC Motor Control

Python subsystem app modeled after `VB6/frmDCMotors.frm`.

## Focus

- Multi-axis DC motor panel (`Changer X`, `Turning`, `Up/Down`, `Changer Y`)
- Position/speed moves with active-axis selection
- Turning motor spin command workflow
- Hole-to-position and position-to-hole conversion helpers

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
