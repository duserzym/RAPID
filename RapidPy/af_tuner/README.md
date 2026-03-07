# RapidPy AF Tuner

Python subsystem app modeled after `VB6/frmAFTuner.frm`.

## Focus

- Axial and transverse coil tuning workflows
- Old/new resonance frequency and max voltage fields
- Separate `Apply` and `Save` behavior matching VB6 panel expectations
- ADWIN-backed auto-tune sweep (frequency scan + best-peak selection)
- Runtime-configurable ADWIN board/process parameters
- Runtime-configurable AF relay digital bit mapping (axial/transverse)
- Coil-specific max ramp/monitor limits fed directly into ADWIN `FPAR` settings

## Run

```bash
cd RapidPy/af_tuner
python -m pip install -r requirements.txt
python main.py
```

## Windows Build

```bat
cd RapidPy\af_tuner
build_windows.bat
```

Notes:

- Uses shared styling from `RapidPy/rapidpy_common/ui.py`
- Stores values in `~/.rapidpy_af_tuner.json`
- Stores backend settings in `~/.rapidpy_af_backend.json`
- Stores sweep settings in `~/.rapidpy_af_autotune.json`
- Requires Windows with `adwin32.dll` available on `PATH` for live auto-tune execution
- Applies AF relay switching on coil selection and before auto-tune starts
