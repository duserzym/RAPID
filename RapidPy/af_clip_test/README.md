# RapidPy AF Clip Test

Python subsystem app modeled after the clipping-test half of `VB6/frmAFTuner`.

## Focus

- VB6-style auto clipping test with ramp-up and ramp-down passes
- Shared ADwin backend configuration and relay mapping used by `af_tuner`
- Sine-fit monitor amplitude plus residual RMS versus ramp voltage plotting
- Highest-voltage waveform preview using the shared board-timed capture path
- Save-back of active coil max ramp/max monitor limits into the shared AF tuner config
- RAPID liquid-glass styling, contrast rules, and lightweight pyqtgraph plotting

## Run

```bash
cd RapidPy/af_clip_test
python -m pip install -r requirements.txt
python main.py
```

## Windows Build

```bat
cd RapidPy\af_clip_test
build_windows.bat
```

Notes:

- Uses shared styling from `RapidPy/rapidpy_common/ui.py`
- Stores shared coil limits in `~/.rapidpy_af_tuner.json`
- Stores shared backend settings in `~/.rapidpy_af_backend.json`
- Stores clip scan settings in `~/.rapidpy_af_clip_test.json`
- Requires Windows with `adwin32.dll` / `adwin64.dll` available for live execution
- Uses `VB6/ADwin/sineout.T91` for all board-timed clipping captures