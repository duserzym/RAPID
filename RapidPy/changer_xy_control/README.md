# RapidPy Changer XY Control

Python subsystem app modeled after `VB6/frmChanger.frm` with XY-focused operations.

## Focus

- Hole/sample list editor (1..100)
- Queue option controls (ascending/descending, holder repeat, return-to-start)
- CSV import/export for sample assignments
- Software-level goto-hole workflow to support operator transitions from VB6

## Run

```bash
cd RapidPy/changer_xy_control
python -m pip install -r requirements.txt
python main.py
```

## Windows Build

```bat
cd RapidPy\changer_xy_control
build_windows.bat
```

Notes:

- Uses shared styling from `RapidPy/rapidpy_common/ui.py`
- Next iteration can attach this panel directly to motor move commands.
