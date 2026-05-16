# RapidPy Changer XY Control

RapidPy Changer XY Control is now a lightweight operator panel for the RAPID X/Y stage, built around the VB6 changer behavior but focused on motion, homing, switch state, and cup calibration rather than sample-list management.

## Current UI

- Large center stage cartoon based on the photographed acrylic holder
- Left-side hardware cards for COM selection, connection state, homing, and stage control
- Right-side cards for jog velocities, calibration capture, and a compact editable cup-position sheet
- Single-click cup selection and double-click move-to-cup interaction
- VB6-correct limit switch interpretation, including the Z-top switch

## Calibration Workflow

- Capture the current X/Y counts into a cup calibration
- Edit cup X/Y counts directly in the built-in position sheet
- Clear a cup by leaving both X and Y cells blank
- Use the stage view and the table together to inspect and refine stored positions

## Related Settings Tool

The companion settings editor now lives at `RapidPy/settings_editor/`.
Use it to browse and edit VB6-compatible INI files such as `VB6/Defaults.ini`, with JSON import/export and internal snapshot history on every save.

## Run

```powershell
c:\Users\Berkeley_QDM\anaconda3\envs\paleomag\python.exe RapidPy/changer_xy_control/main.py
```

## Notes

- Uses shared styling from `RapidPy/rapidpy_common/ui.py`
- Stage motion and switch semantics are checked against `VB6/frmDCMotors.frm`
- Settings-file compatibility work is centered in the separate settings editor app
