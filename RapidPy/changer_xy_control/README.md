# RapidPy Changer XY Control

RapidPy Changer XY Control is now a lightweight operator panel for the RAPID X/Y stage, built around the VB6 changer behavior but focused on motion, homing, switch state, and cup calibration rather than sample-list management.

## Velocity Clarification

The changer keeps the VB6-style raw velocity inputs in the UI, but the live helper text now explains those values in physical terms.

- The raw XY and Z velocity values used by the controller are not the same unit as stage position counts.
- The controller scale is derived from the loaded VB6 settings file as `TurningMotor1rps / abs(TurningMotorFullRotation)`.
- In `VB6/settings/Paleomag_v3.INI`, that is `16,000,000 / 8,000 = 2,000`, so `2,000` raw units correspond to roughly `1` position-count-per-second.
- That means a raw command such as `1,000,000` should be interpreted as about `500` position counts per second, not `1,000,000` position counts per second.
- The app combines that controller scale with the loaded `XYTable` cup spacing and `UpDownMotor1cm` so operators can keep VB6-compatible values while seeing estimated cm/s motion in the panel.

This matters because adjacent cups are only separated by a few thousand position counts. Without the raw-to-position conversion factor, the old estimate overstated physical speed by orders of magnitude.

## Current UI

- Large center stage cartoon based on the photographed acrylic holder
- Left-side hardware cards for COM selection, connection state, homing, and stage control
- Right-side cards for jog velocities, calibration capture, and a compact editable cup-position sheet
- Single-click cup selection and double-click move-to-cup interaction
- VB6-correct limit switch interpretation, including the Z-top switch
- Default startup profile loaded from `VB6/settings/Paleomag_v3.INI`
- Active profile badges showing the loaded INI name, X/Y/Z motor IDs, and VB6 motion limits
- Built-in `Save INI` / `Save INI As` actions that write the current cup table back to a VB6-style settings file

## Calibration Workflow

- Capture the current X/Y counts into a cup calibration
- Edit cup X/Y counts directly in the built-in position sheet
- Clear a cup by leaving both X and Y cells blank
- Use the stage view and the table together to inspect and refine stored positions
- Save the current cup table back into the active INI or a new INI copy directly from the changer app

## Settings Round-Trip

- The changer app now reads motor IDs, motion limits, and the full `XYTable` directly from a VB6-style INI
- Local COM selections remain machine-specific and are not overwritten by the INI loader
- Overwriting an INI creates an internal snapshot under `.rapidpy_history/<ini-stem>/` before the new file is written
- Manual jog controls start below the VB6 maximum motion limits so the app opens in a safer operator state

## Documentation

- [Full manual](../../docs/changer-xy-control-user-manual.md)
- [Website app page](../../docs/site/apps/index.html#changer-xy-control)

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
- Raw velocity conversion is derived from the loaded VB6 settings rather than hard-coding a fixed cm/s assumption
