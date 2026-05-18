# Changer XY Control Roadmap

This app has already been repurposed from a hole/sample list helper into a dedicated X/Y stage test, calibration, and visualization tool. The roadmap below tracks the current state and the next concrete steps.

## Current State

- Sample-list workflow removed from the main UI.
- Live desktop layout is now stage-centered:
   - left sidebar for connections, switch state, homing, and stage actions
   - large center stage panel
   - right sidebar for jog tuning, calibration, and cup-position editing
- The live 3D render panel remains removed so the app stays lightweight while core stage control and calibration behavior are refined.
- Independent COM port assignment and compact On/Off controls exposed for X, Y, Z, and reference hardware.
- Numeric jog controls exposed directly as velocities in counts per second.
- VB6-derived operator limit guidance surfaced in the motion controls via tooltips and help text.
- 2D tray model aligned to the photographed stage geometry, including staggered cup rows and center drop-off hole 46.
- Cup-position sheet now exposes editable X/Y counts for cups 1..100 inside the live changer app.
- Changer app now loads VB6-style INI settings directly on startup and surfaces the active file, motor IDs, and motion limits in the UI.
- Changer app can now save the current cup table back to a VB6-style INI, with internal `.rapidpy_history/<ini-stem>/` snapshots before overwrite.
- Separate `RapidPy/settings_editor` app now exists for section-organized INI editing, JSON import/export, and internal settings snapshot history.

## Verified VB6 Anchors

- XY homing uses full `ChangerSpeed` with negative and positive sweeps of `-30000000` and `30000000` counts.
- XY centering requires the Z / up-down axis to be homed to top first.
- Z homing logic uses the top switch on internal status bit 4.
- X limit status bits are 4 and 5.
- Y limit status bits are 5 and 6.
- Current Berkeley changer profile in `VB6/settings/Paleomag_v3.INI` sets:
   - XY max move speed: `10000000`
   - Z slow / normal / fast: `25000000 / 35000000 / 35000000`
   - changer X / Y / Z motor IDs: `16 / 16 / 16`
- The live changer UI now starts manual jogs lower than those file limits for safer operator use:
   - XY manual jog default: `2000000`
   - Z manual jog default: `2500000`
   - jog step default: derived from the loaded cup spacing, currently about `285` counts for the Berkeley profile

## Next Control / UX Steps

1. Add richer calibration diagnostics so cup offsets and nearest-hole matching can be audited quickly.
2. Decide whether the cup-position sheet should support bulk paste/import from CSV or stay intentionally manual.
3. Consider a lightweight handoff from the changer app into `RapidPy/settings_editor` for deeper settings-file edits.
4. Continue tightening operator-facing copy so the stage view stays visually dominant.

## Longer-Term Goal

Keep the changer UI lightweight for daily motion and calibration work while preserving the option to reintroduce a local in-app digital twin later, only if it can stay performant and clearly supports operator tasks.