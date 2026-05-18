# RapidPy Changer XY Control — User Manual

This manual covers the current Python replacement for the RAPID X/Y sample changer panel. The app follows the VB6 changer behavior closely, but exposes the machine state in a more operator-friendly way with live switch readback, editable cup positions, snapshot-backed settings save, and raw-velocity helper text translated into estimated cm/s.

![RAPID changer tray overview](https://raw.githubusercontent.com/duserzym/RAPID/main/resources/x-y_stage.png)

The published website currently uses the high-resolution tray photograph above as the visual reference for this app page. It matches the physical geometry the Python changer UI is built around and avoids shipping a broken image link while published UI screenshots are still being captured.

---

## Table of Contents

1. [Purpose](#1-purpose)
2. [Launch](#2-launch)
3. [Window Overview](#3-window-overview)
4. [Connections and Hardware Roles](#4-connections-and-hardware-roles)
5. [Current Position and Switch Interpretation](#5-current-position-and-switch-interpretation)
6. [Velocity Inputs and Physical Speed Estimates](#6-velocity-inputs-and-physical-speed-estimates)
7. [Jogging, Centering, and Load Position](#7-jogging-centering-and-load-position)
8. [Cup Calibration Table](#8-cup-calibration-table)
9. [INI and JSON Settings Workflow](#9-ini-and-json-settings-workflow)
10. [Files Written by the App](#10-files-written-by-the-app)
11. [Troubleshooting](#11-troubleshooting)

---

## 1. Purpose

RapidPy Changer XY Control is a lightweight operator panel for the RAPID changer stage. It focuses on:

- connecting to the X, Y, Z, and reference/sensor serial lines
- reading switch state and estimating current stage position
- jogging and homing the stage with conservative startup defaults
- viewing and editing the calibrated cup position table
- saving settings back to a VB6-compatible INI or exchanging settings through JSON

The VB6 behavioral reference for XY changer motion is `VB6/frmDCMotors.frm`, not `frmChanger.frm`.

---

## 2. Launch

From the repository root:

```powershell
conda activate paleomag
python RapidPy/changer_xy_control/main.py
```

The app starts by loading `VB6/settings/Paleomag_v3.INI` unless another settings file was selected previously.

---

## 3. Window Overview

The window is organized around a large central stage cartoon and surrounding operator cards.

- Left side: connections, loaded settings profile, and main stage actions
- Center: stage view with cup layout, target selection, current position indication, and special load/center targets
- Right side: raw jog velocity inputs, live estimated cm/s text, calibration capture, and editable cup-position sheet

Single-click selects a cup or special target. Double-click activates motion to that target.

---

## 4. Connections and Hardware Roles

The app separates the machine roles so they can be diagnosed independently.

Typical Berkeley defaults currently used in this project are:

- X motor: `COM4`
- Y motor: `COM5`
- Z / up-down motor: `COM6`
- reference / light-gauge input: machine-dependent

The app keeps COM assignments in the local machine config file and does not overwrite them when you load an INI. This is intentional: the INI describes the stage and controller configuration, while COM mappings are workstation-specific.

---

## 5. Current Position and Switch Interpretation

The changer has distinct load and center switch states on both X and Y axes. The app reads those states directly and keeps session anchors so that once a switch position has been observed in the current session, the UI can continue to infer current position even when only part of the machine is connected.

Important operator behavior:

- `Read current position` polls the live controller state and updates the position readout
- if load or center switches have been observed during the session, the app can use those anchors to estimate current X/Y position
- the lower-left standalone `LOAD` marker lights when both load switches are active or when the inferred stage position matches the learned load anchor
- the center drop-off hole can be selected directly from the stage view

---

## 6. Velocity Inputs and Physical Speed Estimates

This is the most important clarification for interpreting the current UI.

The numeric velocity inputs are still the raw VB6-compatible controller values. They are not direct stage position counts per second.

The controller scale is derived from the loaded settings file:

- `TurningMotorFullRotation`
- `TurningMotor1rps`

The app computes:

```text
raw velocity scale = TurningMotor1rps / abs(TurningMotorFullRotation)
```

For the default Berkeley profile in `VB6/settings/Paleomag_v3.INI`:

```text
TurningMotorFullRotation = -8000
TurningMotor1rps         = 16000000
raw velocity scale       = 2000
```

That means:

- `2000` raw velocity units ≈ `1` position-count-per-second
- `1,000,000` raw command ≈ `500` position counts per second

The app then combines that controller-scale conversion with the physical geometry:

- XY motion uses the loaded `XYTable` cup spacing to estimate counts per centimeter
- Z motion uses `UpDownMotor1cm` from the loaded INI

This lets the UI show estimated physical speed in cm/s without breaking compatibility with the legacy raw controller values.

### Why the old estimate looked absurd

Adjacent cups differ by only a few thousand XY position counts, roughly around one inch of physical tray spacing. If a raw command such as `1,000,000` were interpreted directly as `1,000,000` position counts per second, the stage would appear to be moving hundreds or thousands of cm/s, which is physically impossible for this mechanism. The missing raw-to-position scale factor was the reason for the inflated estimate.

### Speed confirmation

The app warns before very high operator-entered speeds. This is based on the estimated physical cm/s rather than only the raw controller value, so the confirmation better matches the actual machine risk.

---

## 7. Jogging, Centering, and Load Position

The stage view includes two special targets besides the numbered cups:

- `LOAD` marker at the lower-left of the stage panel
- center drop hole

Operator interactions:

- single-click selects the target
- double-click `LOAD` triggers the VB6-style move-to-corner routine
- double-click the center hole triggers the VB6-style home-to-center routine

The jog panel uses conservative startup defaults even when the loaded INI allows much higher motion limits.

---

## 8. Cup Calibration Table

The cup table is live and directly tied to the current stage model.

- capture the current X/Y position into the selected cup
- edit X/Y counts directly in the table
- clear a cup by leaving both coordinates blank
- save the current map back to INI or export it through JSON

Nearest-hole matching and visual stage selection both use this calibration map.

---

## 9. INI and JSON Settings Workflow

The app supports both legacy and transition-era settings workflows.

### INI

- loads VB6-compatible changer settings
- reads motor IDs, motion defaults, XY table positions, and `UpDownMotor1cm`
- saves directly back to a VB6-style INI
- snapshots the previous file before overwrite

### JSON

- exports section-based JSON compatible with the RapidPy settings editor
- imports that JSON back into the changer app
- useful while gradually transitioning away from direct INI-only editing

---

## 10. Files Written by the App

The app writes several categories of local state:

- machine-local config: `~/.rapidpy_changer_xy_control.json`
- snapshot history for overwritten INIs: `.rapidpy_history/<ini-stem>/`
- saved INI or JSON settings chosen by the operator

The local machine config stores serial-port preferences and operator state. It is intentionally separate from stage settings files.

---

## 11. Troubleshooting

### The stage reads switches but position still looks wrong

- use `Read current position` after connecting both X and Y
- verify the correct settings file is loaded so the app has the proper `XYTable`
- ensure the session has observed a load or center switch state if position inference is being used

### The raw speed number looks huge

- that is expected; the field is a controller raw value, not direct cm/s
- use the helper text under the velocity input for the physical estimate
- if the estimate still looks unreasonable, confirm the loaded INI has the expected `TurningMotor1rps`, `TurningMotorFullRotation`, and `UpDownMotor1cm` values

### Save and load behavior does not match COM assignments

- COM ports are intentionally machine-local and are not overwritten by INI imports
- use the connections panel to set the machine-specific ports on each workstation

### A previous settings file needs to be restored

- look under `.rapidpy_history/<ini-stem>/` next to the overwritten INI
- each overwrite creates a timestamped snapshot before saving the new file