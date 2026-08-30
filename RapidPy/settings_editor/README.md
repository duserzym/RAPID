# RapidPy Settings Editor

RapidPy Settings Editor is a standalone operator tool for viewing and editing RAPID settings files in a section-organized UI instead of a flat text editor.

## What It Does

- Loads VB6-compatible INI files such as `VB6/Defaults.ini`
- Organizes settings by section in a left-hand navigator
- Edits keys and values directly in a table
- Exports the current document to JSON for easier inspection and interchange
- Imports JSON back into the editor and writes it back out as INI
- Captures the previous INI into an internal snapshot history before every save

## Versioning Behavior

The editor uses an internal snapshot history instead of scattering timestamped `.bak` files beside the active INI.

Before every save, if the target INI already exists, the current file is copied to:

`.rapidpy_history/<ini-stem>/YYYY-MM-DD_HH-MM-SS.ini`

For example, saving `VB6/Defaults.ini` creates snapshots under:

`VB6/.rapidpy_history/Defaults/`

This keeps the working directory cleaner than the VB6 `filename_MM-DD-YYYY_HH-MM-SS.bak` pattern while still preserving a full-file checkpoint before each write. The app also exposes that snapshot list in the UI so an older revision can be loaded back into the editor and saved again.

## JSON Format

The editor exports JSON in this shape:

```json
{
  "sections": [
    {
      "name": "Boards",
      "entries": [
        {"key": "BoardsCount", "value": "2"}
      ]
    }
  ]
}
```

Section order and key order are preserved as they appear in the editor.

## Launch

From the repo root:

```powershell
c:\Users\Berkeley_QDM\anaconda3\envs\paleomag\python.exe RapidPy/settings_editor/main.py
```
