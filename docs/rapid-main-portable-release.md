# RapidPy main application: portable pilot release

The glass workspace includes Dashboard, Sample Queue, Sequence, Live Measure,
Settings, and Calibration Center. Its Windows bundle also includes AF Tuner,
Data Viewer, Gaussmeter, Up/Down, VRM, and Webcam helpers. No separate Python
installation or source checkout is required. Vendor drivers, station settings,
and approved calibration remain necessary for hardware operation.

## Launch and verify

Copy the **entire** `dist/RapidPyMain` folder, including `_internal`, to the
operator computer. Launch `RapidPyMain.exe`. The console entry point supports:

```powershell
.\RapidPyMainConsole.exe --check-startup
.\RapidPyMainConsole.exe --smoke-test
.\RapidPyMain.exe --tool data_viewer
```

The smoke test uses temporary configuration and INI preferences, forces
No-Communication, constructs all six main panels, imports all six helper entry
points, prints JSON, and exits. It never uses operator configuration or connects
to instruments. Startup diagnostics are read-only. Configuration normally
lives in `~/.rapid/config.json`; `RAPID_CONFIG` selects a different file.
Windows preferences and queue recovery state belong to the operator account.
Main-window helper launches use separate bundled processes; VRM retains its
ownership leases and handoff environment.

## Operator workflow

1. Open Settings. Select No-Communication for practice and set the data folder.
   For supervised hardware testing, import the station INI and review ports,
   geometry, motion positions, enable flags, and calibration.
2. Load a sample index or enter specimens in Sample Queue. Review specimen
   names, changer positions, and orientations.
3. Build/load a sequence, inspect its preview, save it, and select a specimen
   or queue. Live treatment plans are validated before ordinary preflight.
4. Measure the holder before specimens. Rejected magnetic blocks or failed
   bridge reads cannot replace the accepted holder correction.
5. Watch Live Measure/Step Monitor. Use Pause/Halt and inspect the safe-return
   result before touching the instrument.
6. Review results in Plots/Data Review. Simulation output stays in `SIMULATED`.
   Run bundles retain scientific outputs, provenance, workflow summaries,
   communication evidence, and a hashed artifact index.
7. After restart, review interrupted rows and choose Resume, Re-run, Skip, or
   Abort. Restored queue state alone does not verify physical specimen location.

Calibration Center records proposals, approval, expiry/invalidation, and
rollback. Review/approve calibration before using its coefficients. Help →
“Where did this VB6 control go?” supplies the searchable transition map.

## Rebuild and test

Use an environment with the declared RapidPy dependencies, setuptools, wheel,
and PyInstaller. The default build interpreter is the repository `.venv`:

```powershell
.\installer\build_rapid_main.ps1
.\.venv\Scripts\python.exe tools\test_rapid_main.py
```

The build packages a staged installation wheel, including metadata, icons,
common assets, helpers, and GPL licensing. Tests isolate configuration and
Qt preferences instead of changing the operator's registry or saved queue.

## Remaining full-replacement gates

This is a **pilot candidate**. The source audit identified software integrations
that must not be mistaken for physical acceptance alone:

- Backfield IRM needs the legacy dedicated polarity relay mapping. Negative
  requests cannot run through the unsigned ramp or be silently clamped to zero.
- RRM requires synchronized turning-motor spin, selected-coil AF ramp, and
  verified stop (`VB6/RockmagStep.cls`). Planning labels do not implement it.
- Full AF treatment requires coil centering, axial and two transverse passes,
  orientation changes, and station field-to-voltage calibration. One ADwin
  ramp does not prove parity with `VB6/RockmagStep.cls`.
- ARM/pulse IRM needs independent bias/pulse circuitry integration. The present
  effective-peak calculation is not evidence of equivalent ARM physics or
  pulse-IRM behavior.
- MCC/DAC and transverse motor execution have planning APIs but need production
  adapter binding and loopback/position evidence.
- Thermal treatment remains manual/external; automated furnace operation needs
  an identified controller, protocol, readback, and interlocks.

These integrations remain open. Vendor driver installation, fresh-Windows
deployment, mixed-DPI display review, physical fault/recovery tests, scientific
reference comparisons, and lab approval also remain pending. The portable app
does not install kernel drivers. Follow
`docs/rapid_hardware_acceptance_procedure_2026-08-29.md` for bench qualification.

## Verification checkpoint — 3 October 2026

The continuation suite passes all 599 tests in 82.771 seconds. The copied
portable bundle passes dependency/icon discovery and the offline six-panel,
six-helper UI smoke check on this Windows 10 computer. Packaging excludes
incompatible ICU/CRT shims discovered on the build host's tooling PATH, so Qt
uses the Windows ICU API. Reproduce the checks with:

```powershell
.\installer\verify_rapid_main.ps1 -Bundle .\dist\RapidPyMain
```

This validates offline launch on this computer; it does not replace a fresh
Windows deployment or instrument/scientific acceptance.
