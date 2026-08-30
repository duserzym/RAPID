# RapidPy Hardware/Workflow Acceptance Checklists (2026-07-15)

This file is the acceptance gate set for objective-bound production remediation.
Each workflow below must have:
- code-anchored evidence for operator-accessible paths,
- a hardware evidence record from live bench validation, or
- explicit retirement rationale in `docs/vb6_parity_inventory.md`.

## Core parity workstreams (operator-facing)

- [ ] AF demagnetization execution in `rapid_main`
  - [x] Software evidence: AF labels now route through an explicit AF demagnetizer command planner/backend instead of the IRM treatment path, covered by `RapidPy/rapid_main/tests/test_diagnostic_services.py` and `RapidPy/rapid_main/tests/test_queue_and_bundle.py`.
  - [x] Software evidence: queue safe-state return calls the AF backend reset hook when the AF backend is connected, preserving a software contract for park-after-stop.
  - [ ] Log/telemetry includes target profile, frequency/settle/tumble metadata, and run completion/fault markers from live AF execution.
  - [ ] **Hardware-only evidence required:** live AF rig pass/fail with representative field profile.
- [ ] SQUID communication parity
  - [ ] Real transport adapter path tested with hardware and failure timeout/retry behavior covered.
  - [x] Software evidence: magnetometer quality flags are modeled in `RapidPy/rapid_main/rapid_main/magnetometer.py` and surfaced by `MeasurementWorker` when a backend returns `MagnetometerReading`.
  - [x] Software evidence: shared SQUID readiness snapshots classify connected vs disconnected transport state without taking a measurement sample, covered by `RapidPy/rapid_main/tests/test_diagnostic_services.py`.
  - [x] Software evidence: hardware-mode queue startup and per-command execution block/halt unsafe measurement transitions when SQUID communication is not healthy, covered by `RapidPy/rapid_main/tests/test_queue_orchestration.py`.
  - [ ] **Hardware-only evidence required:** no-comm + valid COM/serial path exercised on SQUID rig.
- [ ] Vacuum control parity
  - [x] Software evidence: `VacuumDialog` reports pressure, pump state, connection/status text, and shared fault classification from `read_vacuum_snapshot`, covered by `RapidPy/rapid_main/tests/test_diagnostic_services.py`.
  - [x] Software evidence: queue startup and per-command execution call the vacuum readiness guard before unsafe transitions, covered by `RapidPy/rapid_main/tests/test_queue_orchestration.py`.
  - [x] Software evidence: `vacuum_fault` equivalent state escalates high pressure/readback failures into queue halt/error flow with existing safe-state return.
  - [ ] **Hardware-only evidence required:** controlled pressure sequence with fault injection.
- [ ] IRM voltage calibration workflow
  - [x] Software evidence: `RapidPy/rapid_main/rapid_main/calibration.py` provides IRM field-to-voltage calibration fits and limit-checked voltage requests, covered by `RapidPy/rapid_main/tests/test_calibration.py`.
  - [x] Software evidence: `IrmArmConfig` and Settings retain IRM voltage slope/intercept/limit values, and `diagnostic_services.py` consumes those values for ADwin DAC planning, covered by `RapidPy/rapid_main/tests/test_diagnostic_services.py`.
  - [x] Software evidence: Calibration Center exposes an `IRM Voltage Calibration` operator path with measured field/voltage points, run context, operator identity, and max-voltage settings; the Diagnostics menu routes `IRM Voltage Calibration` to this panel.
  - [x] Software evidence: IRM voltage calibration artifacts are persisted as timestamped JSON with fit parameters, measured points, run context, and operator metadata, covered by `RapidPy/rapid_main/tests/test_calibration.py` and `RapidPy/rapid_main/tests/test_calibration_panel.py`.
  - [ ] Recovery path handles calibration interruption without leaving coil ownership locked.
  - [ ] **Hardware-only evidence required:** calibration execution on IRM/ARM rig plus output review.
- [ ] 908A Gaussmeter / SQUID baseline verification
  - [x] Software evidence: Calibration Center writes timestamped SQUID/Gaussmeter baseline JSON artifacts with expected value, samples, mean/stdev, residuals, pass/fail status, run context, and operator metadata, covered by `RapidPy/rapid_main/tests/test_calibration.py` and `RapidPy/rapid_main/tests/test_calibration_panel.py`.
  - [ ] **Hardware-only evidence required:** live 908A/SQUID baseline verification against known standard and accepted residual tolerance.
- [ ] DAC / MCC pathway parity
  - [x] Software evidence: `RapidPy/rapid_main/rapid_main/dac.py` defines validated channel commands, safe-zero commands, and blocked out-of-range writes, covered by `RapidPy/rapid_main/tests/test_dac.py`; IRM/ARM ADwin field-to-voltage planning now uses the same validation in `RapidPy/rapid_main/rapid_main/diagnostic_services.py`, covered by `RapidPy/rapid_main/tests/test_diagnostic_services.py`.
  - [ ] DAC/MCC interfaces are either production-backed by hardware adapter or explicitly retired.
  - [x] If retained: software startup readiness reports enumerate safe-zero defaults and surface missing/invalid adapter errors before any hardware write, covered by `RapidPy/rapid_main/tests/test_dac.py`.
  - [ ] **Hardware-only evidence required:** minimum one validated signal channel round-trip.
- [ ] Thermal workflow parity
  - [x] Software evidence: `RapidPy/rapid_main/rapid_main/thermal.py` defines validated thermal treatment steps, safety limits, ramp/hold/cool estimates, and queue-block conversion, covered by `RapidPy/rapid_main/tests/test_thermal.py`.
  - [x] Software evidence: `QueueHardwareBackend` recognizes `TT`/`TH`/`TEMP` labels and routes them to an optional live `apply_thermal(temperature_c, label)` backend hook, covered by `RapidPy/rapid_main/tests/test_queue_and_bundle.py`.
  - [x] Software evidence: Calibration Center exposes `Thermal Routine Planning` with operator/run context, safety limits, target list, queue-compatible labels, and timestamped JSON planning artifacts, covered by `RapidPy/rapid_main/tests/test_thermal.py` and `RapidPy/rapid_main/tests/test_calibration_panel.py`.
  - [ ] Thermal controls are live-integrated or explicit retirement decision recorded with rationale.
  - [x] If active: software planning has preheat/ramp/hold/cool thresholds and rejects unsafe targets before queue conversion.
  - [ ] **Hardware-only evidence required:** thermal routine run with safe abort and alarm handling.
- [ ] Rockmag routines
  - [x] `RockmagStep`/`RockmagSteps` model is represented in queue/compiler path.
  - [x] Software evidence: `RapidPy/rapid_main/rapid_main/rockmag.py` compiles operator routine specs/templates into family-preserving blocks and runner labels, covered by `RapidPy/rapid_main/tests/test_rockmag_routine.py`.
  - [x] Software evidence: the Sequence panel "Rockmag the Works" preset now uses the rockmag compiler and returns to manual control generation after edits, covered by `RapidPy/rapid_main/tests/test_sequence_rockmag.py`.
  - [x] Software evidence: Rockmag routine plans can be written as deterministic JSON sidecar artifacts with queue labels, block metadata, operator/run context, and explicit `hardware_validation_required`, covered by `RapidPy/rapid_main/tests/test_rockmag_routine.py`.
  - [ ] End-to-end hardware launch and physical output generation for at least one rockmag template.
  - [ ] **Hardware-only evidence required:** run captured in live sample flow.
- [ ] VRM acquisition parity
  - [x] Software evidence: `rapid_main` launches VRM acquisition with a persisted `rapidpy.vrm.launch_context.v1` handoff and `vrm_logger` writes a `.vrm.json` sidecar beside the selected CSV, covered by `RapidPy/rapid_main/tests/test_app_bootstrap_window_contracts.py` and `RapidPy/rapid_main/tests/test_vrm_integration.py`.
  - [x] Software evidence: VRM output sidecars preserve CSV path, session start, interval/spacing, display unit, baseline, calibration, and rapid_main handoff metadata for archive/run review.
  - [ ] **Hardware-only evidence required:** logged acquisition session with run association.
- [ ] Measurement cycle statistics (`frmStats`)
  - [x] Software evidence: configured repeated SQUID reads are summarized by `reading_cycle_statistics`, saved as the cycle mean, and emitted with `StepResult`; covered by `RapidPy/rapid_main/tests/test_analysis.py` and `RapidPy/rapid_main/tests/test_measurement_worker.py`.
  - [x] Software evidence: the Measurement panel renders live cycle mean, X/Y/Z ranges, ratios, directional spread, and signal-to-drift values; holder and induced values remain explicitly `N/A` without supplied baselines, covered by `RapidPy/rapid_main/tests/test_measurement_panel.py`.
  - [ ] **Hardware-only evidence required:** a repeat-reading run with real SQUID, holder, and induced-field baselines to validate physical quality thresholds.
- [ ] DC motor live telemetry
  - [x] Software evidence: integrated and standalone tools render input command, output/encoder feedback, velocities, and optional torque traces with aliases and `N/A` fallback; covered by `RapidPy/rapid_main/tests/test_dcmotors_telemetry.py`, `RapidPy/dc_motor_control/tests/test_telemetry_ui.py`, and `RapidPy/dc_motor_control/tests/test_motor_telemetry.py`.
  - [ ] **Hardware-only evidence required:** a live axis command with encoder feedback and torque-readback comparison against the controller's native units.
- [x] Startup guidance (`frmSplash`, `frmTip`)
  - [x] Software evidence: `StartupGuideDialog` is non-modal, appears after the main window has fitted itself, can be reopened from Help > Quick Start, persists its preference, and routes to Settings or Sample Queue; covered by `RapidPy/rapid_main/tests/test_startup_guide.py`.

## Scientific analysis and diagnostics

- [ ] Scientific analysis modules (`modDataAnalysis`, `modListenAndLog`, plotting math parity)
  - [x] Regression tests for formula/mapping path exist for analysis outputs used in production paths (`RapidPy/rapid_main/tests/test_analysis.py`).
  - [x] Software evidence: measurement runs now emit `communication.tsv` transcript artifacts through `MeasurementWorker`, covered by `RapidPy/rapid_main/tests/test_measurement_worker.py`.
  - [x] Software evidence: transport-neutral communication transcript and serial-like bridge behavior is covered by `RapidPy/rapid_main/tests/test_communication_log.py`.
  - [ ] Output plots/transforms match legacy schema where still required.
  - [x] Operator-visible diagnostics include stable error classification primitives (`info`, `warning`, `error`, `safety`) through `RapidPy/rapid_main/rapid_main/status_codes.py` and worker `status_event` emissions.
  - [x] Software evidence: Debug Console can refresh a structured backend readiness snapshot for Vacuum, AF Demag, IRM/ARM, SQUID, and DC Motors, using `collect_diagnostic_status` to surface connected/simulated/status/fault state; covered by `RapidPy/rapid_main/tests/test_diagnostic_services.py`.
  - [ ] **Hardware-only evidence required:** analysis on at least one real measurement bundle.
- [ ] Plotting/quicklook workflow (`frmPlots*`, rapid review panel)
  - [x] Software evidence: on-demand quicklook path is available from the Measurement panel and feeds completed run steps into `PlotsDialog`, covered by `RapidPy/rapid_main/tests/test_measurement_panel.py`.
  - [x] Software evidence: `PlotsDialog` records a backend-independent quicklook data contract for labels, vectors, intensity, declination, and inclination, covered by `RapidPy/rapid_main/tests/test_plots_dialog.py`.
  - [x] Software evidence: quicklook outputs can be written as a JSON summary sidecar for reproducible metadata review, covered by `RapidPy/rapid_main/tests/test_plots_dialog.py`.

- [ ] Susceptibility workflow (`frmSusceptibilityMeter`, `modSusceptibility`)
  - [x] Software evidence: measurement runs persist susceptibility readings in `.rmg` and write `susceptibility.json` with per-step values, summary statistics, and hardware-validation caveat, covered by `RapidPy/rapid_main/tests/test_measurement_worker.py`.
  - [ ] **Hardware-only evidence required:** calibrated susceptibility bridge reading captured from live hardware and reconciled against expected standard.
  - [x] Software evidence: successful, non-aborted measurement runs with completed steps automatically write `quicklook.json` into the same output bundle path passed to `MeasurementWorker`, covered by `RapidPy/rapid_main/tests/test_measurement_panel.py`.
  - [ ] **Hardware-only evidence required:** real bundle with plotting review pass.

## Topology and geometry proof requirements (all first-party apps)

- [ ] All 16 app entrypoints launch within current monitor working area at startup on compact and mixed monitors.
   - [x] Software evidence: static bootstrap contract test confirms shared `apply_window_bounds_guard(app)` call and absence of startup maximize/fullscreen paths in `RapidPy/rapid_main/tests/test_app_bootstrap_window_contracts.py`.
   - [x] Software evidence: offscreen geometry tests cover compact, small, and scaled-DPI representative profiles in `RapidPy/rapid_main/tests/test_window_layout.py` and startup panel safety checks in `RapidPy/rapid_main/tests/test_sequence_ui_smoke.py`, including persisted-geometry normalization, synthetic compact-screen resize clamping, tiny-profile clamping below the legacy fixed 360 px floor, oversized-minimum normalization, and direct `MainWindow._fit_window_to_current_screen` coverage for hostile full-width restores plus oversized sidebar state.
  - [x] Software evidence: DC-motor telemetry UX contract is regression-tested in `RapidPy/rapid_main/tests/test_dcmotors_telemetry.py`, including visible initial value-tile layout plus compact/narrow reflow for input command, output feedback, residual, velocity, and torque readouts.
  - [ ] Hardware-only evidence: manual startup/relocation verification on:
    - at least one `1366x768` or lower display,
    - one mixed-resolution multi-monitor set,
    - one DPI-scaled profile (125% / 150% Windows scaling).

- [ ] Screen-topology change handling keeps every app safe.
  - [x] Software evidence: the shared guard re-fits top-level windows after screen add/remove, working-area/geometry changes, and logical-DPI changes; `RapidPy/rapid_main/tests/test_window_layout.py` covers synthetic topology re-clamping and `test_app_bootstrap_window_contracts.py` guards the event subscriptions.
  - [ ] Hardware-only evidence: moving/re-connecting displays while each app is open does not force zero-size or unusable minimums; operator can continue interaction.

## Interruption and recovery semantics

- [ ] Interrupted-run recovery
  - [ ] `rapid_main` persists queue state in a safe checkpoint that includes sample index + phase marker.
  - [x] Software evidence: recovery UI exposes `Resume`, `Re-run`, `Skip`, and `Abort` operators for interrupted rows through the Sample Queue recovery menu and interrupted-row context menu.
  - [ ] Restart path verifies mechanical ownership is re-acquired (or intentionally released) before continuation.
  - [x] Software evidence: automated restoration tests in `RapidPy/rapid_main/tests/test_queue_orchestration.py` verify interrupted `Running` rows are restored as `Interrupted` with resume metadata; `RapidPy/rapid_main/tests/test_sample_queue_helpers.py` verifies unresolved interrupted rows cannot be queued until an explicit recovery action is chosen.
  - [x] Software evidence: completed measurement workers write `workflow_summary.json` with emitted workflow phases, final aborted/completed state, labels, status codes, and hardware-validation caveat, covered by `RapidPy/rapid_main/tests/test_measurement_worker.py`.
  - [x] Software evidence: measurement workers write `artifact_index.json` for successful, halted, unavailable-backend, and preflight-failed exits so operators can distinguish created, missing, optional UI-generated, and hardware-validation-required run artifacts.
  - [ ] **Hardware-only evidence required:** interrupt + resume exercise in live queue run.

## Layout/responsiveness gates for all first-party windows

- [ ] Window startup sizing
  - [x] Default startup geometry uses compact, off-screen-safe clamp.
  - [x] Restored geometries are normalized to current screen bounds (`availableGeometry`).
  - [x] Fullscreen/maximized restore artifacts cannot block startup at full width.
- [ ] Screen topology changes
  - [x] Shared guard subscriptions re-clamp windows after screen add/remove, work-area, geometry, and logical-DPI updates; synthetic negative-origin, compact, tiny, high-resolution, and scaled-DPI profiles are covered.
  - [ ] **Hardware-only evidence required:** move/attach and mixed-DPI workflows on native Windows displays keep each app usable.
- [ ] Rapid-main UI geometry
  - [x] Left icon rail remains readable within a `48-72 px` compact envelope and never exceeds its restore cap, including direct regression coverage for an oversized restored splitter/sidebar state.
  - [x] Stacked panels preserve only a screen-capped vertical minimum and do not inflate startup width.
  - [x] DC motor diagnostics telemetry panel shows input/output and torque visibility at compact startup geometries, with N/A fallback when torque is not reported.

## 2026-07-16 Software Acceptance Record

The following focused suites were run with `E:\Github\RAPID\.venv\Scripts\python.exe`; Qt suites used `QT_QPA_PLATFORM=offscreen` and the relevant `PYTHONPATH` package roots.

| Scope | Executed suites | Result |
|---|---|---|
| Measurement statistics | `test_analysis.py`, `test_measurement_worker.py`, `test_measurement_panel.py` | 28 passing tests |
| Startup guidance | `test_startup_guide.py` | 2 passing tests |
| Window/layout | `test_window_layout.py`, `test_app_bootstrap_window_contracts.py`, `test_sequence_ui_smoke.py` | 51 passing tests |
| DC motor telemetry | `test_motor_telemetry.py`, `test_telemetry_ui.py`, `test_dcmotors_telemetry.py`, `test_diagnostic_services.py` | 40 passing tests |
| Focused total | all suites above | 121 passing tests |

This is software acceptance evidence only. It does not replace bench validation for live motor torque/encoder feedback, physical SQUID baselines, or native Windows monitor/DPI/hotplug behavior.

## Evidence artifact index (expected updates)

- Code evidence:
  - `RapidPy/rapidpy_common/ui.py`
  - `RapidPy/rapid_main/rapid_main/app.py`
  - `RapidPy/adwin_comms/adwin_comms/app.py`
  - `RapidPy/vrm_logger/vrm_logger/app.py`
- Test evidence:
  - `RapidPy/rapid_main/tests/test_window_layout.py`
  - `RapidPy/rapid_main/tests/test_app_bootstrap_window_contracts.py`
  - `RapidPy/rapid_main/tests/test_sequence_ui_smoke.py`
  - `RapidPy/rapid_main/tests/test_dcmotors_telemetry.py`
  - `RapidPy/rapid_main/tests/test_status_codes.py`
  - `RapidPy/rapid_main/tests/test_analysis.py`
  - `RapidPy/rapid_main/tests/test_calibration.py`
  - `RapidPy/rapid_main/tests/test_communication_log.py`
  - `RapidPy/rapid_main/tests/test_dac.py`
  - `RapidPy/rapid_main/tests/test_magnetometer.py`
  - `RapidPy/rapid_main/tests/test_measurement_worker.py`
  - `RapidPy/rapid_main/tests/test_measurement_panel.py`
  - `RapidPy/rapid_main/tests/test_startup_guide.py`
  - `RapidPy/dc_motor_control/tests/test_motor_telemetry.py`
  - `RapidPy/dc_motor_control/tests/test_telemetry_ui.py`
  - `RapidPy/rapid_main/tests/test_rockmag_routine.py`
  - `RapidPy/rapid_main/tests/test_thermal.py`
  - `RapidPy/rapid_main/tests/test_calibration_panel.py`
  - `RapidPy/rapid_main/tests/test_transverse.py`
- Gap/tracker evidence:
  - `docs/vb6_parity_inventory.md`
