# RapidPy production-readiness assessment

> **October 2, 2026 update — read this first.** The July 16 material below is
> retained for history. Where the two disagree, this section wins.

## Code-complete versus hardware-validated (August 29, 2026)

RapidPy is **transition/testing software**, not a VB6 replacement. The P0
measurement-integrity path is code-complete and covered by replay and fault
tests; none of it has been validated against the physical RAPID system.

| P0 area | Code-complete | Hardware-validated | Evidence |
|---|---|---|---|
| Bracketed SQUID acquisition (`Measure_ReadSample` order) | Yes | **No** | `rapid_main/acquisition.py`, `tests/test_acquisition.py` |
| Coherent latch/counter/DVM observation evidence | Yes | **No** | `rapid_main/magnetometer.py`, `rapid_main/squid_transport.py`, `tests/test_squid_transport.py` |
| Flux-count rejection on X, Y, Z | Yes | **No** | `tests/test_acquisition.py`, `tests/test_vb6_parity.py` |
| Safe recovery and retry exhaustion | Yes | **No** | `tests/test_acquisition.py`, `tests/test_output_integrity.py` |
| Measured, persisted holder correction | Yes | **No** | `rapid_main/holder_state.py`, `rapid_main/holder_measurement.py`, `tests/test_holder_state.py` |
| Holder-gated sample measurement | Yes | **No** | `tests/test_queue_hardware_backend.py` |
| Fail-closed hardware mode | Yes | **No** | `rapid_main/diagnostic_services.py`, `tests/test_diagnostic_services.py`, `tests/test_queue_hardware_backend.py` |
| Simulation isolation and labelling | Yes | n/a | `rapid_main/io/measurement_bundle.py`, `tests/test_output_integrity.py` |
| Transactional output, duplicate-free resume | Yes | **No** | `tests/test_output_integrity.py`, `tests/test_vb6_parity.py` |
| Specimen metadata resolution | Yes | **No** | `rapid_main/specimen_metadata.py`, `tests/test_output_integrity.py` |
| Replay of recorded blocks (acceptance step D1) | Yes | n/a | `rapid_main/replay.py`, `tests/fixtures/`, `tests/test_replay_fixtures.py` |
| Motion and interlock behavior | Verified in software only | **No** | Every motion is verified and a failure aborts the block |
| Output parity against VB6 | Fixtures only | **No** | VB6 now compiles; side-by-side comparison still requires the no-communication smoke test and a physical reference run |

Full suite: **440 tests passing** (`python -m unittest discover -s tests -p
'test_*.py'` from `RapidPy/rapid_main`), up from a 261-test baseline. The
October 2 shell slice adds truthful Dashboard backend snapshots, responsive
glass-card reflow, workflow/session menu wiring, shutdown ordering, atomic
sequence documents, truthful empty operator states, real sample-index loading,
AF simulation isolation, atomic settings backup/restore, and searchable VB6
transition-help tests. It also covers immutable calibration artifact snapshots,
versioned approvals, expiry/invalidation, event-based rollback, integrity
checks, and measurement-provenance linkage.

### Behavior changes an operator will notice

- Hardware mode blocks instead of falling back to a simulator. Missing
  packages, ports, adapters, or calibration are named as preflight blockers.
- `motion.zero_pos` / `motion.meas_pos` default to unconfigured, so hardware
  preflight blocks until the legacy `[SteppingMotor]` values are imported or
  real measured values are entered. RapidPy will not invent lift positions.
- The `Holder` queue command now measures the holder. A rejected holder block
  aborts the queue and keeps the previous correction.
- Help now includes a searchable VB6-to-RapidPy task map with readiness labels
  and direct routing to the corresponding real panel or diagnostic launcher.
- Calibration Center now distinguishes a recorded artifact from an approved,
  active calibration. Approvals, activations/rollbacks, and invalidations are
  append-only; expired, invalidated, missing, or hash-mismatched artifacts are
  excluded from new measurement provenance.
- Sample measurement is blocked when no valid holder correction exists.
- Measurement output is published only when a run completes; an aborted run
  leaves the production output path untouched.
- Simulated runs publish into a `SIMULATED` subdirectory with a marker file.
- Sig/Holder and Sig/Induced now show real values when a bracketed block is
  available, plus holder identity, magnitude, and age.
- The main Dashboard reads the same backend snapshot as the debug console and
  visibly distinguishes live, disconnected, unavailable, faulted, and
  **SIMULATED** devices. It refreshes on demand, every ten seconds, and after a
  communication-mode change.
- Imported sequences remain the active executable/saveable document; writes
  are atomic, malformed files report actionable errors, unsaved edits prompt
  on replacement/exit, and the Hawaiian preset emits its documented AF25–AF800
  steps.
- Main-shell Flow and session actions invoke real pause/resume/halt/session
  behavior. The legacy Code Grey equivalent is explicitly visual-only and
  cannot bypass preflight or interlocks.

### Still open in software

- 2G range-letter mapping unconfirmed against the instrument.
- Holder direction policy defaults to `shared` (VB6 parity); `strict` is
  available but is a deviation.
- RapidPy issues one blocking verified motion where VB6 issues a non-blocking
  start followed by a blocking repeat.

The physical procedure and evidence schema are in
`docs/rapid_hardware_acceptance_procedure_2026-08-29.md`. VB6 build/launch gate
status is in `docs/rapidpy_transition_readiness_2026-08-29.md`.

---

# RapidPy production-readiness assessment (July 16, 2026)

## Scope
This file tracks parity and hardening evidence for the active `rapid_main` workflow and the 16 standalone RapidPy apps.

## High-level status (July 16, 2026 update)

- **VB6 parity goal:** not yet complete.
- **Current readiness posture:** core workflows and shared UI/window foundations are implemented, but multiple domains still need hardware-backed acceptance before production closure.
- **Window/layout readiness:** compact sizing, restore guards, and runtime topology/DPI re-fit handling are covered in software; full operator validation on varied Windows DPI + monitor topologies remains outstanding.

## Direct feasibility assessment against current user questions (July 15/16, 2026)

- **VB6 functionality parity for RapidPy package:** **Not fully complete.** `vb6_parity_inventory.md` still tracks unresolved items across AF transport completion, thermal, rockmag, DAC/MCC, SQUID transport hardening, vacuum production transport edge-cases, and scientific analysis paths. Automatic email dispatch (`frmSendMail`) is explicitly retired in favor of operator-controlled sharing of deterministic measurement bundle outputs. Current status is best described as **mapped + partial + simulated-safe fallback** in many domains, not end-to-end production parity.
- **16 app windows displaying/initializing safely on all monitors:** partially complete. The shared boundary guard and startup contract coverage is present for 16 `app.py`/`main.py` entrypoints, and software now re-fits top-level windows after screen add/remove, work-area, geometry, and logical-DPI changes. Runtime validation against real multi-monitor + DPI-change permutations is still a physical evidence gap.
- **Auto-resize/layout across Windows machines:** partially complete. The shared geometry clamp is global (`rapidpy_common/ui.py`), `rapid_main` adds compact restore logic, and offscreen profile tests cover negative origins, small/tiny displays, scaled DPI, and topology re-clamping. Native field validation still requires actual display hardware.
- **Main app fully functional on every aspect:** **No.** Workflow scaffolding exists, but multiple critical paths are still operator- or hardware-acceptance-gated:
  - no-comm / simulator behavior is intentionally preserved in several services,
  - queue/changer and diagnostic hardware transitions need live transport acceptance evidence,
  - thermal/rockmag/VRM/scientific plotting domains now have mapped or integrated software foundations, but still require live hardware/run-bundle acceptance where retained.
- **DC motor telemetry requirement status:** met in software for both integrated dialog and standalone controller (`rapid_main/dialogs/dc_motors.py`, `dc_motor_control/app.py`, `rapid_main/tests/test_dcmotors_telemetry.py`, `dc_motor_control/tests/test_telemetry_ui.py`, `dc_motor_control/tests/test_motor_telemetry.py`): live input/output/velocity telemetry and optional torque fallback/alias extraction are wired with plotted traces. A real-axis encoder/torque comparison remains a hardware-only gate.

## Verified entrypoint hardening status (static evidence)

- `16` application entrypoints under `RapidPy/*/*/app.py` were discovered and all are expected to go through the shared startup guard contract checked in `rapid_main/tests/test_app_bootstrap_window_contracts.py`.
- Startup wrappers (`main.py`) exist for `16` projects and are expected to dispatch through `app.main()`, avoiding direct `QApplication` bootstrapping.
- No startup path in `main.py` or `app.py` is currently allowed to call `showMaximized()` or `showFullScreen()` by contract.

## Layout and multi-monitor readiness status (evidence-backed)

- `clamp_window_geometry` enforces a compact profile (max width fraction + min/max floors) in
  [`rapidpy_common/ui.py`](rapidpy_common/ui.py), with an added safeguard for very small/offset viewports.
- `test_window_layout.py` now includes tiny-viewport cap regressions so geometry remains bounded by
  the available screen on very small monitors and unusual taskbar placements.
- All app entrypoints register the window-bounds guard in `main()`:
  [`rapidpy/common/ui.py`] is used by every discovered entrypoint and validated by
  [`rapid_main/tests/test_app_bootstrap_window_contracts.py`](rapid_main/tests/test_app_bootstrap_window_contracts.py).
- `rapid_main` adds explicit compact-fit and restore-clamp flow:
  [`MainWindow._fit_window_to_current_screen`](rapid_main/app.py),
  [`_restore_layout_state`](rapid_main/app.py),
  [`_clamp_main_window_size`](rapid_main/app.py),
  and sidebar cap behavior in
  [`_build_central`](rapid_main/app.py).
- The current glass workspace keeps a readable text sidebar with
  `_DEFAULT_SIDEBAR_WIDTH = 252`, `_MIN_SIDEBAR_WIDTH = 240`,
  `_MAX_SIDEBAR_WIDTH = 288`, and `_SIDEBAR_RESTORE_RATIO = 0.25`. The main
  workspace uses up to 94% of the monitor with explicit 1680×1100 caps rather
  than the superseded 12%-width compact profile.
- `rapidpy_common.ui._screen_area_for_widget` now prefers the nearest-screen candidate when a widget is
  positioned off-screen before falling back to the primary screen, reducing stale-geometry monitor
  selection risk after multi-monitor changes.
- `rapid_main` now uses a translucent glass backdrop, elevated cards, explicit
  focus/disabled states, readable icon-plus-text navigation, and responsive
  Dashboard cards. The compact profile reflows readiness cards to two columns
  and stacks Current Run/Quick Actions; the wide profile restores five columns
  and a split run/action row.
- New regression coverage added:
  - `rapid_main/tests/test_window_layout.py` (compact geometry, nearest-screen fallback, negative-origin/topology re-clamping, scaled-DPI, and screen-capped stack-hint checks).
  - `rapid_main/tests/test_sequence_ui_smoke.py` (icon rail readability, wrapped text surfaces, stacked-panel hint safety, and restore clamping tests).
  - `rapid_main/tests/test_app_bootstrap_window_contracts.py` (entrypoint bootstrap contracts include `app.py` and `main.py` launchers plus topology/DPI guard subscriptions).
  - `rapid_main/tests/test_startup_guide.py` (first-run startup guidance preference and Help-panel routing).
  - `rapid_main/tests/test_dashboard_shell.py` (truthful live/simulated/fault states, responsive reflow, main run-state synchronization, and halt-after-confirm shutdown ordering).
  - `rapid_main/tests/test_sequence_documents.py` (atomic sequence persistence, strict import errors, compiled preset behavior, and unsaved-state tracking).

## July 16 Focused Acceptance Evidence

Focused software acceptance completed with 121 passing tests across measurement statistics, startup guidance, layout/topology, and DC motor telemetry. The complete command-to-result table is maintained in `docs/hardware_acceptance_checklists_2026-07-15.md`.

- Measurement statistics: repeated SQUID-cycle mean and quality evidence are rendered live by `MeasurementPanel`. *(Superseded 2026-08-29: holder and induced values are rendered from a bracketed block, and holder identity/magnitude/age/validity are shown.)*
- Startup guidance: `StartupGuideDialog` replaces the unimplemented splash/tip tracker entries with a non-modal first-run guide, persisted preference, and Help > Quick Start route.
- Physical-only remaining evidence: live SQUID holder/induced baselines, motor encoder/torque comparison, and native Windows monitor add/remove plus mixed-DPI behavior.

## VB6 parity ledger (current)

| Domain | Status | Evidence | Notes |
|---|---|---|---|
| AF workflows | **Integrated + explicit AF treatment contract** | `MainWindow._prepare_af_workflow`, `Queue options`, AF labels, `diagnostic_services.AfDemagBackend`, `plan_af_demag_command`, `QueueHardwareBackend.set_demag_step`, `tests/test_app_af_workflow.py`, `tests/test_diagnostic_services.py`, `tests/test_queue_and_bundle.py` | AF labels route through an AF demagnetizer planner/backend instead of the IRM path. Synthetic AF examples are unmistakably named, disabled and refused in hardware mode, and never auto-start live hardware. Live AF rig pass/fail, transport telemetry, and hardware fault evidence remain acceptance gates. |
| SQUID communication/plots | **Partial (mixed HW + fallback + software readiness guard)** | `diagnostic_services.SquidBackendAdapter`, `diagnostic_services.SquidNoCommBackend`, `read_squid_snapshot`, `require_squid_ready`, `dialogs/squid_comm.py`, `measurement_worker.py`, `tests/test_diagnostic_services.py`, `tests/test_queue_orchestration.py` | SQUID live and treatment control can run with live transport when available; simulation fallback remains for no-comm/unavailable hardware. Hardware-mode queues now block or halt measurement transitions when SQUID communication is disconnected. Live serial timeout/retry validation remains open. |
| Magnetometer normalization (`modMagnetometer`) | **Integrated foundation** | `magnetometer.py`, `measurement_worker.py`, `tests/test_magnetometer.py`, `tests/test_measurement_worker.py` | Raw SQUID axis voltages can be background-corrected and converted to calibrated moment vectors with quality flags, and the measurement loop now accepts `MagnetometerReading` payloads while surfacing read-quality flags as operator warnings. Live SQUID adapter rollout plus timeout/retry evidence remains open. |
| Vacuum | **Partial (sim + transport fallback + software fault guard)** | `diagnostic_services.VacuumBackendAdapter`, `VacuumNoCommBackend`, `dialogs/vacuum.py`, `tests/test_diagnostic_services.py`, `tests/test_queue_orchestration.py` | Real transport is used when `updown_control` dependency and port are available; otherwise pressure simulation is retained. Shared vacuum snapshots now report pressure, pump state, status, and high-pressure/readback faults, and queue automation blocks or halts through the existing safe-state path on vacuum faults. Live pressure fault-injection acceptance is still required. |
| IRM/ARM calibration | **Partial** | `diagnostic_services.IrmArmBackendAdapter`, `IrmArmNoCommBackend`, `dialogs/irm_arm.py` | ADWIN-backed calibration path exists; fallback simulator remains if adapter unavailable. |
| DAC/MCC / AF ramping | **Partial** | `rapidpy_common/adwin_af.py`, `diagnostic_services.IrmArmBackendAdapter` | ADWIN ramp API is used when present. Full DAC coverage confidence still depends on lab hardware integration verification. |
| DC motor control | **Integrated (live telemetry complete)** | `dialogs/dc_motors.py`, `diagnostic_services.DCMotorBackendAdapter`, `dc_motor_control/app.py` | Standalone dialog and standalone app include live command, output, velocity, and torque traces; torque extraction is alias-safe (`actual_torque`/`feedback_torque`/`torque`) and renders `% FS` on known scale. |
| VRM logging | **Integrated via external launch + run manifest sidecar** | `MainWindow._launch_vrm`, `rapid_main/vrm.py`, `vrm_logger/main.py`, `vrm_logger/session_manifest.py`, `tests/test_app_bootstrap_window_contracts.py`, `tests/test_vrm_integration.py` | Functionality remains in the dedicated app module; rapid_main now writes a launch handoff manifest and passes it to `vrm_logger`, which writes a `.vrm.json` sidecar beside the selected CSV with session/output metadata. Live logged acquisition with physical run association remains an acceptance gate. |
| Interrupted queue recovery | **Integrated + explicit operator choices + workflow summary artifact** | `rapid_main/panels/sample_queue.py`, `rapid_main/dialogs/step_monitor.py`, `measurement_worker`, queue state restore, `tests/test_queue_orchestration.py`, `tests/test_sample_queue_helpers.py`, `tests/test_measurement_worker.py` | Queue interruption persistence exists; restored `Running` rows become `Interrupted` and require an operator choice of Resume, Re-run, Skip, or Abort before queue start. Measurement workers now write `workflow_summary.json` phase evidence. Field replay and stressed restart behavior still need hardware acceptance evidence. |
| Thermal / scientific plotting | **Partial (integrated foundation)** | `measurement.py`, `runtime_estimator.py`, `thermal.py`, `hardware_contracts.py`, `data_model.MeasurementBlock`, `data_model.MeasurementBlocks`, `tests/test_thermal.py`, `tests/test_queue_and_bundle.py` | Thermal labels and planning now have validated temperature targets, ramp/hold/cool estimates, safety limits, queue-compatible labels, measurement-block metadata, and an optional queue-backend `apply_thermal` hook. Live furnace/oven control and full scientific plotting parity still need expanded hardware-in-the-loop coverage. |
| Plotting (measurement/stats) | **Integrated foundation** | `dialogs/plots.py`, `panels/measurement.py`, `data_viewer` app, `tests/test_plots_dialog.py`, `tests/test_measurement_panel.py` | Measurement panel emits live treatment-vs-step and Zijderveld traces, stores completed run steps, passes them to the plots dialog for Zijderveld/equal-area/intensity quicklook review, and writes `quicklook.json` into successful non-aborted measurement output bundles. The dialog starts empty, validates vector shape/finite values, and labels its optional synthetic example as non-hardware evidence. Advanced scientific plot parity and real-bundle review remain acceptance gates. |
| Email/report dispatch | **Retired with replacement** | `io/measurement_bundle.py`, `measurement_worker.py`, `docs/vb6_parity_inventory.md` | VB6 automatic mail dispatch is not retained; operator reporting uses `.sample`, `.rmg`, `measurements.txt`, and `specimens.txt` bundle outputs for review/sharing. |
| Status/error code taxonomy | **Mapped** | `status_codes.py`, `measurement_worker.py`, `tests/test_status_codes.py` | `MeasurementWorker` now emits structured `OperatorStatus` objects for workflow phases, preflight warnings, SQUID errors, file/output errors, and safety-return states while preserving current UI strings. |
| Debug diagnostics (`frmDebug`) | **Integrated foundation** | `dialogs/debug_console.py`, `diagnostic_services.collect_diagnostic_status`, `tests/test_diagnostic_services.py` | Debug Console can refresh a structured readiness snapshot for Vacuum, AF Demag, IRM/ARM, SQUID, and DC Motors, including connected/simulated/status/fault state. Live DAQ/ADwin adapter troubleshooting and raw-transport evidence remain acceptance gates. |
| Data analysis primitives (`modDataAnalysis`) | **Mapped** | `analysis.py`, `tests/test_analysis.py` | RapidPy now extracts measurement vectors from `MeasurementStep.sdx/sdy/sdz`, reports centroid/PCA-style principal-axis line-fit evidence, orientation/variance/residual metrics, and moment statistics/decay summaries. Advanced plotting integration and hardware-run bundle review remain acceptance caveats. |
| Communication listener/logger (`modListenAndLog`) | **Integrated foundation** | `communication_log.py`, `measurement_worker.py`, `tests/test_communication_log.py`, `tests/test_measurement_worker.py` | RapidPy now has deterministic communication transcript events, payload normalization, file export, a serial-like logging bridge, and per-run `communication.tsv` artifacts from the measurement loop. Full production coverage still requires adopting raw-transport bridges inside each live hardware adapter and collecting hardware transcripts. |
| Susceptibility (`frmSusceptibilityMeter`, `modSusceptibility`) | **Integrated foundation + run summary artifact** | `susceptibility.py`, `measurement_worker.py`, `tests/test_measurement_worker.py` | Measurement runs preserve susceptibility readings in legacy `.rmg` output and write `susceptibility.json` with per-step values, summary statistics, and explicit hardware-validation caveat. Live bridge calibration and standard-sample acceptance remain open. |
| 908A Gaussmeter / SQUID baseline (`frm908AGaussmeter`) | **Integrated foundation + operator artifact path** | `gaussmeter_control`, `calibration.py`, `panels/calibration.py`, `tests/test_calibration.py`, `tests/test_calibration_panel.py` | Standalone gaussmeter app remains available; rapid_main Calibration Center can record automated or manual SQUID/Gaussmeter baseline JSON artifacts with samples, residuals, pass/fail status, run context, and operator metadata. Live meter/standard acceptance remains open. |
| IRM voltage calibration (`frmIRM_VoltageCalibration`) | **Integrated foundation + operator artifact path** | `calibration.py`, `config.py`, `settings_panel.py`, `panels/calibration.py`, `diagnostic_services.py`, `tests/test_calibration.py`, `tests/test_calibration_panel.py`, `tests/test_diagnostic_services.py` | RapidPy now has auditable IRM field-to-voltage calibration points, linear fit output, retained slope/intercept/limit settings, DAC planning that uses those coefficients while blocking unsafe voltages, and a Calibration Center operator path that records timestamped JSON artifacts with run context/operator metadata. Live DAC binding, interruption recovery, and hardware execution acceptance remain gates. |
| DAC communication (`frmDAC_Comm`) | **Integrated foundation + startup readiness contract** | `dac.py`, `diagnostic_services.py`, `tests/test_dac.py`, `tests/test_diagnostic_services.py` | RapidPy now has a stable DAC command surface for channel limits, safe-zero requests, blocked out-of-range voltage commands without silent clamping, startup readiness reports for safe-zero defaults/missing adapter errors, and IRM/ARM ADwin field-to-voltage planning routes through that validation. Live MCC/DAQ binding and loopback acceptance remain open. |
| Interpolation/fitting utilities | **Mapped** | `geometry.InterpolationRange`, `geometry.InterpolationRanges`, `tests/test_geometry.py` | Linear interpolation and ordered range lookup are ported with validation, clamping, and descending-axis support; transverse workflow integration remains separately open. |
| Transverse auto-position (`frmTransverseProbeAutoPosition`) | **Mapped foundation** | `data_model.AngleVsFieldPoint`, `data_model.AngleVsFieldCollection`, `data_model.ProbeAngleOptimizer`, `transverse.py`, `tests/test_queue_and_bundle.py`, `tests/test_transverse.py` | Point and collection models preserve ordered observations; optimizer returns auditable peak-field angle results; auto-position planning now computes shortest signed move, hold/tolerance state, and confidence blocking. Motorized execution and physical probe acceptance remain open. |
| Rockmag routine (`frmRockmagRoutine`) | **Integrated foundation + planning artifact path** | `data_model.RockmagStep`, `data_model.RockmagSteps`, `rockmag.py`, `panels/sequence.py`, `tests/test_queue_and_bundle.py`, `tests/test_rockmag_routine.py`, `tests/test_sequence_rockmag.py` | Rockmag labels are parsed into family/value/unit metadata; operator-level routine specs/templates compile to family-preserving measurement blocks plus runner-compatible labels; deterministic JSON planning artifacts preserve queue labels, block metadata, and operator/run context; the Sequence panel "Rockmag the Works" preset now uses that compiler. Hardware execution and reproducible run-bundle acceptance remain open. |
| AF/IRMS/SQC queue recovery | **Integrated, needs acceptance** | queue persistence in `rapid_main/tests/test_queue_and_bundle.py`, explicit interrupted-row recovery in `rapid_main/tests/test_queue_orchestration.py`, `rapid_main/tests/test_sample_queue_helpers.py`, and `workflow_summary.json` evidence from `tests/test_measurement_worker.py` | Interruption persistence, explicit software recovery choices, and per-run workflow phase summaries exist, but hardware-driven recovery paths need field replay evidence. |

## Pending simulation-only behavior (explicitly identified)

- No-comm / fallback behavior is intentionally retained in:
  - `diagnostic_services.VacuumNoCommBackend`
  - `diagnostic_services.IrmArmNoCommBackend`
  - `diagnostic_services.SquidNoCommBackend`
  - `diagnostic_services.DCMotorNoCommBackend`
  - `hardware_contracts.NoCommBackend`
- `queue/changer` path still accepts adapter-less defaults where no transport is available:
  `QueueHardwareBackend` and `build_measurement_backend` remain operationally safe but may mask missing hardware at startup.

## Offscreen/multi-monitor evidence requirements to complete

- Confirm launch across at least the following geometric classes with hardware attached:
  - Native multi-monitor primary/secondary boundary transitions.
  - Taskbar-offset available geometry with top/left non-zero origin.
  - Tiny displays below 1366x768.
  - Mixed-DPI profiles (1.25, 1.5, 2.0 scaling) using same monitor rectangle assumptions.

## Physical-hardware acceptance checklist (to mark as complete)

1. AF demo and queued AF workflow runs with all three supported AF hardware transports.
2. IRM/ARM calibration path applies axis/ramp values and resets hardware state reliably.
3. SQUID transport connects in the configured serial profile and reports stable moments.
4. Vacuum transitions (pump on/off + pressure dynamics) show expected closed-loop behavior.
5. DAC/MCC/ADWIN command execution and limits are verified against known-good VB6 parameters.
6. DC motor dialog and measurement path execute: command + encoder + torque readback on one axis at least.
7. Interrupted queue run: crash/kill/restart test and resume from persisted queue state on same station.
8. Main-window startup and restore in:
    - full-HD and 4K monitors,
    - dual-monitor with negative/offset origins,
    - small/low-resolution displays.

## Explicit unclosed gap list against the June-to-July objective

- Original unassessed workflow components now have mapped, integrated-foundation, or retired statuses in `docs/vb6_parity_inventory.md`; remaining closure is acceptance evidence rather than unknown inventory.
- 42 partially mapped component behaviors still need acceptance artifacts or explicit retirement evidence.
- No-op/No-Comm fallbacks remain intentionally preserved in hardware-offline scenarios and must be clearly distinguished from production-hardened paths.
