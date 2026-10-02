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

Full suite: **587 tests passing** (`python -m unittest discover -s tests -p
'test_*.py'` from `RapidPy/rapid_main`, with `RapidPy`, `RapidPy/updown_control`,
and `RapidPy/vrm_logger` importable), up from a 261-test baseline. The
October 2 shell slice adds truthful Dashboard backend snapshots, responsive
glass-card reflow, workflow/session menu wiring, shutdown ordering, atomic
sequence documents, truthful empty operator states, real sample-index loading,
AF simulation isolation, atomic settings backup/restore, and searchable VB6
transition-help tests. It also covers immutable calibration artifact snapshots,
versioned approvals, expiry/invalidation, event-based rollback, integrity
checks, and measurement-provenance linkage.
The current distribution checkpoint also builds an installable
`berkeley-rapidpy` wheel, excludes checkout bytecode, packages application
icons, exposes seven console entry points, diagnoses dependencies/configuration
headlessly, and constructs the main window from an isolated install target.
The first reachable-dialog completion slice also moves Vacuum, IRM/ARM, and
SQUID onto the shared glass/semantic-status contract, adds named accessibility
metadata, verifies compact 360x520 geometry, and keeps unavailable hardware
fail-closed rather than presenting it as ready.
The second slice moves Login, About, Quick Start, and Transition Help onto the
same shared dialog system, removes rigid widths/heights, verifies compact action
visibility, and adds useful accessible names to every primary control.
The third and final slices cover Plots, Sample Selection, Webcam, Debug Console,
Step Monitor, and DC Motors. They add truthful real/simulated plot state,
fail-closed unavailable motor controls, written semantic execution/connection
states, accessible safety descriptions, compact action coverage, and correct
`QtGui.QScreen` fitting. Opening Debug Console now uses PySide6's supported rich
text API and escapes backend-provided text; disconnecting DC Motors now refreshes
all enabled states rather than leaving motion controls active.

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
- Installed builds resolve helper applications through packaged modules rather
  than repository-relative scripts; checkout scripts remain the development
  preference. A dependency-clean/operator-account deployment is still pending.
- Sample measurement is blocked when no valid holder correction exists.
- Measurement output is published only when a run completes; an aborted run
  leaves the production output path untouched.
- Simulated runs publish their core data and every workflow, communication,
  susceptibility, artifact-index, and quicklook sidecar into a `SIMULATED`
  subdirectory with explicit provenance. A live-declared backend that returns a
  simulated block is rejected before accepted state or output can change.
- Sig/Holder and Sig/Induced now show real values when a bracketed block is
  available, plus holder identity, magnitude, and age.
- The main Dashboard reads the same backend snapshot as the debug console and
  visibly distinguishes live, disconnected, unavailable, faulted, and
  **SIMULATED** devices. It refreshes on demand, every ten seconds, and after a
  communication-mode change.
- Vacuum, IRM/ARM, and SQUID diagnostics now use shared dialog surfaces and
  textual `READY`, `SIMULATED`, `WARNING`, `ERROR`, and `UNAVAILABLE` state
  prefixes rather than color-only communication. Unavailable IRM/ARM actuation
  controls are disabled, and the Vacuum diagnostic remains open to explain an
  unavailable backend instead of failing during construction.
- Login, About, Quick Start, and VB6 transition help use the same responsive
  glass surfaces. Operator identity remains mandatory, No-Comm explains its
  simulation-only meaning to assistive technology, and an empty transition-help
  search disables its otherwise inapplicable Open Destination action.
- Plots, Sample Selection, Webcam, Debug Console, Step Monitor, and DC Motors
  complete the reachable-dialog shared-glass audit. Simulation/unavailable and
  step execution states are written in words, motor motion controls fail closed,
  and diagnostic text is rendered safely without interpreting backend markup.
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

- **VB6 functionality parity for RapidPy package:** **Not fully complete.** `vb6_parity_inventory.md` still tracks physical acceptance for SQUID, vacuum, AF/IRM/ADwin, motors, rockmag, manual-external thermal results, susceptibility bridge/coil/standard acceptance, DAC/MCC, and retained auxiliary instruments. Automatic email dispatch (`frmSendMail`) is explicitly retired in favor of operator-controlled sharing of deterministic measurement bundle outputs. Current status is best described as **software-hardened across the core SQUID, vacuum, treatment, and motor paths, with named hardware and auxiliary-integration gates**, not end-to-end production parity.
- **16 app windows displaying/initializing safely on all monitors:** partially complete. The shared boundary guard and startup contract coverage is present for 16 `app.py`/`main.py` entrypoints, and software now re-fits top-level windows after screen add/remove, work-area, geometry, and logical-DPI changes. Runtime validation against real multi-monitor + DPI-change permutations is still a physical evidence gap.
- **Auto-resize/layout across Windows machines:** partially complete. The shared geometry clamp is global (`rapidpy_common/ui.py`), `rapid_main` adds compact restore logic, and offscreen profile tests cover negative origins, small/tiny displays, scaled DPI, and topology re-clamping. Native field validation still requires actual display hardware.
- **Main app fully functional on every aspect:** **No.** Workflow scaffolding exists, but multiple critical paths are still operator- or hardware-acceptance-gated:
  - no-comm / simulator behavior is intentionally preserved in several services,
  - queue/changer and diagnostic hardware transitions need live transport acceptance evidence,
  - thermal/rockmag/VRM/scientific plotting domains now have mapped or integrated software foundations, but still require live hardware/run-bundle acceptance where retained. Specimen thermal treatment is explicitly manual/external because the active VB6 source supplies no furnace protocol; automated thermal and other treatment families without a validated production actuator are rejected during queue plan preflight rather than silently acknowledged.
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
| AF workflows | **Software-complete live adapter evidence; physical acceptance pending** | `MainWindow._prepare_af_workflow`, `diagnostic_services.AfDemagBackendAdapter`, `plan_af_demag_command`, `QueueHardwareBackend.set_demag_step`, `hardware_contracts.py`, `tests/test_app_af_workflow.py`, `tests/test_diagnostic_services.py`, `tests/test_queue_hardware_backend.py` | AF labels route through the dedicated ADwin planner/backend. Hardware mode requires a nonzero ADwin version probe, records complete structured ramp request/result/error evidence, rejects incomplete results, confirms safe relay reset, and contributes current-run events to `communication.tsv`; simulated sources are excluded. Synthetic AF examples remain no-communication-only. Physical rig execution, interlock, decay, abort, and transcript acceptance remain gates. |
| SQUID communication/plots | **Software-complete transport recovery; physical acceptance pending** | `acquisition.py`, `squid_transport.RawSquidTransport`, `squid_transport.BracketedSquidBackend`, `hardware_contracts.QueueHardwareBackend`, `measurement_worker.py`, `updown_control.RawSquidClient`, `tests/test_acquisition.py`, `tests/test_squid_transport.py`, `tests/test_measurement_worker.py` | Production bracketed acquisition records exact raw TX/RX/errors, rejects stale/partial replies, and classifies transport failures separately from motion failures. A transport fault returns to verified zero, clears counters, waits with bounded exponential backoff, and reacquires a newly reset/latching whole block; exhaustion and recovery failure return no block. The unsafe uncancellable eight-second outer read timeout was removed. Current-run recovery commands, timestamps, and failure details appear in operator warnings, provenance, `workflow_summary.json`, and `communication.tsv`. Simulated adapters cannot satisfy live evidence. Physical fault injection and transcript acceptance remain open. |
| Magnetometer normalization (`modMagnetometer`) | **Integrated foundation** | `magnetometer.py`, `measurement_worker.py`, `tests/test_magnetometer.py`, `tests/test_measurement_worker.py` | Raw SQUID axis voltages can be background-corrected and converted to calibrated moment vectors with quality flags, and the measurement loop accepts `MagnetometerReading` payloads while surfacing read-quality flags as operator warnings. Production bracketed transport recovery is software-verified; live SQUID/reference acceptance remains open. |
| Vacuum | **Software fail-closed command path; physical/pressure acceptance pending** | `updown_control.VacuumController`, `diagnostic_services.VacuumBackendAdapter`, `measurement_worker.py`, `dialogs/vacuum.py`, `tests/test_vacuum_transport.py`, `tests/test_diagnostic_services.py`, `tests/test_measurement_worker.py` | Hardware mode never falls back to simulation and records exact commands, CR-terminated acknowledgments, and errors. The legacy controller has no pressure telemetry, so RapidPy never fabricates live pressure: a positive threshold blocks, while threshold zero plus confirmed pump-on is the explicit command-only opt-out. Current-run vacuum events can join `communication.tsv`. Physical pump/valve round-trip, independent pressure adapter/interlock if retained, and fault injection remain open. |
| IRM/ARM calibration | **Software-complete ADwin treatment evidence; physical acceptance pending** | `diagnostic_services.IrmArmBackendAdapter`, `IrmArmNoCommBackend`, `hardware_contracts.py`, `dialogs/irm_arm.py`, `tests/test_diagnostic_services.py`, `tests/test_queue_hardware_backend.py` | Live construction requires an ADwin readiness probe. Every incremental IRM/ARM ramp records its complete structured request, result, or propagated error; incomplete result evidence fails closed. Safe relay reset requires confirmation, and live current-run events merge into `communication.tsv` while simulated sources are excluded. Physical calibration, interruption, safe abort, and execution acceptance remain open. |
| DAC/MCC / AF ramping | **Integrated ADwin ramp evidence; MCC/DAC acceptance pending** | `rapidpy_common/adwin_af.py`, `diagnostic_services.AfDemagBackendAdapter`, `diagnostic_services.IrmArmBackendAdapter`, `tests/test_diagnostic_services.py` | ADwin ramp requests/results and safe relay reset are validated and auditable in software. Production MCC/DAC binding, electrical loopback, limit/interlock behavior, and physical fault acceptance remain open. |
| DC motor control | **Software-complete transport/state evidence; physical acceptance pending** | `rapidpy_common.hardware.MotorSerialClient`, `diagnostic_services.DCMotorBackendAdapter`, `hardware_contracts.QueueHardwareBackend`, `dialogs/dc_motors.py`, `dc_motor_control/app.py`, `tests/test_motor_transport.py`, `tests/test_diagnostic_services.py`, `tests/test_queue_hardware_backend.py`, `tests/test_dcmotors_telemetry.py` | The shared Quicksilver boundary records exact CRLF commands, CR-terminated replies, and errors; empty, unterminated, non-ASCII, and malformed replies fail closed. Partial connection setup closes the port. Diagnostic and queue paths reject unsuccessful move results before state mutation. Safe return attempts AF reset, drop-off, and every axis halt, then reports a combined error instead of skipping later halts. Current-run motor evidence joins `communication.tsv`. Telemetry and accessible glass controls remain integrated. Physical direction, encoder/torque, limit/stall, transfer, safe-abort, and transcript acceptance remain open. |
| VRM logging | **Integrated via external launch + run manifest sidecar** | `MainWindow._launch_vrm`, `rapid_main/vrm.py`, `vrm_logger/main.py`, `vrm_logger/session_manifest.py`, `tests/test_app_bootstrap_window_contracts.py`, `tests/test_vrm_integration.py` | Functionality remains in the dedicated app module; rapid_main now writes a launch handoff manifest and passes it to `vrm_logger`, which writes a `.vrm.json` sidecar beside the selected CSV with session/output metadata. Live logged acquisition with physical run association remains an acceptance gate. |
| Interrupted queue recovery | **Integrated + explicit operator choices + workflow summary artifact** | `rapid_main/panels/sample_queue.py`, `rapid_main/dialogs/step_monitor.py`, `measurement_worker`, queue state restore, `tests/test_queue_orchestration.py`, `tests/test_sample_queue_helpers.py`, `tests/test_measurement_worker.py` | Queue interruption persistence exists; restored `Running` rows become `Interrupted` and require an operator choice of Resume, Re-run, Skip, or Abort before queue start. Measurement workers now write `workflow_summary.json` phase evidence. Field replay and stressed restart behavior still need hardware acceptance evidence. |
| Thermal / scientific plotting | **Manual-external planning and run evidence integrated; automated actuation blocked** | `thermal.py`, `hardware_contracts.py`, `measurement_worker.py`, Sequence/Measurement/Calibration panels, `docs/thermal_integration_decision.json`, `tests/test_thermal.py`, `tests/test_sequence_thermal.py`, `tests/test_calibration_panel.py`, `tests/test_queue_and_bundle.py`, `tests/test_measurement_worker.py` | The source-backed decision distinguishes VB6 AF coil sensor protection from specimen furnace control and records manual/external-only treatment until authoritative hardware/protocol evidence exists. Atomic `rapidpy.thermal.plan.v1` artifacts load into Sequence with explicit UI wording; exact plan identity reaches provenance/workflow and indexed `rapidpy.thermal.run.v1` completion/abort evidence. Hardware mode blocks before ordinary preflight without a furnace adapter; No-Communication execution is explicitly simulated and claims no furnace control. Manual-external physical result reconciliation remains open; automated control additionally requires controller documentation, interlocks, safe abort/cooldown, and signed bench acceptance. |
| Plotting (measurement/stats) | **Integrated + PCA/decay review + validated exports** | `analysis.py`, `dialogs/plots.py`, `panels/measurement.py`, `data_viewer` app, `tests/test_analysis.py`, `tests/test_plots_dialog.py`, `tests/test_measurement_panel.py`, `tests/test_output_integrity.py` | Measurement panel opens Zijderveld/equal-area/intensity review plus an accessible Analysis tab backed by the shared principal-axis and moment-decay services. North/East/Down PCA direction, variance fraction, RMS residual, moment statistics, decay ratios, monotonicity, and slope are included in atomic provenance-marked JSON/CSV exports. `artifact_index.json` is reconciled after quicklook publication. Simulated/replay quicklooks remain inside `SIMULATED`; live-declared backends cannot inject simulated blocks. Additional legacy-specific surfaces and real-bundle review remain acceptance gates. |
| Email/report dispatch | **Retired with replacement** | `io/measurement_bundle.py`, `measurement_worker.py`, `docs/vb6_parity_inventory.md` | VB6 automatic mail dispatch is not retained; operator reporting uses `.sample`, `.rmg`, `measurements.txt`, and `specimens.txt` bundle outputs for review/sharing. |
| Status/error code taxonomy | **Mapped** | `status_codes.py`, `measurement_worker.py`, `tests/test_status_codes.py` | `MeasurementWorker` now emits structured `OperatorStatus` objects for workflow phases, preflight warnings, SQUID errors, file/output errors, and safety-return states while preserving current UI strings. |
| Debug diagnostics (`frmDebug`) | **Integrated foundation** | `dialogs/debug_console.py`, `diagnostic_services.collect_diagnostic_status`, `tests/test_diagnostic_services.py`, `tests/test_remaining_dialog_glass.py` | Debug Console can refresh a structured readiness snapshot for Vacuum, AF Demag, IRM/ARM, SQUID, and DC Motors, including connected/simulated/status/fault state. The dialog uses the shared glass/accessibility contract and safely escapes backend text before rich-text rendering. Production SQUID, vacuum, ADwin treatment, and DC motor traffic is retained in run bundles; equivalent retained-auxiliary evidence plus physical troubleshooting remain gates. |
| Data analysis primitives (`modDataAnalysis`) | **Mapped** | `analysis.py`, `tests/test_analysis.py` | RapidPy now extracts measurement vectors from `MeasurementStep.sdx/sdy/sdz`, reports centroid/PCA-style principal-axis line-fit evidence, orientation/variance/residual metrics, and moment statistics/decay summaries. Advanced plotting integration and hardware-run bundle review remain acceptance caveats. |
| Communication listener/logger (`modListenAndLog`) | **Integrated + core live-adapter evidence** | `communication_log.py`, `squid_transport.py`, `diagnostic_services.py`, `hardware_contracts.py`, `measurement_worker.py`, `rapidpy_common/hardware.py`, `tests/test_communication_log.py`, `tests/test_squid_transport.py`, `tests/test_motor_transport.py`, `tests/test_diagnostic_services.py`, `tests/test_measurement_worker.py`, `tests/test_queue_hardware_backend.py` | Per-run `communication.tsv` accepts immutable, current-run evidence from production SQUID, legacy vacuum, AF/IRM ADwin, and DC motor paths. Quicksilver logging retains exact commands/replies and parse/transport errors; ADwin requests/results are structured and complete; event sources merge chronologically and simulated sources are excluded. Remaining software coverage is retained auxiliary adapters; physical transcript acceptance remains open. |
| Susceptibility (`frmSusceptibilityMeter`, `modSusceptibility`) | **Automated acquisition code-complete; physical acceptance pending** | `susceptibility_transport.py`, `susceptibility_acquisition.py`, `diagnostic_services.SusceptibilityBackendAdapter`, `dialogs/susceptibility.py`, `config.py`, `legacy_ini.py`, `holder_state.py`, `holder_measurement.py`, `hardware_contracts.py`, `measurement_worker.py`, `docs/susceptibility_integration_decision.json`, `tests/test_susceptibility_transport.py`, `tests/test_susceptibility_acquisition.py`, `tests/test_susceptibility_queue.py`, `tests/test_diagnostic_services.py`, `tests/test_hardware_dialog_glass.py`, `tests/test_measurement_worker.py` | The dedicated glass dialog now replaces the previously misrouted SQUID window. The typed transport sends exact legacy `Z`/`M` + CRLF commands, requires CR-terminated ASCII replies, rejects empty/partial/non-ASCII/non-numeric/non-finite data, applies the configured scale factor, and exposes immutable TX/RX/error evidence. VB6 port/framing/scale/moment/coil/enabled settings migrate into persisted configuration. The live SQUID adapter no longer mislabels magnetic moment as susceptibility; the worker reads only explicit `SUSC` steps and aborts a requested read failure instead of silently writing zero. Live SUSC now follows the VB6 zero / slow move to `Int(SCoilPos + SampleHeight/2)` / measure / `(sample - holder) * MomentFactorCGS` sequence with a verified home before and after (a labelled RapidPy safety improvement). One bridge instance is shared between the dialog and the queue under the ownership manager and connects only inside an operator-started run. Holder commands stage the bridge value with the magnetic block and install both atomically with persisted evidence; any failure keeps the previous holder. The worker reads before treatment and SQUID, honors halt cooperatively, and publishes every current-run acquisition as an indexed immutable `rapidpy.susceptibility.acquisition.v1` artifact; failures never become values or output, and safe-return failure is a distinct error. Plan preflight still blocks without an enabled live bridge, valid geometry/factors, and an accepted holder susceptibility. Bridge, coil-centring, reference-standard, and mechanical acceptance remain open (procedure section 11a). |
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
4. Vacuum pump/valve commands receive the expected physical acknowledgments; if pressure gating is retained, an independent pressure adapter/interlock proves pressure and controlled fault-to-halt behavior.
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
