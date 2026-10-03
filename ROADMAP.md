# RAPID Modernization Roadmap

## Mission

Fully migrate every actively used capability of the RAPID VB6 control program into one production-ready `RapidPy/rapid_main` application. The final application must preserve the scientific behavior, hardware safeguards, terminology, and workflows that experienced VB6 operators rely on while providing a clean, modern, fast, and lightweight interface.

The finished application must remain usable across supported monitor sizes and Windows display-scaling settings. Controls must not overlap, clip, or become too compressed to operate safely. Routine instrument operation, calibration, diagnostics, data review, and recovery must be possible from `rapid_main` without requiring separate subsystem applications.

The standalone RapidPy applications remain valuable reference implementations and bench tools during migration. Their proven hardware behavior should be moved behind shared services and integrated into `rapid_main`, not independently reimplemented in its widgets.

## Current Baseline

RapidPy already contains substantial standalone implementations for the gaussmeter, VRM logging, ADwin communication, COM mapping, XY changer, Z-axis and vacuum control, AF tuning, AF clipping, data viewing, and settings editing. The repository also contains shared Quicksilver motor, ADwin AF, gaussmeter, configuration, and UI code.

`rapid_main` has a useful application shell, sample and sequence models, queue compilation, runtime estimation, dialogs, and CIT/RMG/MagIC input and output support. It also contains a hardware-independent measurement-worker design and a simulation backend.

The principal remaining gap is integration and qualification. The unified application does not yet orchestrate the real hardware through a complete production workflow. Some controls, plots, coordinate calculations, error calculations, and device dialogs remain placeholders or simulation-oriented. Automated test coverage is concentrated in file I/O, queue logic, and Data Viewer analysis rather than hardware behavior and end-to-end control.

The current baseline should therefore be understood as:

- Strong standalone subsystem and diagnostic progress.
- A useful shared hardware foundation that still needs formal contracts and tests.
- A broad unified-application prototype with incomplete real-hardware orchestration.
- Insufficient evidence to retire VB6 or run the full instrument autonomously.

## Non-Negotiable Product Requirements

### Complete VB6 capability parity

Every capability in the active `VB6/Paleomag v3.vbp` project must be:

1. Migrated into `rapid_main`.
2. Deliberately replaced by a documented and validated improved workflow.
3. Or formally classified as obsolete and approved for exclusion.

No capability may disappear merely because it was overlooked. Old, duplicate, and superseded VB6 files should be recorded during the audit, but they do not require separate Python implementations when their active behavior is already represented elsewhere.

### Low-friction transition for VB6 operators

- Preserve recognizable task names, sample terminology, sequence labels, machine positions, calibration concepts, and safety language.
- Offer a Classic workflow navigation option or a clearly visible VB6 task map.
- Provide a searchable "Where did this VB6 control go?" reference.
- Import existing VB6 INI files, calibrations, sample definitions, sequences, and supported result files.
- Show physical units alongside legacy raw controller values when both are useful.
- Provide a transition sheet for every migrated VB6 form or operator workflow.
- Explain intentional differences at the point of use rather than only in developer documentation.
- Include a simulator-backed practice mode for operator training.

Scientific or mechanical behavior must not change solely to modernize the interface. Any intentional behavioral change requires documentation, validation, and operator approval.

### Clean and responsive user interface

- Use layout managers, splitters, tabs, drawers, and scroll areas instead of fixed-position interfaces.
- Never overlap, clip, or shrink safety-critical controls below a usable size.
- Adapt multi-column pages to fewer columns when space is constrained.
- Collapse secondary controls and diagnostics before compressing primary workflow controls.
- Keep Pause, Halt, system state, current sample, and current step visible during active runs.
- Resize tables intelligently while preserving access to important columns through scrolling.
- Support keyboard navigation and visible focus states.
- Do not use color as the only indicator of hardware state or severity.
- Persist workstation-specific window and splitter arrangements and provide a Reset Layout action.

The required display test matrix is:

- Minimum supported workspace: 1280 x 720 at 100% scaling.
- Standard workspace: 1920 x 1080.
- Windows display scaling at 100%, 125%, 150%, and 200% where supported by the target hardware.
- Reduced-height laptop displays.
- Ultrawide and multi-monitor configurations.

### Fast and lightweight operation

- Hardware I/O, file watching, camera capture, plotting, and long calculations must not block the UI thread.
- Optional tools must load only when opened.
- Polling and plot refresh rates must be bounded and configurable.
- Long-running logs, camera buffers, tables, and plots must have bounded memory growth.
- Dependencies must be justified by essential functionality.
- The final Windows package must run without access to the source repository.
- Startup, idle CPU use, memory use, and long-run responsiveness must be measured on the oldest supported control computer.

### Safety and data integrity

- All hardware commands must pass through shared services and centralized interlocks.
- Device ownership must be exclusive so two panels cannot command the same hardware concurrently.
- Every automated operation must have preconditions, timeouts, cancellation behavior, and a defined safe outcome.
- The software must never guess a specimen or mechanism position after losing authoritative state.
- Measurement output must be written atomically and associated with the exact configuration and calibration versions used.
- Pause, controlled Halt, and emergency-stop behavior must be defined separately and tested.

## Capability Migration Map

| Functional area | Principal VB6 sources | Target in `rapid_main` |
|---|---|---|
| Main operation and flow | `frmMagnetometerControl`, `modFlow`, `modProg` | Dashboard, run controller, application state machine |
| Operator session | login, messages, help, splash, and shutdown forms | Operator session, contextual help, safe shutdown |
| Sample management | changer order, sample selection, registry, rerun, and queue forms | Sample workspace and persistent queue |
| Programs and sequences | `frmProgram`, sample commands, measurement blocks | Sequence editor, validator, and templates |
| Measurement | `frmMeasure`, `modMeasure`, `frmStats` | Live Measurement workspace and measurement service |
| SQUID | `frmSquid`, magnetometer modules | SQUID service, live readings, calibration, diagnostics |
| XY changer | changer and XY-homing forms and modules | Integrated changer workspace and service |
| Z and turning motors | `frmDCMotors`, `modMotor` | Integrated motion workspace and service |
| Vacuum | `frmVacuum` | Vacuum service and compact operator panel |
| AF treatment | AF, ADwin AF, tuner, clipping, and calibration forms | AF treatment, tuning, and calibration workspaces |
| IRM and ARM | `frmIRMARM`, IRM voltage calibration | IRM/ARM treatment and calibration workspaces |
| Thermal | `modThermal` | Thermal workflow and device service |
| Susceptibility | susceptibility form and module | Susceptibility measurement workflow |
| Gaussmeter | 908A form and module | Integrated gaussmeter tool and meter service |
| Coil and rod calibration | coil and calibration-rod forms | Calibration Center |
| Rock magnetics | rockmag routine and data classes | Rock-magnetic sequence support |
| VRM | `frmVRM` | Integrated VRM acquisition and plotting |
| Plots and analysis | plot forms and data-analysis modules | Live plots and review workspace |
| Webcam | `frmWebcam` | Dockable camera view |
| Configuration | settings, options, DAQ, channel, and communication forms | Typed settings and hardware configuration |
| Files and export | file-save and directory forms and modules | Project/output manager and atomic data writing |
| Diagnostics | debug, DAC/ADwin communications, step monitor | Diagnostics and role-controlled service mode |

## Delivery Status Model

Every capability in the migration matrix must use one of these evidence-based statuses:

- **Not assessed**: behavior and dependencies have not been inventoried.
- **Designed**: destination architecture and acceptance criteria are approved.
- **Standalone implemented**: behavior exists in a subsystem app but is not integrated.
- **Integrated**: behavior is accessible through `rapid_main` and shared services.
- **Simulator-tested**: deterministic normal and failure cases pass without hardware.
- **Bench-tested**: direct device behavior passes on the bench.
- **Operator-validated**: experienced operators have completed the workflow successfully.
- **Production-qualified**: safety, recovery, scientific results, packaging, and endurance gates pass.
- **Approved obsolete**: exclusion is documented and approved.

A compiled executable or polished UI does not by itself qualify a capability for production.

## Current Execution Snapshot (Goal-Driven Run)

**As of Friday, July 10, 2026:**

- Active goal: execute roadmap phases in order while preserving low-friction VB6 transition constraints.
- Phase 1: inventory artifact exists and is linked from this roadmap.
- Phase 2: measurement hardware contract is in place and enforced in the measurement worker preflight path.
- Immediate next milestone: close the remaining contract/UI integration gaps, then harden AF treatment state and preflight-driven workflow behavior.
- Regression policy for this execution run: every executed roadmap item should add/extend tests and a short evidence artifact under `docs/` or `rapid_main/tests/`.

### Execution checklist for this run

1. Confirm and update `docs/vb6_parity_inventory.md` row statuses to move any `Not assessed`/`Mapped` items with owners.
2. Add service-level ownership checks so two widgets cannot command the same device.
3. Expand preflight checks with timeout and cancellation semantics per real backend implementations.
4. Run deterministic simulator scenarios for AF and output bundle writes in CI-style test execution.
5. Capture operator review notes and archive as an execution artifact before advancing to broader calibration automation.
6. Begin Phase 3 execution with responsive-shell reliability: persist layout and add a reset layout path.
7. Add queue-safe validation for sample-order workflows (`frmChangerSampOrder` equivalent): duplicate/invalid-order detection and strict-mode blocking for invalid queue definitions.
8. Wire queue execution from `SampleQueuePanel` through `MainWindow` into `MeasurementPanel` for automatic per-sample runs, including run/warn/cancel transitions and sample status updates.
9. Make queue control ergonomics and safety actions deterministic: connect toolbar queue pause/halt and app header controls to shared measurement queue pause/abort paths; ensure halt aborts both automation and active worker.
10. Finish the VB6 configuration migration workstream by mapping `CIniFile` and `frmINIConverter` parity into `rapid_main` (settings import workflow, mapping report, and warnings for unmapped keys).

Execution status on this run:

- Step 1: **In progress**. Parity inventory and remediation tracking is continuing; `frmShutdownMsg` is now mapped to the controlled shutdown flow in `MainWindow._request_shutdown` / `_confirm_shutdown`, while `CIniFile` / `frmINIConverter` are already covered by the `Import VB6 INI…` workflow.
- Step 1a: **Implemented**. Sample-index migration is now mapped/evidenced in inventory (`SampleIndexRegistration`, `SampleIndexRegistrations`, `frmSampleIndexRegistry` parity path).
- Step 2: **Implemented**. Shared ownership now covers measurement workflows and the remaining shell-level diagnostic launchers: measurement execution uses ownership-aware panel locking, and SQUID/IRM/Vacuum/DC Motors dialogs all route through the same resource lock path.
- Step 3: **Implemented**. Measurement preflight + step commands now run with halt-aware timeout guards, and backend availability is explicitly checked before run start.
- Step 4: **In progress**. AF workflow now has canonical transition-aware phase tracking (`positioning`, `validating`) plus safe-state completion via `return_to_safe_state()` on all exits; `AF Demag` diagnostics now also provides a runnable demo path (`Run AF Demo Sequence`) that loads AF defaults, sets specimen context, and starts the run in Live Measurement.
- Step 5: **Implemented (internal checkpoint)**. Required operator-review evidence is now archived in `docs/operator_review_notes_2026-07-10.md`; simulation paths remain validated and human operator sign-off is pending.
- Step 6: **Implemented**. Main window layout and sidebar width now persist across runs via `QSettings`; `Reset Layout` is available under `View`.
- Step 7: **Implemented**. Queue compilation now validates sample metadata and duplicate positions before command synthesis through `validate_queue_samples`, with strict-mode failure paths covered by tests.
- Step 8: **Implemented**. Queue execution now runs through shared app flow: queue rows validate, options compile into `QueueOptions`, `MainWindow.start_queue_run()` advances each sample through `MeasurementPanel.start_measurement_for_sample()`, and row statuses update from pending→running→done/error with warnings/abort support.
- Step 9: **Implemented**. Queue control actions are now plumbed to shared runtime state: app-header and queue toolbar Pause/Halt buttons are wired to queue-aware measurement control methods, and queue halt now requests measurement worker halt before clearing queue state.
- Step 10: **Implemented**. VB6 INI migration workstream completed in this phase via `legacy_ini.py` + Settings import flow, with mapping report, unmapped-key warnings, and test-backed importer behavior.
- Step 11: **Implemented**. Sample-index registry migration completed for sample-loading parity:
  - Added typed registry models (`SampleIndexRegistration`, `SampleIndexRegistrations`) and `.sam` parser path in `rapid_main`.
  - Routed `SampleSelectDialog` to the new registry parser and preserved existing specimen metadata fallback behavior.
  - Added fixture-backed regression coverage for VB6-style headers and plain-list `.sam` registry files.

Progress evidence for this checkpoint:

- [Execution notes for this checkpoint](/docs/roadmap_execution_notes_2026-07-10.md)
- [Operator review checkpoint notes](/docs/operator_review_notes_2026-07-10.md)
- `RapidPy/rapid_main/rapid_main/app.py`
- `RapidPy/rapid_main/rapid_main/measurement_worker.py`
- `RapidPy/rapid_main/rapid_main/panels/measurement.py`
- `RapidPy/rapid_main/tests/test_measurement_worker.py`
- `RapidPy/rapid_main/tests/test_measurement_panel.py`
- `RapidPy/rapid_main/tests/test_device_ownership.py`
- `RapidPy/rapid_main/tests/test_queue_and_bundle.py`
- `RapidPy/rapid_main/tests/test_workflow.py`
- `RapidPy/rapid_main/rapid_main/workflow.py`
- `RapidPy/rapid_main/rapid_main/hardware_contracts.py` (timeout metadata + safe-return hook default)
- `RapidPy/rapid_main/rapid_main/panels/sample_queue.py`
- `RapidPy/rapid_main/tests/test_sample_queue_helpers.py`
- `RapidPy/rapid_main/rapid_main/legacy_ini.py`
- `RapidPy/rapid_main/rapid_main/panels/settings_panel.py`
- `RapidPy/rapid_main/tests/test_legacy_ini.py`
- `RapidPy/rapid_main/rapid_main/app.py` (shutdown safety flow and controlled exit path; AF demo launch path and runnable AF demo action)
- `RapidPy/rapid_main/rapid_main/data_model.py` (`SampleIndexRegistration`, `SampleIndexRegistrations`)
- `RapidPy/rapid_main/rapid_main/io/sample_index.py` (registry parser + sample-order mapping)
- `RapidPy/rapid_main/rapid_main/dialogs/sample_select.py` (registry-backed load path)
- `RapidPy/rapid_main/tests/test_sample_index.py` (registry parser regression tests)
- `RapidPy/rapid_main/tests/test_app_af_workflow.py` (AF preset contract test)

## Phase 1 - Authoritative Parity Inventory

Create a traceability matrix for every active VB6 form, module, class, command, setting, data format, scientific calculation, and hardware interaction.

For each capability, record:

- VB6 source and operator purpose.
- Hardware dependencies and communication protocol.
- Inputs, outputs, settings, calibrations, and persisted data.
- Safety assumptions, interlocks, limits, and recovery behavior.
- Existing RapidPy implementation and known deviations.
- Target location in `rapid_main`.
- Migration status and responsible owner.
- Test method and acceptance criteria.
- Approved differences or reason for retirement.

Canonical evidence tracker for this phase is now:

- [VB6 parity inventory](/docs/vb6_parity_inventory.md)

Create representative golden fixtures from real VB6 settings, sequences, samples, calibrations, and outputs. Record the expected calculations and file representations before changing them.

### Exit criteria

- Every active VB6 component is mapped.
- Operators have reviewed the common task inventory.
- No feature is silently excluded.
- Golden input/output fixtures are version-controlled.
- Standalone implementations are mapped to reusable services.

## Phase 2 - Unified Hardware and Workflow Architecture

Move proven behavior into UI-independent services for:

- SQUID.
- X, Y, Z, and turning motors.
- Changer orchestration.
- Vacuum.
- ADwin and AF.
- IRM and ARM.
- Susceptibility.
- Thermal equipment.
- Gaussmeter and other digitally controlled meters.
- Webcam.
- Configuration and calibration.
- Measurement processing and output.

Each service must expose connection state, read-only status, validated commands, timeouts, cancellation, actionable errors, event logging, and both simulated and real-hardware implementations. Widgets must not open ports or command hardware directly.

Implement a central workflow state machine such as:

```text
Idle
-> Preflight
-> Loading
-> Treating
-> Positioning
-> Measuring
-> Validating
-> Saving
-> Returning
-> Complete
```

The detailed state machine must include failure and recovery branches and must distinguish physical state, requested state, and confirmed state.

### Exit criteria

- `rapid_main` can execute deterministic workflows against simulators.
- Each real backend passes a connection, status, command, timeout, and stop contract.
- Hardware services have no widget dependencies.
- Conflicting device ownership is prevented.
- Pause, Halt, and emergency-stop semantics are centralized.
- Configuration errors are detected before motion or treatment begins.

## Phase 3 - Responsive Application Shell and Design System

Build the final shell around a persistent top status bar, adaptive navigation, a central task workspace, an optional details/diagnostics drawer, and a collapsible event area. Allow plots, calibration tools, diagnostics, and the webcam to open as docked, full-screen, or detached views where useful.

Define reusable UI components and rules for:

- Connection and instrument-state indicators.
- Primary, secondary, destructive, and emergency actions.
- Parameter entry with units, ranges, and validation.
- Hardware preflight and operator confirmations.
- Responsive cards, forms, tables, plots, and split views.
- Empty, loading, disconnected, warning, fault, and recovery states.
- Classic terminology and modern terminology where they differ.

Automated UI smoke tests must resize important pages through the supported display matrix. Manual tests must cover real Windows scaling and the oldest supported monitor.

### Exit criteria

- No overlap or clipping occurs in supported display configurations.
- Primary workflows remain usable at 1280 x 720.
- Essential values and warnings are never silently truncated.
- Keyboard navigation and visible focus work throughout primary workflows.
- Operators can reset layout state.
- UI updates remain responsive during simulated long-running work.

## Phase 4 - Core Measurement Workflow

Implement the first complete production-shaped workflow inside `rapid_main`:

1. Start an operator session.
2. Select or create a project and output destination.
3. Import or enter samples and build a queue.
4. Select, edit, and validate a sequence.
5. Run hardware and configuration preflight.
6. Retrieve and position a sample.
7. Apply treatment.
8. Move to measurement position.
9. Read SQUID and optional susceptibility data.
10. Apply coordinate transformations and quality calculations.
11. Validate and atomically commit the measurement bundle.
12. Return the sample and advance the queue.
13. Support pause, halt, skip, rerun, and recovery.

Wire the existing measurement-worker concept into the main application and replace simulated countdown behavior with actual workflow state. Complete the live measurement plots, specimen-to-geographic/core transformations, noise/error calculations, and persistent run journal.

Begin with a fully validated AF-demagnetization specimen workflow because it exercises motion, treatment, SQUID measurement, data output, recovery, and operator interaction together.

### Exit criteria

- Simulated runs are deterministic and cover normal and failure cases.
- Supervised real-hardware AF runs complete entirely in `rapid_main`.
- UI progress and readings reflect confirmed hardware state.
- CIT/specimen, RMG, and MagIC output match accepted fixtures.
- Interrupted or rejected measurements are never represented as complete.

## Phase 5 - Complete Treatment and Measurement Migration

Integrate and qualify workflows incrementally:

1. AF demagnetization.
2. IRM and ARM.
3. Susceptibility measurements.
4. Thermal workflows.
5. Rock-magnetic routines.
6. VRM acquisition.
7. Remaining specialized measurement routines identified in Phase 1.

For every workflow:

- Preserve supported VB6 labels, parameters, and sequence semantics.
- Validate ranges before enabling execution.
- Show treatment progress and authoritative hardware state.
- Support timeout, cancellation, safe stop, recovery, and rerun.
- Record commands, readings, configuration, calibration, and software versions.
- Compare scientific and mechanical results with VB6 or an approved reference.

### Exit criteria

- Every active VB6 treatment and measurement type is integrated or approved obsolete.
- Operators do not need standalone apps for production runs.
- Unsupported or invalid commands cannot enter an executable queue.
- Each workflow passes simulator, bench, operator, and output-comparison gates.

## Phase 6 - Automated and Manual Calibration Center

Create a role-controlled Calibration and Diagnostics workspace that uses digital control of the instrument and meters to automate repeatable calibration procedures while retaining complete supervised manual operation.

### Calibration modes

- **Automated mode**: the application controls motion, treatments, meters, data collection, fitting, validation, and proposed configuration changes.
- **Guided mode**: the application advances through a checklist while the operator confirms physical setup and selected actions.
- **Manual mode**: authorized users directly operate individual devices, inspect readings, and perform custom calibration steps.
- **Simulation mode**: users can rehearse and test the complete procedure without commanding hardware.

Automated mode should be the default for validated routine calibrations. Manual mode remains available for troubleshooting, research, unusual configurations, and recovery.

### Shared automated calibration workflow

```text
Select procedure
-> verify authorization
-> load active calibration
-> verify instrument configuration
-> run hardware and safety preflight
-> request physical setup where required
-> establish baseline
-> command controlled calibration steps
-> collect synchronized meter readings
-> detect unstable or invalid readings
-> fit the calibration model
-> calculate uncertainty and residuals
-> compare with acceptance limits
-> present old and proposed calibration
-> obtain operator approval
-> save a versioned calibration
-> generate a report
-> verify with an independent check
```

All automated procedures must support configurable settling time and sample count, averaging, drift and outlier detection, automatic range selection where supported, failed-point retries, safe cancellation, defined resume checkpoints, live plots, estimated completion time, and automatic rejection when uncertainty or residuals exceed approved limits.

No new calibration may be applied silently. The application must display the old and proposed values, fit quality, uncertainty, operating range, and acceptance result before an authorized operator approves activation.

### Procedures to automate where hardware allows

- XY cup positions and changer geometry.
- Z measurement-position optimization using SQUID response.
- Turning-axis position and orientation checks.
- SQUID baseline, axis response, noise, drift, and range checks.
- AF coil frequency tuning.
- AF clipping and maximum safe-output determination.
- AF field-versus-command calibration using digital gaussmeter readings.
- Axial and transverse coil calibration.
- Transverse-probe angle and position optimization.
- IRM/ARM voltage-versus-field calibration.
- Gaussmeter zero, range, stability, and comparison checks.
- Susceptibility-meter baseline and reference-sample checks.
- Vacuum response, valve-state, and threshold verification.
- Calibration-rod measurement routines.
- Motor reference-position and homing repeatability.
- Temperature-sensor and thermal-system checks where digital control is available.
- End-to-end reference-specimen verification.

Automation must pause for physical handling or independent observation when it cannot verify those conditions digitally. A requested fixture change is not proof that the fixture is present.

### Manual and service controls

Authorized users must retain direct access to motor jog and homing, raw target movement, cup capture and editing, Z scans, live SQUID readings, AF relay/frequency/voltage/ramp/waveform controls, IRM/ARM control and readback, gaussmeter controls, susceptibility readings, vacuum controls, ADwin/DAQ communications, COM discovery, webcam setup, and raw consoles where genuinely necessary.

Manual commands must use the same limits, interlocks, exclusive device ownership, audit logging, and emergency-stop behavior as automated workflows. Manual mode is not an unlogged bypass around safety rules.

### Calibration records

Store the procedure and calibration ID, hardware identifiers, operator and approver, timestamp, software version, source configuration, previous calibration, raw commands and readings, setup notes, fit method, coefficients, residuals, uncertainty, accepted operating range, pass/fail result, verification result, plots, and override reasons.

Calibration records must be immutable after approval, versioned, reversible without destroying history, exportable in human- and machine-readable forms, linked to production measurements, capable of being marked expired or invalid, and compatible with imported VB6 values.

### Exit criteria

- Routine calibrations run automatically from preflight through report generation.
- Guided and manual alternatives exist for every supported procedure.
- All modes produce the same calibration record and audit format.
- Repeatability equals or exceeds the corresponding VB6/manual procedure.
- Acceptance thresholds and fitting methods are documented and tested.
- Failed or interrupted calibration cannot overwrite an active calibration.
- Operators explicitly approve proposed changes.
- Production measurements record exact calibration versions.
- Service actions cannot bypass interlocks or audit logging.
- Standalone calibration utilities are unnecessary for normal work.

## Phase 7 - Data Review, Plotting, and Scientific Parity

Integrate the useful lightweight Data Viewer capabilities into `rapid_main`:

- Live directional components.
- Zijderveld plots.
- Equal-area stereonets.
- Intensity decay.
- Three-dimensional vectors and PCA.
- AF and thermal plots.
- IZZI/Thellier quicklooks and checks.
- Measurement statistics, noise, and error indicators.
- Automatic refresh after committed measurements.
- CIT, RMG, legacy CSV, and MagIC compatibility.

Advanced review tools should load on demand so normal acquisition remains lightweight. Calculations must use tested, shared analysis functions rather than duplicating logic inside plots.

### Exit criteria

- All operational VB6 plots and statistics are represented.
- Plot updates do not block acquisition or device polling.
- Saved results reopen reproducibly.
- Scientific calculations pass fixture-based tests.
- Advanced views do not materially degrade routine run performance when closed.

## Phase 8 - VB6 Operator Transition

Provide:

- First-run import of VB6 INI, calibration, sequence, and supported sample files.
- Optional Classic navigation organized around familiar VB6 tasks.
- Tooltips and search aliases using old form and control terminology.
- A VB6-to-RapidPy task lookup inside the application.
- Side-by-side transition sheets.
- Simulator-backed practice workflows.
- Short guides for normal runs, reruns, halt, recovery, calibration, and shutdown.
- Contextual explanations for intentional changes.

Conduct structured usability sessions with both experienced VB6 operators and new users. Track task completion time, errors, questions, and points of confusion.

### Exit criteria

- Experienced operators complete standard workflows with minimal assistance.
- Every common VB6 task has a documented destination.
- Approved terminology is consistent across the application and manuals.
- Critical usability findings are resolved before VB6 retirement.
- Training and recovery exercises are available without live hardware.

## Phase 9 - Safety, Recovery, and Qualification

Build and execute a fault-injection matrix covering:

- Serial disconnects and incorrect port assignments.
- Motor timeouts, unexpected stops, and limit switches.
- Vacuum or valve failure.
- Invalid, noisy, saturated, or drifting SQUID readings.
- ADwin and treatment failures.
- Meter range and communication failures.
- File-write and disk-space failures.
- Application crash and workstation restart.
- Power interruption.
- Invalid or changed configuration.
- Partially completed queue and treatment.
- Unknown physical specimen or mechanism position.

Implement startup self-test, pre-run validation, atomic file writes, an append-only recovery journal, configuration/calibration snapshots, safe shutdown, explicit recovery checkpoints, and clear severity levels for warning, recoverable fault, run-stopping fault, and emergency state.

### Exit criteria

- The approved fault-injection matrix passes on simulators and relevant hardware.
- The application never guesses uncertain physical state.
- Recovery neither duplicates nor loses accepted measurements.
- Halt prevents subsequent treatment and motion.
- High-risk operations have verified preconditions, limits, and timeouts.
- Audit records explain what happened before and after each fault.

## Phase 10 - Performance, Packaging, Parallel Operation, and VB6 Retirement

Measure and optimize startup time, idle resource use, polling, plot refresh, camera buffers, log growth, queue endurance, and UI responsiveness on the oldest supported workstation.

Deliver:

- A reproducible PyInstaller build for `rapid_main`.
- A clean-machine Windows installation test.
- A signed and versioned installer where infrastructure permits.
- Configuration migration, backup, export, and rollback tools.
- Operator, calibration, recovery, and service manuals.
- A release checklist and hardware compatibility record.
- A documented rollback path during parallel operation.

Run representative samples and workflows through VB6 and RapidPy under matched conditions. Compare motion positions, treatment parameters, meter readings, SQUID results, coordinate calculations, timestamps, and output files against documented tolerances.

### Exit criteria

- Every traceability-matrix row is production-qualified or approved obsolete.
- Scientific and mechanical results meet documented tolerances.
- At least 20 successful supervised runs per major workflow are completed, followed by appropriate endurance testing for routine workflows.
- Operators pass normal-operation, calibration, and recovery exercises.
- The packaged application runs without the source repository or standalone apps.
- The lab formally approves `rapid_main` as the operational replacement.
- VB6 is archived and reproducible for historical reference but is no longer operationally required.

## Testing Strategy Across All Phases

Testing is cumulative and must include:

- Unit tests for protocols, conversions, calculations, validation, and data formats.
- Contract tests shared by simulated and real hardware services.
- Deterministic simulator tests for workflows and fault cases.
- UI smoke tests at supported sizes and scaling configurations.
- Golden-file tests for VB6-compatible and MagIC output.
- Bench tests for individual devices.
- End-to-end tests on the complete instrument.
- Fault-injection and recovery tests.
- Performance and long-duration queue tests.
- Operator acceptance and transition tests.

Continuous integration should run all hardware-independent tests. Hardware test results should be recorded in repeatable qualification reports rather than informal notes.

## Documentation Deliverables

- VB6 capability traceability matrix.
- Architecture and hardware-service contracts.
- Workflow state-machine diagrams.
- Configuration and calibration schemas.
- Operator manual and quick-start guide.
- VB6 transition sheets and task lookup.
- Calibration procedures and acceptance limits.
- Diagnostics and service manual.
- Fault recovery and safe-shutdown guide.
- Scientific validation report.
- Hardware qualification and compatibility record.
- Release, installation, upgrade, backup, and rollback instructions.

## Final Definition of Done

The modernization is complete only when:

1. Every actively used VB6 capability is available in `rapid_main` or formally retired.
2. The application safely controls the complete instrument across all approved workflows.
3. Existing settings, calibrations, sequences, samples, and supported data can be migrated.
4. Experienced VB6 operators can locate and perform familiar tasks without extensive retraining.
5. The UI remains readable, responsive, and non-overlapping across the supported display matrix.
6. Hardware activity, plotting, camera use, and file operations never freeze the interface.
7. Automated, simulator, bench, fault-injection, scientific-output, performance, and operator-acceptance tests pass.
8. Routine operation, automated or manual calibration, diagnostics, review, and recovery require only `rapid_main`.
9. A clean supported Windows computer can install and run the application without the repository.
10. Interrupted work preserves physical safety, auditability, and data integrity.
11. Production measurements identify the exact software, settings, and calibration versions used.
12. The lab formally approves `rapid_main` as the full replacement for the VB6 application.

## Immediate Next Milestone

Complete Phase 1 and the core contracts of Phase 2, then build the responsive shell and deliver the first complete AF measurement workflow through Phase 4.

The immediate milestone is complete when the project has an operator-reviewed parity matrix, golden VB6 fixtures, shared simulator-backed hardware contracts, responsive primary screens, and one supervised AF specimen run that executes entirely in `rapid_main` and produces validated output. This creates a production-shaped vertical foundation on which every remaining VB6 capability can be migrated without creating another generation of disconnected applications.

## RAPID-v4 Execution Plan (Current)

**Last updated:** 10 July 2026

### Executive objective

Execute this repository roadmap so that `RapidPy/rapid_main` becomes the practical replacement for `VB6/Paleomag v3.vbp`: complete VB6-capability parity, provide a responsive modern shell, support both automated and manual workflows, and deliver a validated AF workflow with safe stop/recovery before any production migration.

### Current capability status (as observed in current workspace)

| Area | `rapid_main` status | Evidence | What is still required |
|---|---|---|---|
| Shell & navigation | Partially complete | 5-panel shell, status bar, diagnostics menu, responsive side/main layout | Multi-resolution stress validation, persistent layouts, global compact mode, recovery-safe startup |
| Measurement sequencing | Partially complete | Sequence builder, label preview, file save/load, runtime estimates | Queue compiler integration, full per-step validation against device contracts, rerun/skip/abort flow |
| Live measurement execution | Partially complete | `MeasurementWorker`, run countdown, step/result signals, specimen output paths | Hardware command/state binding for each step, robust pause/halt semantics, progress persistence |
| SQUID acquisition | Simulated | `NoCommBackend` read path integrated | Real serial/SDK backend and quality/error checks |
| AF demag execution | Simulated | AF labels generated by sequence panel; AF diagnostics launcher exists | ADwin AF command service, demag state tracking, coil safety interlocks |
| IRM / ARM | Not yet integrated | Settings + stub dialog entries exist | Full workflow/state machine and protocol wrappers |
| XY changer / Z / turn motors | Not yet integrated | Standalone apps remain separate; shell has launch stubs | Migrated queue orchestration and exclusive ownership in rapid_main services |
| Vacuum | Not yet integrated in `rapid_main` | Stub/launcher dialog only | Vacuum state polling and fault propagation into flow control |
| Calibration tools | Not yet implemented | Calibration tab exists for parameter storage | Automated + guided + manual calibration center with audit trail |
| Gaussmeter / meter checks | Not yet integrated | Standalone apps exist; shell launch path exists | Meter read abstraction and meter-driven calibration/verification workflows |
| Data review/plots | Partially complete | Measurement plot placeholder now opens the built-in `PlotsDialog` and a launcher for the `data_viewer` utility is available from the main View menu | Advanced plotting and statistical checks (CSD, Zijderveld, quicklooks) with background refresh and queue-linked artifact filtering |
| Sample IO & queue | Partial | Sample table + queue compiler module + bundle writer tests | Queue-state recovery is now implemented (interrupted `Running` rows auto-reset to `Pending`) and sample metadata parsing exists; remaining work is richer queue reuse of `.sam` metadata in queue assembly and review/report modules. |
| Persistence/config | Partially complete | JSON config model + load/save + settings-to-runtime binding | Migration from legacy INI and settings diff checks |
| Safety + recovery | In progress | Pause/halt hooks exist; ownership enforced for measurement + diagnostic dialogs; preflight and per-step timeout/cancel guards active in worker | Expand to all device services, emergency-stop state, restart-safe recovery, and explicit physical-fault provenance |
| Packaging | Not started in this phase | Standalone app packaging exists for other tools | `rapid_main` packaging and smoke install test |

### Current phase execution order (short-form)

- **Phase 1 (Parity inventory):** Complete remaining audit pass and mark each active VB6 form/function as mapped or approved obsolete.
- **Phase 2 (Contracts + service layer):** Keep one stable `MeasurementBackend` contract and expand into real backend classes for SQUID, changer, AF, IRM/ARM, vacuum, and gaussmeter; remove direct widget-to-device coupling. Device ownership for measurement and diagnostic dialogs is in place, with timeout-aware preflight/step operations wired in worker path.
- **Phase 3 (Responsive shell):** Finalize adaptive layouts and panel resizing behavior for 1280×720+ and scaling 100–200%, plus persistent layout presets.
- **Phase 4 (Core AF workflow):** Deliver end-to-end AF workflow using real backends: preflight, sample selection, orientation/positioning, treatment, measurement, write outputs atomically, advance queue, and return to safe state.
- **Phase 5 (Treatment completion):** Add IRM/ARM, susceptibility, thermal, VRM, and remaining specialized routines with VB6 parity evidence.
- **Phase 6 (Calibration center):** Implement shared Calibration Center with automated and guided/manual modes. Default to automated when digital control is available; keep manual overrides explicit and logged.
- **Phase 7–10:** data review hardening, transition tooling, fault-injection tests, and production packaging/approval.

### Calibration center requirements (reframed per your instruction)

The calibration workspace should treat automation and manual operation as first-class:

1. **Automated mode by default** where instrument+meter endpoints are controllable digitally.
2. **Guided mode** for physical inspection steps the software cannot guarantee.
3. **Manual mode** for service and anomaly recovery with controlled direct commands and full audit logging.
4. One shared workflow engine so automated and manual modes produce the same output schema and report format.
5. All fits/results remain proposal-only until operator approval, and no calibration can silently change live coefficients.

### This plan as a concrete "goal card"

#### Goal

Deliver a validated `rapid_main` AF-first pilot within the next two sprints:
- One supervised AF specimen run executes completely in rapid_main in both simulation and bench/hardware modes.
- The UI remains usable at 1280×720 and selected higher sizes without clipping or control crowding.
- Operators can still perform a manual calibration path for any automated routine.

#### Milestones

1. **Sprint 1:** close remaining Phase-1 gaps
   - Finish traceability matrix and tag each VB6 behavior as mapped, replaced, or retired.
   - Add golden fixtures (SAM, INI, queues, output expectations) with versioned references.
2. **Sprint 2:** real-backend phase for AF + SQUID + motion preflight
   - Implement AF/SQUID backend classes and queue-safe ownership.
   - Connect preflight checks to "Start Run" and block run when unsafe.
3. **Sprint 3:** AF workflow hardening
   - Add persistent specimen/run state machine transitions (Idle/Preflight/Treating/Measuring/Saving/Returning).
   - Implement stop-safe atomic writes and recovery journal entries.
   - Add failure-path tests (timeout, disconnect, halt).
4. **Sprint 4:** Calibration center skeleton
   - Add procedure registry and templates.
   - Implement one automated routine end-to-end (coil zero/gain or meter baseline check) plus manual equivalent mode.
5. **Sprint 5:** transition-to-production prep
   - Add operator transition map, "where did VB6 go" lookup, and quick-start run sheets.
   - Stabilize packaging/build and produce a release-candidate install check.

### Tracking and maintenance

- Keep this file as the canonical execution ledger. Update each milestone with pass/fail evidence (tests, fixtures, and signed operator check-off).
- Record any deviations from roadmap as explicit phase re-planning notes rather than silently dropping features.
- Tie every code change to a roadmap line item to keep feature creep and regression risk visible.

## Concrete patch plan for current run (13 July 2026)

### Current phase progression

- **Overall phase in progress:** Phase 2 (Contracts + service layer) into Phase 3 (Responsive shell).
- **Current objective:** remove major UI/comms usability blockers in sequence and diagnostics, finish DC motor/vacuum/IRM-ARM integration quality, and complete migration-tracker evidence updates.

### Concrete patch sequence (clean, separable commits)

1. `feat/rapid-main-sequence-layout`
   - Files: `RapidPy/rapid_main/rapid_main/panels/sequence.py`, `RapidPy/rapid_main/rapid_main/app.py`
   - Scope:
     - Reorganize presets + measurement cards so values remain visible without scrollbars.
     - Fix truncated toggle-box values (`AF / IRM`, log-factor rows, etc.).
     - Expand and lock sidebar width to fit icon+label text consistently.
   - Validation: layout smoke at 1280×720 and 1920×1080.

2. `feat/rapid-main-dc-motors-hardware-path`
   - Files: `RapidPy/rapid_main/rapid_main/dialogs/dc_motors.py`, `RapidPy/rapid_main/rapid_main/device_ownership.py`, `RapidPy/rapid_main/rapid_main/diagnostic_services.py`
   - Scope: eliminate remaining placeholder behavior in motor diagnostics and ensure ownership-safe, explicit status/error propagation.
   - Validation: no regression in diagnostic smoke workflow and backend tests.

3. `ui/icons-and-application-identity`
   - Files: main app and other app entry points using top-level windows
   - Scope:
     - Ensure main icon is distinct from COM-port app.
     - Apply icon to both window frame and taskbar/titlebar for every shipped executable window.
   - Validation: launch each app and verify non-default icon presence.

4. `docs/parity-evidence-hardening`
   - File: `docs/vb6_parity_inventory.md`
   - Scope:
     - Update statuses/evidence for `frmVacuum`, SQUID, IRM/ARM, DC motor/changer pathways, and automated queue/sample handling behaviors.
     - Record acceptance criteria or approved-retirement notes where still pending.

5. `feat/rapid-main-workflow-hardening`
   - Files: `RapidPy/rapid_main/rapid_main/queue_compiler.py`, `RapidPy/rapid_main/rapid_main/app.py`, `RapidPy/rapid_main/rapid_main/measurement_worker.py`, `RapidPy/rapid_main/tests/*`
   - Scope:
     - Finalize VB6-to-rapid_main automated sample-handling parity checks (compile, preflight, queue halt/recovery, resume behavior).
     - Add/update regression tests for compile failures, recovery transitions, and halt semantics.
   - Validation: `python -m unittest discover -s RapidPy/rapid_main/tests -p 'test*.py' -v`

### Acceptance gates for current run

- No `QPushButton.setWordWrap` runtime errors.
- No main-sequence value clipping on baseline window widths.
- No side-menu icon/text clipping.
- Main app presents a custom icon (not default Python icon) and all modules use explicit window/icon identity.
- VB6 parity tracker updated with objective evidence references and no silent exclusions.

## Continuation checkpoint — 3 October 2026

The main glass workspace, susceptibility/holder lifecycle, transactional
bundles, calibration governance, and helper integration are retained. New work
adds a wheel-backed portable build, frozen helper dispatch, startup diagnostics,
and an isolated offline UI smoke check. Treatment validation blocks missing,
simulated, or disconnected actuators and invalid requests before execution.
IRM/ARM requests cannot silently clamp or partially apply invalid ramps.
AFZ/AFMAX numeric labels retain their requested field. All 599 tests pass with
isolated Qt preferences; the source UI/helper smoke check passes.

The July snapshot above is historical. See
`docs/rapid-main-portable-release.md` for the current launch/test procedure.
Full replacement remains open: AF multi-pass choreography/field calibration,
independent ARM bias and pulse-IRM circuitry, backfield polarity, synchronized
RRM, MCC/DAC binding, and transverse motor execution are software integrations
still needing closure. Physical/operator acceptance is additional; do not
classify all remaining work as hardware-only testing.

### Calibrated AF source checkpoint — 3 October 2026

Live queue AF now uses the full source-backed axial/transverse/transverse
lifecycle, AFZ and independent-coil AFMAX, accepted calibration imports,
verified specimen centering, cooperative native ADwin cancellation, independent
cleanup attempts and immutable indexed treatment evidence. The main-app suite
passes 613 tests. See docs/af-treatment-integration.md for evidence and limits.
The portable pilot predates these source changes. Calibration tooling and bench
acceptance remain open; independent ARM bias, pulse IRM, backfield, RRM and
MCC/DAC integration still require development. The full-system goal remains active.

### Independent ARM bias source checkpoint — 3 October 2026

Queue ARM now uses one calibrated axial AF pass and independent MCC DC bias,
with corrected legacy calibration units/bindings, active-low gate sequencing,
cooperative setup cancellation, safe output cleanup and indexed phase evidence.
The MCC Universal Library ctypes adapter probes read-only and checks each native
conversion/output status. All 622 main-app tests pass. Settings show imported
AF/ARM calibration summaries. See docs/arm-bias-integration.md. Manual ARM control,
calibration tooling, pulse IRM/backfield, RRM and bench qualification remain open.
The full-system goal is active; the portable pilot has not yet been rebuilt.

### Manual ARM and pulse IRM planning checkpoint — 3 October 2026

Manual ARM now reuses the complete calibrated lifecycle on a GUI-safe worker,
retains device leases until cancellation cleanup finishes, requires specimen
identity, and publishes immutable treatment records with SHA-256 indexes. Success
and failure evidence, unwritable destinations, publication failure and close
cancellation have injected coverage. Queue context is restored after each attempt.

Pulse IRM now has separate persisted configuration/imports and tested Old/ASC
capacitor-field planning, ASC boost, module/backfield gating and calibrated limits.
The common MCC driver has checked ADC read/conversion support. Pulse execution,
Matsusada scaling, channel binding, polarity/positioning, pulse evidence and manual
IRM integration remain open. See docs/pulse-irm-development.md. The main suite
passes 633 tests; the additional ADC check and UI smoke pass. The goal remains
active and the portable pilot still predates these source changes.

### Capacitor pulse execution and integration — 3 October 2026

Live queue and manual IRM now use calibrated capacitor charge/readback/trim/fire/
discharge with ADwin coil/polarity readback, source Matsusada scaling, enabled
coil-temperature interlocks, two residual pulses at load and verified specimen
positioning. Backfield changes the polarity relay and retains non-negative DAC
voltages. The old generic AF waveform IRM implementation is removed/inhibited.

Unverified pulse safe state inhibits motor return, further treatments and device
ownership until a recorded discharge-only recovery succeeds. Immutable pulse,
motion and recovery records are retained/indexed for worker and manual runs.
Legacy capacitor-voltage/field mapping and axis normalization are corrected.
All 651 main-app tests pass in 88.038 seconds; six-panel/six-helper UI smoke passes.
See docs/pulse-irm-development.md. Zero-field queue semantics, RRM, motor station
configuration, calibration tooling, probe motion, release rebuilding and physical/
scientific acceptance remain open. The full-system goal remains active.

### Native motor station and RRM lifecycle — 3 October 2026

Imported native motor wiring now routes four address-16 controllers on separate
ports rather than assuming addresses 1..4 on COM3. Both queue and manual motor
diagnostics use native travel, speed, torque and 57600 baud calibration.
Fractional belt steps and negative specimen-centre floor semantics are retained.
Slot number is no longer misinterpreted as lift speed. Incomplete station profiles
block preflight without constructing a motor transport.

RRM/RRMZ now have explicit field/signed-speed labels, optional independent ARM
bias, calibrated AF while spinning, cooperative rotation readback checks,
verified stop and independent field cleanup. Unsafe rotation/field inhibits lift
return and further treatments until recorded stop recovery succeeds. Worker and
manual recovery publish immutable RRM evidence with SHA-256 indexes. See
docs/motor-station-and-rrm-development.md for requirements and remaining work.

Verification: 673 main-app tests passed in 84.794 seconds; the additional long-spin
native reference regression and current motor/RRM subset pass (23 tests).
Six-panel/six-helper source UI smoke passes. The full-system goal remains active;
the portable pilot has not yet been rebuilt from this newer source.

### Zero-field IRM, typed field units and routine controls — 3 October 2026

Queue IRM0 now performs two source residual discharges at load and positions the
specimen without charging or applying a specimen pulse. Zero-treatment records
retain residual count, explicit charged-pulse status, bounded discharge and
required immutable artifact digests. Recovery remains discharge-only.

Typed G/MT labels convert once at the actuator boundary, support explicit pulse
axes and keep ARM peak/bias units independent. The sequence panel now uses RRM
field/speed/coil/bias and ARM AF-peak controls; logarithmic IRM and numeric
backfield steps preserve gauss inputs. The documented Hawaiian series is explicit
25–800 G rather than 25–800 mT. Compact panels stack controls above the preview
and scroll vertically. See docs/zero-irm-and-treatment-units.md for contracts,
source discrepancies and remaining work. The full-system goal remains active.

Verification: 696 isolated main-app tests pass in 90.593 seconds. Source UI smoke
passes six panels and six helpers. Temperature is rechecked before each residual
fire and immediately before charging after positioning. The full-system goal
remains active; persistence of unsafe hardware state across restarts is a next
required development task, and the portable pilot remains older than this source.

### Durable field-treatment safety across restarts — 3 October 2026

Live AF, ARM, pulse IRM and RRM persist unfinished operations before treatment
motion/output and retain station identity, plan and physical recovery evidence.
Atomic checksummed storage, stale-token checks and an OS lifetime ownership
lease prevent restart amnesia and concurrent treatment/recovery. Changed wiring,
corrupt state and failed publication block further operation. Recovery does not
reapply fields, charge/fire pulses or start rotation. Pulse discharge precedes
verified reference/lift return; field recovery stops ADwin processes and checks
relay clear without rebooting the board. See docs/hardware-safety-persistence.md.

The full-system goal remains active. Standalone helper ownership, non-treatment
motion persistence, calibration/probe tooling, portable rebuild and physical/
scientific acceptance remain open. Development changes are being committed and
pushed in verified milestones at the user's request.

Verification: 724 isolated main-app tests pass in 92.420 seconds, including 28
durable-safety tests. Six-panel/six-helper source UI smoke passes. The portable
pilot still predates this source checkpoint and must be rebuilt before release.
