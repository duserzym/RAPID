# VB6 Capability Parity Inventory (RapidPy Migration Tracker)

## Status

This matrix is the executable evidence artifact for roadmap Phase 1.  
All capabilities that currently drive routine scientific workflows in `VB6/Paleomag v3.vbp` are mapped to an evidence-backed target in `RapidPy/rapid_main` or marked for explicit retirement.

Legend for `Migration Status`:

- **Not assessed** — inventory exists in VB6 but no matching implementation has been inventoried yet.
- **Mapped** — destination exists in `rapid_main` but not yet fully production-ready.
- **Stand-alone implemented** — implemented in a separate standalone app; not yet integrated into `rapid_main`.
- **Integrated** — implemented in `rapid_main` with a real-or-simulated service path.
- **Not required** — no longer relevant to current active VB6 workflow.

## Current Baseline Mapping

| VB6 Source / Function | Mapping in RapidPy | Migration Status | Evidence (`rapid_main`) | Notes |
|---|---|---|---|---|
| `frmMagnetometerControl` + `modFlow` + `modProg` | Main app control model in `rapid_main.app`, `rapid_main.panels.dashboard`, runtime controls in `MainWindow`, `RuntimeEstimator` | Mapped | `RapidPy/rapid_main/rapid_main/app.py`, `panels/dashboard.py` | Session flow/state machine is partially implemented (run timer + labels), but not full state-machine semantics yet. |
| `frmProgram` + Rockmag routine logic | `rapid_main.panels.sequence` | Mapped | `RapidPy/rapid_main/rapid_main/panels/sequence.py` | Generates labels, save/load, and validation preview. Needs full executable mapping to queue/state logic. |
| `frmMeasure` / `modMeasure` / `frmStats` | `rapid_main.panels.measurement` + `rapid_main.measurement_worker` | Mapped | `RapidPy/rapid_main/rapid_main/panels/measurement.py`, `measurement_worker.py` | Worker supports stop/pause/halt and file writing; missing full coordinate pipeline and complete state transitions. |
| `frmSquid` + magnetometer helpers | `rapid_main.hardware_contracts` (+ future backend adapters) | Mapped | `RapidPy/rapid_main/rapid_main/hardware_contracts.py`, `dialogs/squid_comm.py` | Current path is simulated (`NoCommBackend`) plus dialog scaffolding. |
| `frmChanger*` set | Not in `rapid_main` shell, dedicated standalone app | Stand-alone implemented | `RapidPy/changer_xy_control` | Functionality present and usable but still outside unified app in this phase. |
| `frmDCMotors` + `modMotor` + XY changers | Mapped | Mapped | `RapidPy/dc_motor_control`, `RapidPy/updown_control`, `RapidPy/changer_xy_control`, `RapidPy/rapid_main` | Motor controls are now available in `rapid_main` via `DCMotorDialog` + ownership-safe diagnostics path. |
| `frmVacuum` | Stand-alone implemented | Mapped | `RapidPy/updown_control` (vacuum logic), `RapidPy/rapid_main/rapid_main/dialogs/vacuum.py`, `diagnostic_services.py` | Vacuum is now owned/routed through `rapid_main` dialog and has no-comm + transport adapter behavior; production hardening remains for transport/failure recovery. |
| AF treatment (`frmAF`, ADwin AF controls, AF tuners) | Mapped | Mapped | `RapidPy/af_tuner`, `RapidPy/af_clip_test`, `RapidPy/updown_control`, `_launch_af` in `app.py` | AF workflow has shell launch path and demo workflow; transport contract remains pending for fully in-app end-to-end execution. |
| IRM/ARM (`frmIRMARM`, voltage calibration forms) | Mapped | Mapped | `RapidPy/rapid_main/rapid_main/dialogs/irm_arm.py`, `dialogs` + `diagnostic_services.py` | Backend adapter + ownership path is available; remaining work is production transport and voltage-calibration workflow closure in main shell. |
| `frm908AGaussmeter` | Launcher path / existing standalone | Stand-alone implemented | `RapidPy/gaussmeter_control`, `rapid_main/dialogs/menus` | Needs meter-aware calibration/verification inside rapid_main contracts. |
| Thermal (`modThermal` + associated forms) | Not yet integrated | Not assessed | (no robust mapping yet) | Requires migration if thermal routine remains active in scope. |
| Susceptibility (`Susceptibility` forms) | Settings + measurement read path | Mapped | `panels/measurement.py`, `measurement_worker.py`, susceptibility read optional path | Backend-only optional reading implemented; user workflow not yet parity complete. |
| `frmVRM` (VRM routines) | Not yet integrated in `rapid_main` | Not assessed | (no direct mapping) | VRM app exists separately: `RapidPy/vrm_logger`. |
| Calibration rod / AF/IRM tuning forms | Parameter tabs + launchers | Mapped | `RapidPy/rapid_main/rapid_main/panels/settings_panel.py`, `RapidPy/rapid_main/rapid_main/panels/calibration.py`, `RapidPy/rapid_main/rapid_main/calibration.py` | Calibration workflows now route through the rapid_main Calibration Center (automated/manual modes). |
| Plots/stats (`frmPlots*`, Zijderveld, stereonet, quicklooks) | Placeholder + standalone preview | Stand-alone implemented / Not assessed | `RapidPy/data_viewer`, `rapid_main.dialogs.plots` | Need on-demand integration into main workflow path. |
| Data saving/export (`frmFileSave`, directories) | `rapid_main.io.measurement_bundle` + output writers | Integrated (core path) | `RapidPy/rapid_main/rapid_main/io/measurement_bundle.py`, `measurement_worker.py` | `.sample`, `.rmg`, `measurements.txt`, `specimens.txt` writes work in tandem. |
| Webcam (`frmWebcam`) | Launcher + own app | Stand-alone implemented | `RapidPy/webcam_viewer`, `RapidPy/rapid_main/rapid_main/dialogs/webcam_dialog.py` | Needs docked/optional integration path in main shell. |
| Settings/INI editors (`frmSettings*`, `frmOptions`, channel config) | Structured JSON settings + settings panel | Mapped | `RapidPy/rapid_main/rapid_main/config.py`, `panels/settings_panel.py`, `RapidPy/rapid_main/rapid_main/legacy_ini.py` + `panels/settings_panel.py` (Import VB6 INI…) | VB6 INI import now routes into app settings with mapping report and warnings for unmapped keys. |
| diagnostics (`frmDebug`, DAQ/ADwin comm debug, step monitor, messages) | Dialog stubs + separate standalone apps | Mapped | `RapidPy/rapid_main/rapid_main/dialogs` | Core diagnostic logic still simulated/stubbed. |

## Full `VB6/Paleomag v3.vbp` Component Sweep (2026-07-10)

The table below captures every `Form`, `Module`, and `Class` entry from
`VB6/Paleomag v3.vbp` and current migration status to support complete
Phase-1 inventory closure.

| VB6 item | Type | Migration status | RapidPy location / notes |
|---|---|---|---|
| `frmMagnetometerControl` | Form | Mapped | `rapid_main` shell + dashboard + flow scaffolding |
| `frmAbout` | Form | Mapped | `rapid_main` dialogs/help roadmap placeholder |
| `frmLogin` | Form | Mapped | `rapid_main` main menu + `dialogs/login.py` launcher; session handling remains a refinement target |
| `frmSplash` | Form | Mapped | `rapid_main` app startup + loading states (placeholder: startup guidance screen not yet implemented) |
| `frmTip` | Form | Mapped | `rapid_main` help/status messaging and planned in-app guidance artifacts |
| `modProg` | Module | Mapped | flow/runtime scaffolding in shell |
| `modMeasure` | Module | Mapped | `measurement_worker` + `panels.measurement` |
| `modVector3d` | Module | Mapped | `RapidPy/rapid_main/rapid_main/geometry.py`; `measurement_worker.py`; `io/specimen_reader.py` |
| `modMotor` | Module | Stand-alone implemented | `updown_control` / `dc_motor_control` launcher path |
| `frmChangerSampOrder` | Form | Mapped | `rapid_main/rapid_main/queue_compiler.py`, sample-queue panel UI | Duplicate position and invalid sample-step validation are now tracked in the compiler contract (`QueueValidationResult`) with regression tests. |
| `modChanger` | Module | Stand-alone implemented | `changer_xy_control`/`updown_control` |
| `frmMeasure` | Form | Mapped | `panels.measurement` |
| `frmStats` | Form | Mapped | `panels.measurement` placeholders pending full stats parity |
| `modPrint` | Module | Mapped | `rapid_main/io/measurement_bundle.py`, `measurement_worker.py`, `measurement output writers` |
| `frmVacuum` | Form | Mapped | `updown_control` launcher |
| `frmDCMotors` | Form | Mapped | `rapid_main` dialog launcher (`dialogs/dc_motors.py`) + `updown_control` / `dc_motor_control` reference implementation. |
| `frmAF_2G` | Form | Stand-alone implemented | AF calibration/treatment apps exist; main path not integrated |
| `frmSendMail` | Form | Not assessed | Emailing/report dispatch mapping pending |
| `frmSquid` | Form | Stand-alone implemented | `gaussmeter`/meter abstraction pending in shell |
| `frmOptions` | Form | Mapped | `panels.settings_panel` |
| `modFlow` | Module | Mapped | state-machine draft and timer scaffolding in `app.py` |
| `frmStepMonitor` | Form | Mapped | dialog launcher exists (`dialogs.step_monitor`) |
| `modDataAnalysis` | Module | Not assessed | Scientific calculations still incomplete in workflow |
| `modMagnetometer` | Module | Not assessed | Requires SQUID backend parity |
| `modAF_2G` | Module | Stand-alone implemented | AF treatment app remains external |
| `frmIRMARM` | Form | Stand-alone implemented | `dialogs.irm_arm` placeholder |
| `frmRockmagRoutine` | Form | Not assessed | Rock-magnetic routine migration pending |
| `RockmagStep` | Class | Not assessed | Domain model mapping pending |
| `RockmagSteps` | Class | Not assessed | Domain model mapping pending |
| `frmSusceptibilityMeter` | Form | Mapped | susceptibility read path in measurement worker/settings |
| `SampleCommand` | Class | Mapped | queue/sequencer pipeline uses newer equivalents |
| `SampleCommands` | Class | Mapped | queue/sequencer pipeline uses newer equivalents |
| `SampleIndexRegistration` | Class | Mapped | `rapid_main/data_model.py`; `rapid_main/io/sample_index.py` (`read_sample_index_registrations`) |
| `SampleIndexRegistrations` | Class | Mapped | `rapid_main/data_model.py`; `rapid_main/io/sample_index.py` (`read_sample_index_registrations`, `SampleIndexRegistrations`) |
| `frmSampleIndexRegistry` | Form | Mapped | Modern equivalent is `SampleSelectDialog` loading workflow in `rapid_main/dialogs/sample_select.py`, with registry metadata preserved via `SampleIndexRegistration`; full add/edit registry operations are deferred to a later milestone with parity evidence tracking. |
| `frmProgram` | Form | Mapped | `panels.sequence` |
| `Samples` | Class | Mapped | specimen model in `data_model.py` |
| `Sample` | Class | Mapped | specimen model in `data_model.py` |
| `Cartesian3D` | Class | Mapped | `RapidPy/rapid_main/rapid_main/geometry.py`; `measurement_worker.py`; `io/specimen_reader.py` |
| `frmChanger` | Form | Stand-alone implemented | `changer_xy_control` |
| `frmRerunSamples` | Form | Mapped | queue sample rerun/skip control and persistence now live in `panels.sample_queue` with failure-state resets (`Error`→`Pending`) and `QSettings` row restoration. |
| `frmSampleSelect` | Form | Mapped | sample queue/panel behavior |
| `modSusceptibility` | Module | Mapped | susceptibility read path |
| `Angular3D` | Class | Mapped | `RapidPy/rapid_main/rapid_main/geometry.py` |
| `MeasurementBlock` | Class | Not assessed | sequence block model migration pending |
| `MeasurementBlocks` | Class | Not assessed | sequence block model migration pending |
| `modStatusCode` | Module | Not assessed | status/error code taxonomy harmonization pending |
| `frmSampleQueueMonitor` | Form | Mapped | queue monitor panel scaffolding |
| `frmVRM` | Form | Not assessed | not yet integrated in `rapid_main` shell |
| `frmPlots` | Form | Stand-alone implemented | data review UI in `data_viewer` |
| `frmWebcam` | Form | Stand-alone implemented | `dialogs.webcam` launch path |
| `modConfig` | Module | Mapped | `config.py` migration started |
| `CIniFile` | Class | Mapped | `RapidPy/rapid_main/rapid_main/legacy_ini.py` + Settings import workflow (`Import VB6 INI…`) with mapping notes and validation. |
| `frmAFTuner` | Form | Stand-alone implemented | `RapidPy/af_tuner` |
| `frmCalibrateCoils` | Form | Mapped | `rapid_main/panels/calibration.py`, `rapid_main/calibration.py`, `rapid_main/app.py` stack/navigation wiring |
| `frmDAC_Comm` | Form | Not assessed | DAC setup parity pending |
| `frmFileSave` | Form | Mapped | output writers + measurement bundle |
| `ADWIN` | Module | Stand-alone implemented | ADwin apps + planned real backend |
| `modMCC` | Module | Stand-alone implemented | DAQ abstraction migration pending |
| `mod908AGaussmeter` | Module | Stand-alone implemented | meter control in external app |
| `frm908AGaussmeter` | Form | Stand-alone implemented | launcher + standalone exists |
| `modFileSave` | Module | Mapped | `io` writers (RMG/samples/MagIC) |
| `Board` | Class | Stand-alone implemented | MCC/DAQ abstraction not integrated |
| `Boards` | Class | Stand-alone implemented | MCC/DAQ abstraction not integrated |
| `Channel` | Class | Stand-alone implemented | MCC/DAQ abstraction not integrated |
| `Channels` | Class | Stand-alone implemented | MCC/DAQ abstraction not integrated |
| `Range` | Class | Stand-alone implemented | MCC/DAQ abstraction not integrated |
| `Wave` | Class | Stand-alone implemented | DAC waveform config not integrated |
| `Waves` | Class | Stand-alone implemented | DAC waveform config not integrated |
| `ChannelDescs` | Class | Stand-alone implemented | DAQ mapping not integrated |
| `frmSettings_new` | Form | Mapped | `panels.settings_panel` |
| `frmADWIN_AF` | Form | Stand-alone implemented | AF runtime in external ADwin apps |
| `frmDialog` | Form | Mapped | generic dialog patterns in `rapid_main.dialogs` |
| `frmDebug` | Form | Mapped | `dialogs` placeholder |
| `frmIRM_VoltageCalibration` | Form | Not assessed | IRM calibration procedure mapping pending |
| `frmShutdownMsg` | Form | Mapped | `rapid_main/rapid_main/app.py` shutdown action path (`_confirm_shutdown`, `_request_shutdown`, menu/toolbar exit wiring) and `DeviceOwnershipManager`-safe halt flow on exit. |
| `frmCalRod` | Form | Stand-alone implemented | coil/rod calibration parity pending |
| `modAF_DAQ` | Module | Stand-alone implemented | AF DAQ bridge in external project |
| `frmINIConverter` | Form | Mapped | VB6 INI converter behavior is replaced by the `Import VB6 INI…` action in `SettingsPanel`, with mapping/reporting and preserved operator review of unresolved keys. |
| `modThermal` | Module | Not assessed | thermal workflow migration pending |
| `frmXYHoming` | Form | Stand-alone implemented | changer homing logic in standalone app |
| `IRMData` | Class | Stand-alone implemented | IRM model migration pending |
| `IrmDataPoint` | Class | Stand-alone implemented | IRM model migration pending |
| `InterpolationRange` | Class | Not assessed | fitting utility parity pending |
| `InterpolationRanges` | Class | Not assessed | fitting utility parity pending |
| `XYCup` | Class | Stand-alone implemented | XY changer model external |
| `XYCup_Positions` | Class | Stand-alone implemented | XY changer model external |
| `frmTransverseProbeAutoPosition` | Form | Not assessed | transverse probe auto-position parity pending |
| `AngleVsField_Point` | Class | Not assessed | transverse automation math pending |
| `AngleVsFieldCollection` | Class | Not assessed | transverse automation math pending |
| `ProbeAngleOptimizer` | Class | Not assessed | transverse optimization workflow pending |
| `AdwinAfInputParameters` | Class | Stand-alone implemented | AF parameter model in external AF stack |
| `AdwinAfOutputParameters` | Class | Stand-alone implemented | AF parameter model in external AF stack |
| `AdwinAfParameter` | Class | Stand-alone implemented | AF parameter model in external AF stack |
| `AdwinAfPauseConstants` | Class | Stand-alone implemented | AF pause constants in external AF stack |
| `AdwinAfRampStatus` | Class | Stand-alone implemented | AF ramp-status model external |
| `modLogAFParameters` | Module | Stand-alone implemented | AF log parser logic external |
| `modListenAndLog` | Module | Not assessed | communication logger abstraction pending |

## Immediate action items for objective completion

1. Complete remaining "Not assessed" VB6 behaviors with owner + acceptance criteria.
2. Add one row per active VB6 workflow that has evidence of:
   - preflight checks
   - failure/timeout behavior
   - atomic output writes and recovery checkpoints
   - evidence of AF parity in `rapid_main`.
3. Link every new row to explicit tests/fixtures in `RapidPy/rapid_main/tests` as they are added.

## RapidPy current gap snapshot (2026-07-13)

### Hardware control readiness

| VB6 subsystem | RapidPy status | What is complete | Remaining to close parity |
|---|---|---|---|
| SQUID comm + readings | Stand-alone placeholder / contract scaffold | Diagnostic dialog exists, no-comm backend contract and ownership-safe launch path | Real transport adapter in `rapid_main` shell + read-quality flags + timeout/retry + safe abort behavior for active runs. |
| Motors (changer/XY/up-down/turning) | In-process diagnostic dialog + dual-path backend | `DCMotorDialog` + `DCMotorNoCommBackend` and adapter path wired in `diagnostic_services.py`; ownership lock integrated | Persisted diagnostics settings, richer safety interlocks for sample transfer operations, and bench validation against hardware command set. |
| Vacuum control | Placeholder + contract scaffold | Dialog exists and backend contract responds in no-comm/sim mode | Hardware transport layer for live vacuum readback/pump controls + pressure fault propagation into queue/run halt states. |
| IRM / ARM + calibrations | Contract-backed no-comm path | `irm_arm.py` dialog executes simulated IRM/ARM commands and supports field reset path | Real IRM/ARM calibration + protocol-specific voltage calibration + queue-safe run context. |
| AF treatment/tuning | Demo/launcher path with queue-safe presets | AF demo labels and diagnostics launcher exist; preflight hooks improved | Production AF/DAQ adapter in-shell (currently still external AF tooling dominates this workflow). |
| Webcam/DAQ/Susceptibility/VRM | Launch-and-review scaffolds | Stand-alone entry points remain available | Unified workflow entry from main app and parity-mapped output capture; some components still pending integration. |

### Automated sample handling readiness

| VB6 behavior | RapidPy status | Remaining parity work |
|---|---|---|
| `frmSampleQueueMonitor` + run control loop | Integrated and simulator-driven | Add hardware-safe pause/halt semantics across all queue-run phases and persistable recovery after operator stop/exception. |
| `frmChangerSampOrder` validation | Queue compiler validation in place (`compile_queue`, strict mode, tests) | Expand duplicate/invalid-position diagnostics in operator-facing diagnostics, include hard stop if mandatory prerequisites missing. |
| `frmRerunSamples` | Integrated | Improve restart semantics for partially failed automation and provide explicit recovery action for unsafe machine states. |
| `MeasurementBlock` / step mapping | Mapped in compiler output path | Continue adding domain model parity for rock-magnetic block metadata and VB6-equivalent report annotations in exported bundles. |

#### Current blockers to close before main-shell parity claim

- No production transport adapters for several core instruments in `diagnostic_services.py` (SQUID, vacuum, IRM/ARM, changer motors).
- No full end-to-end proof that queue-run execution can resume from persistent interrupted states.
- No explicit operator-facing evidence artifact for every queue transition (`preflight -> loading -> treating -> measuring -> validating -> saving -> returning -> complete/error`).

### Not Assessed Remediation Queue (owner + acceptance criteria)

The list below tracks the remaining "Not assessed" capabilities and gives explicit ownership plus acceptance gates.

| VB6 item | Owner | Acceptance criteria |
|---|---|---|
| IRM/ARM (`frmIRMARM`, voltage calibration forms) | IRM/ARM Workflow Owner | Stand-alone controls become `rapid_main`-routed workflows with contract-backed run-state, ownership, and evidence that preflight/fail-safe checks run before treatment and calibration steps. |
| Thermal (`modThermal` + associated forms) | Thermal Migration Owner | Thermal setup/calibration/edit forms are either mapped into an integrated `rapid_main` service or explicitly approved as out-of-scope with rationale in this tracker. |
| `frmVRM` (VRM routines) | VRM/Logger Integration Owner | A `rapid_main` pathway exists that runs VRM acquisition from the shell with deterministic queue/run metadata and output compatibility with `Rapid` VRM artifacts. |
| Plots/stats (`frmPlots*`, Zijderveld, stereonet, quicklooks) | Data Review Owner | `rapid_main` exposes on-demand review panels for all phase-appropriate outputs (moment vectors, Zijderveld, quicklooks, PCA), backed by reusable analysis services and stable tests. |
| `modVector3d` | Math Utilities Owner | Completed | Orientation/matrix math is ported to `RapidPy/rapid_main/rapid_main/geometry.py` with tests in `RapidPy/rapid_main/tests/test_geometry.py`; used in `measurement_worker` and specimen parsing. |
| `frmSendMail` | Reporting/Run-log Owner | Operator reporting is documented with one modern equivalent or marked Not required with traceable justification. |
| `modDataAnalysis` | Data Quality Owner | Statistical/sanity checks used by measurement review are implemented in shared analysis modules with unit tests and clear output schema links. |
| `modMagnetometer` | SQUID/Measurement Owner | Full SQUID contract implementation includes preflight, safe stop/retry behavior, and measurement read quality flags. |
| `frmRockmagRoutine` | Rock-Magnetic Owner | Rock-magnetic routine behavior is mapped to sequence/measurement flow in `rapid_main` with at least one regression fixture. |
| `RockmagStep` | Rock-Magnetic Owner | Rock-magnetic step structures are represented in new or existing `rapid_main` domain models and serialized into compatible outputs. |
| `RockmagSteps` | Rock-Magnetic Owner | Rock-magnetic batch-step container structure is represented in `rapid_main` queue/compiler workflows. |
| `SampleIndexRegistration` | Sample-Index Owner | Completed | `rapid_main/data_model.py` defines `SampleIndexRegistration`; `rapid_main/io/sample_index.py` maps `.sam` registry entries with preserved specimen order metadata. |
| `SampleIndexRegistrations` | Sample-Index Owner | Completed | `rapid_main/data_model.py` defines `SampleIndexRegistrations`; `rapid_main/io/sample_index.py` and `rapid_main/tests/test_sample_index.py` verify order preservation and header handling. |
| `frmSampleIndexRegistry` | Sample-Index Owner | Completed | `.sam` registry loading flow is represented in `SampleSelectDialog` (`rapid_main/dialogs/sample_select.py`), with registry metadata used where present and fixture-backed parsing tests. |
| `Cartesian3D` | Math Utilities Owner | Completed | Vector math primitives and conversion helpers are covered in `RapidPy/rapid_main/geometry.py` and `RapidPy/rapid_main/tests/test_geometry.py`. |
| `frmRerunSamples` | Queue/Workflow Owner | Completed | queue rerun/skip controls are implemented in `SampleQueuePanel` (failure resets, queue-persistent status restore, and skip-aware run extraction) with regression coverage in `RapidPy/rapid_main/tests/test_sample_queue_helpers.py` (`_normalize_status`) and queue execution tests via `test_queue_and_bundle` coverage patterns. |
| `Angular3D` | Math Utilities Owner | Completed | Angular direction conversion helpers are covered in `RapidPy/rapid_main/geometry.py` and validated by `RapidPy/rapid_main/tests/test_geometry.py`. |
| `MeasurementBlock` | Sequence-Model Owner | Block model is represented in `rapid_main` queue compiler data structures and output generation. |
| `MeasurementBlocks` | Sequence-Model Owner | Measurement block container behavior is represented in queue assembly/validation and regression-tested. |
| `modStatusCode` | Status-Taxonomy Owner | Error/status code mapping is harmonized with shell logging and operator-facing state transitions. |
| `frmVRM` | VRM/Logger Integration Owner | Duplicate: one owner-driven plan created for the above `frmVRM` behavior; evidence captured in single merged plan entry. |
| `frmDAC_Comm` | DAQ/Interface Owner | DAC communication settings are represented through unified config/service behavior or explicitly retired. |
| `frmIRM_VoltageCalibration` | IRM/ARM Workflow Owner | IRM voltage calibration workflow is available in `rapid_main` with approval checkpoints and output recording. |
| `frmShutdownMsg` | Safety/Shutdown Owner | Completed | `rapid_main/rapid_main/app.py` now blocks close when active automation is running until operator confirms; close/menu exit path funnels through `_request_shutdown`, halts worker+queue, and persists layout safely. |
| `modThermal` | Thermal Migration Owner | Same as above thermal row; thermal contracts appear in shared services or are explicitly retired. |
| `InterpolationRange` | Fitting Utility Owner | Interpolation model behavior used by workflow computations is replaced by tested replacements. |
| `InterpolationRanges` | Fitting Utility Owner | Array/range interpolation behavior in sequence and measurement math is represented and regression-tested. |
| `frmTransverseProbeAutoPosition` | Transverse Mechanics Owner | Transverse probe auto-position flow is represented with manual equivalent controls if automation remains not possible. |
| `AngleVsField_Point` | Transverse Mechanics Owner | Transverse angle-vs-field point model is represented in one source of truth with transform tests. |
| `AngleVsFieldCollection` | Transverse Mechanics Owner | Transverse curve/collection model is represented in shared config or sequence services with fixture parity. |
| `ProbeAngleOptimizer` | Transverse Mechanics Owner | Optimization behavior is represented in `rapid_main` and produces auditable outputs or is intentionally retired. |
| `modListenAndLog` | Telemetry/Logging Owner | Logging transport bridge abstraction is integrated or formally retired with replacement logging path documented. |

## Evidence index (first-party)

| Artifact | Purpose |
|---|---|
| `RapidPy/rapid_main/rapid_main/measurement_worker.py` | Sequence engine + error/abort behavior |
| `RapidPy/rapid_main/rapid_main/panels/measurement.py` | Live measurement operator controls |
| `RapidPy/rapid_main/rapid_main/hardware_contracts.py` | Backend contract and preflight abstraction |
| `RapidPy/rapid_main/rapid_main/queue_compiler.py` | Queue command compiler from sample metadata |
| `RapidPy/rapid_main/tests/test_queue_and_bundle.py` | Contract test coverage for bundle + queue compile path |
| `RapidPy/rapid_main/rapid_main/io/measurement_bundle.py` | Multi-format output atomicity by step |

## Ownership and status tracking

This matrix is designed to be updated as each cell transitions toward "Integrated":
- update `Migration Status`
- add evidence rows under `Evidence`
- track validation artifacts under `Notes` until operator review is complete.

Current roadmap target is to complete the AF workflow in this file first, then scale through adjacent core controls (SQUID, changer, vacuum, IRM/ARM).

