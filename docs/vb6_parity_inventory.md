# VB6 Capability Parity Inventory (RapidPy Migration Tracker)

## Status

This matrix is the executable evidence artifact for roadmap Phase 1.  
All capabilities that currently drive routine scientific workflows in `VB6/Paleomag v3.vbp` are mapped to an evidence-backed target in `RapidPy/rapid_main` or marked for explicit retirement.

Legend for `Migration Status`:

- **Not assessed** — inventory exists in VB6 but no matching implementation has been inventoried yet.
- **Mapped** — destination exists in `rapid_main` but not yet fully production-ready.
- **Stand-alone implemented** — implemented in a separate standalone app; not yet integrated into `rapid_main`.
- **Integrated** — implemented in `rapid_main` with a real-or-simulated service path.
- **Not required** — no longer relevant to current active VB6 workflow, with explicit replacement/retirement rationale recorded.

## Current Baseline Mapping

| VB6 Source / Function | Mapping in RapidPy | Migration Status | Evidence (`rapid_main`) | Notes |
|---|---|---|---|---|
| `frmMagnetometerControl` + `modFlow` + `modProg` | Main app control model in `rapid_main.app`, `rapid_main.panels.dashboard`, runtime controls in `MainWindow`, `RuntimeEstimator` | Integrated (software-verified, hardware-unvalidated) | `RapidPy/rapid_main/rapid_main/app.py`, `panels/dashboard.py`, `tests/test_dashboard_shell.py` | Dashboard state follows the actual workflow/sample/step timer; Flow menu actions invoke pause/resume/halt; New Session safely resets transient work only while idle; shutdown halts only after confirmation. Backend cards render live, simulated, disconnected, unavailable, and fault states from the shared diagnostic snapshot. Physical workflow acceptance remains open. |
| `frmProgram` + Rockmag routine logic | `rapid_main.panels.sequence` | Integrated foundation | `RapidPy/rapid_main/rapid_main/panels/sequence.py`, `rapid_main/rockmag.py`, `rapid_main/io/sequence_io.py`, `tests/test_sequence_rockmag.py`, `tests/test_rockmag_routine.py`, `tests/test_sequence_documents.py` | Generates labels, atomic save/load, strict import errors, unsaved-state prompts, validation preview, deterministic Rockmag planning artifacts, a compiled AF25–AF800 Hawaiian preset, and an operator-facing "Rockmag the Works" preset. Imported steps remain the active executable/saveable sequence. Live hardware bundle acceptance remains open. |
| `frmMeasure` / `modMeasure` / `frmStats` | `rapid_main.panels.measurement` + `rapid_main.measurement_worker` + `rapid_main.acquisition` | Integrated (code-complete, hardware-unvalidated) | `RapidPy/rapid_main/rapid_main/acquisition.py`, `panels/measurement.py`, `measurement_worker.py`, `magnetometer.py`, `tests/test_acquisition.py`, `tests/test_vb6_parity.py`, `tests/test_measurement_panel.py` | `BracketedAcquisitionService` runs the exact `Measure_ReadSample` order: verified lift to zero, `CLP`/`RC`, ARC delay, zero-before latch, lift to measurement position, four verified 0/90/180/270 orientations, return to zero, closing 360 turn, zero-after latch. Reduction applies the `i/5` baseline weights, holder subtraction, the `CorrectedSample` rotations for both directions, and reports Average/induced/Fischer CSD/Sig ratios. Sig/Holder and Sig/Induced are rendered from the block. Physical acceptance is still required. |
| `frmSquid` + magnetometer helpers | `rapid_main.squid_transport` + `rapid_main.acquisition` + `updown_control.RawSquidClient` | Integrated (code-complete, hardware-unvalidated) | `RapidPy/rapid_main/rapid_main/squid_transport.py`, `RapidPy/updown_control/updown_control/app.py`, `tests/test_squid_transport.py` | The atomic 2G operations are exposed separately, mirroring `frmSQUID`: `latch` (LC then LD with the legacy pauses), `read_axis` (counter before DVM, both components and both raw replies kept), `clear_and_reset` (CLP/RC), and `set_range` (CSE+CR1 for flux mode). Each latch mints an identity stamped onto its three axis replies, so a partial, stale, or mismatched read is rejected. Live serial timeout/retry evidence remains a hardware gate. |
| `frmChanger*` set | Not in `rapid_main` shell, dedicated standalone app | Stand-alone implemented | `RapidPy/changer_xy_control` | Functionality present and usable but still outside unified app in this phase. |
| `frmDCMotors` + `modMotor` + XY changers | Integrated foundation | Software-complete transport/state evidence; hardware-unvalidated | `RapidPy/rapidpy_common/hardware.py`, `RapidPy/dc_motor_control`, `RapidPy/updown_control`, `RapidPy/changer_xy_control`, `RapidPy/rapid_main`, `tests/test_motor_transport.py`, `tests/test_diagnostic_services.py`, `tests/test_queue_hardware_backend.py`, `tests/test_dcmotors_telemetry.py` | The shared Quicksilver client records exact commands/replies/errors and rejects empty, partial, non-ASCII, or malformed replies. Partial connection setup closes cleanly. Integrated and queue adapters reject unsuccessful results before state mutation, and safe return attempts every halt after earlier failures. Live telemetry, torque fallback, truthful semantic states, and accessible controls remain. Physical direction, encoder/torque, limits/stall, transfer, abort, and transcript acceptance remain open. |
| `frmVacuum` | Stand-alone implemented | Software-complete command path; hardware-unvalidated | `RapidPy/updown_control` (vacuum logic), `RapidPy/rapid_main/rapid_main/dialogs/vacuum.py`, `diagnostic_services.py`, `measurement_worker.py`, `tests/test_vacuum_transport.py`, `tests/test_diagnostic_services.py`, `tests/test_measurement_worker.py` | Hardware mode has no silent simulation fallback, requires CR-terminated acknowledgments, records exact TX/RX/errors, and never fabricates pressure. The legacy controller has no pressure telemetry: positive thresholds block, while threshold zero plus confirmed pump-on is the explicit command-only opt-out. Physical pump/valve round-trip and an independent pressure adapter/interlock remain open. |
| AF treatment (`frmAF`, ADwin AF controls, AF tuners) | Mapped | Software-complete live adapter evidence; hardware-unvalidated | `RapidPy/af_tuner`, `RapidPy/af_clip_test`, `_launch_af` in `app.py`, `diagnostic_services.AfDemagBackendAdapter`, `QueueHardwareBackend.set_demag_step`, `tests/test_diagnostic_services.py`, `tests/test_queue_hardware_backend.py` | AF workflow has shell/queue routes and explicit command planning. Live-mode construction requires an ADwin readiness probe; complete structured ramp request/result/error evidence and confirmed relay reset merge into the run transcript, while malformed results and simulated sources fail closed. Physical rig/interlock/abort acceptance remains pending. |
| IRM/ARM (`frmIRMARM`, voltage calibration forms) | Mapped | Software-complete live adapter evidence; hardware-unvalidated | `RapidPy/rapid_main/rapid_main/dialogs/irm_arm.py`, `diagnostic_services.IrmArmBackendAdapter`, `hardware_contracts.py`, `panels/calibration.py`, `tests/test_diagnostic_services.py`, `tests/test_queue_hardware_backend.py` | Each incremental IRM/ARM ADwin ramp records complete structured request/result/error evidence after a readiness probe, safe relay reset requires confirmation, and live events merge into `communication.tsv`. The glass dialog remains truthful when unavailable. Physical calibration, interruption/safe-abort, and execution acceptance remain open. |
| `frm908AGaussmeter` | Launcher path + Calibration Center baseline artifact | Integrated foundation | `RapidPy/gaussmeter_control`, `rapid_main/dialogs/menus`, `rapid_main/calibration.py`, `panels/calibration.py`, `tests/test_calibration.py`, `tests/test_calibration_panel.py` | Standalone gaussmeter app remains available; rapid_main Calibration Center now records SQUID/Gaussmeter baseline JSON artifacts with samples/residuals, operator, and run context. Live meter/standard validation remains open. |
| Thermal (`modThermal` + associated forms) | Integrated planning; live route fail-closed | Software-verified planning, hardware-unvalidated | `rapid_main/thermal.py`, `hardware_contracts.py`, `measurement_worker.py`, `panels/calibration.py`, `tests/test_thermal.py`, `tests/test_calibration_panel.py`, `tests/test_queue_and_bundle.py`, `tests/test_measurement_worker.py`, existing TT/TH/TEMP IO/runtime support | Thermal labels and planning are represented with validated temperature targets, ramp/hold/cool estimates, safety limits, queue-compatible labels, measurement-block metadata, a Calibration Center planning artifact path, and an optional typed `apply_thermal` hook. Hardware-mode queues now block before ordinary preflight when no production furnace/oven adapter exists. Live control, readback/interlocks, safe abort, and acceptance remain open. |
| Susceptibility (`Susceptibility` forms) | Dedicated glass diagnostics + typed serial boundary | Auditable diagnostics; queue route fail-closed | `susceptibility_transport.py`, `diagnostic_services.py`, `dialogs/susceptibility.py`, `config.py`, `legacy_ini.py`, `hardware_contracts.py`, `measurement_worker.py`, `tests/test_susceptibility_transport.py`, `tests/test_diagnostic_services.py`, `tests/test_hardware_dialog_glass.py`, `tests/test_measurement_worker.py` | The former menu misroute to SQUID is replaced by a dedicated bridge dialog. Exact VB6 Z/M commands, CR replies, scale factor, connection/errors, and immutable evidence are implemented; legacy serial/calibration/coil settings migrate. SQUID moment is explicitly rejected as susceptibility, and requested SUSC failures abort rather than becoming zero. Automated live SUSC remains blocked pending validated coil motion, holder subtraction, calibration, safe return, and physical acceptance. |
| `frmVRM` (VRM routines) | Integrated via ownership-safe external launch | Software-verified; hardware-unvalidated | `rapid_main.MainWindow._launch_vrm` validates context, releases retained SQUID clients, reserves process-lifetime ownership, and launches `vrm_logger`; `test_vrm_integration.py` covers the handoff and finalized manifest contract. | VRM remains in the dedicated `RapidPy/vrm_logger` app. Physical acquisition and production-run association remain acceptance gates. |
| Calibration rod / AF/IRM tuning forms | Parameter tabs + launchers | Mapped | `RapidPy/rapid_main/rapid_main/panels/settings_panel.py`, `RapidPy/rapid_main/rapid_main/panels/calibration.py`, `RapidPy/rapid_main/rapid_main/calibration.py` | Calibration workflows now route through the rapid_main Calibration Center (automated/manual modes). |
| Plots/stats (`frmPlots*`, Zijderveld, stereonet, quicklooks) | Integrated + PCA/decay review + validated export | Mapped | `RapidPy/data_viewer`, `rapid_main.analysis`, `rapid_main.dialogs.plots`, `tests/test_analysis.py`, `tests/test_plots_dialog.py`, `panels/measurement.py`, `tests/test_measurement_panel.py`, `tests/test_output_integrity.py` | Measurement panel retains completed steps and opens Zijderveld/equal-area/intensity plus accessible PCA/decay review. North/East/Down principal direction, variance, residual, moment statistics, and decay evidence are carried into atomic provenance-marked JSON/CSV. Run quicklooks share the bundle's simulation boundary, and `artifact_index.json` is reconciled after publication. Real-bundle review and any legacy-specific plot types remain acceptance gates. |
| Data saving/export (`frmFileSave`, directories) | `rapid_main.io.measurement_bundle` + output writers | Integrated (core path) | `RapidPy/rapid_main/rapid_main/io/measurement_bundle.py`, `measurement_worker.py` | `.sample`, `.rmg`, `measurements.txt`, `specimens.txt` writes work in tandem; `artifact_index.json` now records required, optional, present, and missing run artifacts for completed and early-aborted worker exits. |
| Webcam (`frmWebcam`) | Integrated optional dialog + own app | Integrated | `RapidPy/webcam_viewer`, `RapidPy/rapid_main/rapid_main/dialogs/webcam_dialog.py`, `tests/test_data_dialog_glass.py` | Main-shell launcher opens the shared glass dialog, fits it with the correct `QtGui.QScreen` contract, and explains a missing optional dependency instead of failing. Live camera/device acceptance remains open. |
| Settings/INI editors (`frmSettings*`, `frmOptions`, channel config) | Structured JSON settings + settings panel | Mapped | `RapidPy/rapid_main/rapid_main/config.py`, `panels/settings_panel.py`, `RapidPy/rapid_main/rapid_main/legacy_ini.py` + `panels/settings_panel.py` (Import VB6 INI…) | VB6 INI import routes into app settings with mapping warnings. Versioned JSON settings backup/restore is atomically written, strictly validated before mutation, blocked during active work, and explicitly requires restart before restored hardware settings are used. |
| Cross-version operator task lookup | Searchable VB6 → RapidPy task map | Integrated | `RapidPy/rapid_main/rapid_main/dialogs/transition_help.py`, Help menu routing in `app.py`, `tests/test_transition_help.py` | Common legacy forms and operator tasks are searchable in-app; each result shows its RapidPy destination and hardware/software readiness, and Open Destination routes to the real panel or launcher. |
| diagnostics (`frmDebug`, DAQ/ADwin comm debug, step monitor, messages) | Integrated foundation + separate standalone apps | Mapped | `RapidPy/rapid_main/rapid_main/dialogs`, `diagnostic_services.collect_diagnostic_status`, `tests/test_diagnostic_services.py`, `tests/test_remaining_dialog_glass.py` | Debug Console refreshes structured backend readiness, escapes backend text before rich rendering, and exposes accessible filtering/copy/refresh actions. Step Monitor has written execution state, normalized progress, compact actions, and correct screen fitting. Live adapter-specific raw DAQ/ADwin troubleshooting still requires hardware acceptance evidence. |

## P0 measurement-integrity addendum (2026-08-29)

| VB6 element | RapidPy destination | Status | Notes |
|---|---|---|---|
| `MeasurementBlock` (baselines, samples, holder, direction) | `magnetometer.BracketedMeasurementBlock` + `BracketedMeasurementResult` | Code-complete | Adds coherent per-axis counter/DVM evidence, a latch identity, and a block audit record. |
| `MeasurementBlock.BaselineAdjustedSample` | `reduce_bracketed_measurement` | Code-complete | `(1 - i/5)` / `i/5` weights and per-position holder subtraction pinned by `tests/test_vb6_parity.py`. |
| `MeasurementBlock.CorrectedSample` | `reduce_bracketed_measurement` | Code-complete | Both `mvarDirection` values pinned. |
| `MeasurementBlock.induced` / `Kappa` / `FischerSD` / `SigDrift` / `SigHolder` / `SigInduced` | `BracketedMeasurementResult` | Code-complete | Same `1e-9` denominator guards as VB6. |
| `Measure_ReadSample` motion and latch order | `acquisition.BracketedAcquisitionService` | Code-complete | One blocking verified motion replaces the VB6 non-blocking-start/blocking-repeat pair. |
| `Measure_HasZeroDiscontinuity` / `Measure_HasHolderStaircase` | `validate_zero_pair` + monotonic-staircase guard | Code-complete | Categorical rejection; never relaxed after N retries. |
| `CLP` / `ResetCount` retry loop | `recover_flux_count_discontinuity` | Code-complete | Returns to verified zero, resets, waits, re-latches, restarts the whole block. |
| Global `Holder` object | `holder_state.HolderCorrection` + `HolderStateStore` | Code-complete | Persisted, atomically installed, retained on any failure, and required before sample measurement. |
| `modMeasure` holder averaging (`MeasurementBlocks.AverageBlock`) | `holder_measurement.HolderMeasurementService` | Code-complete | Averages baseline-adjusted vectors across cycles, as VB6 does. |
| `frmSendMail.MailNotification` retry mails | Retired | Retired | Rejection is surfaced as an operator error and recorded in the run artifacts instead. |

None of these rows is hardware-validated. See
`docs/rapid_hardware_acceptance_procedure_2026-08-29.md`.

## Full `VB6/Paleomag v3.vbp` Component Sweep (2026-07-10)

The table below captures every `Form`, `Module`, and `Class` entry from
`VB6/Paleomag v3.vbp` and current migration status to support complete
Phase-1 inventory closure.

| VB6 item | Type | Migration status | RapidPy location / notes |
|---|---|---|---|
| `frmMagnetometerControl` | Form | Mapped | `rapid_main` shell + dashboard + flow scaffolding |
| `frmAbout` | Form | Integrated | `rapid_main/dialogs/about.py`, `MainWindow` Help menu, `tests/test_general_dialog_glass.py`; responsive shared glass surface with runtime-derived version details and accessible actions |
| `frmLogin` | Form | Mapped | `rapid_main` main menu + `dialogs/login.py` launcher, `tests/test_general_dialog_glass.py`; required identity, explicit No-Comm semantics, responsive form, and accessible session actions are covered |
| `frmSplash` | Form | Integrated | `rapid_main/dialogs/startup_guide.py`, non-modal first-run display after main-window fit, `tests/test_startup_guide.py`, and `tests/test_general_dialog_glass.py`; responsive glass guidance and routing actions are covered |
| `frmTip` | Form | Integrated | `StartupGuideDialog` is available from Help > Quick Start, persists its startup preference, and routes to Settings or Sample Queue; the adjacent searchable `TransitionHelpDialog` answers “Where did this VB6 control go?” and routes to real destinations; covered by `tests/test_startup_guide.py` and `tests/test_transition_help.py`. |
| `modProg` | Module | Mapped | flow/runtime scaffolding in shell |
| `modMeasure` | Module | Mapped | `measurement_worker` + `panels.measurement` |
| `modVector3d` | Module | Mapped | `RapidPy/rapid_main/rapid_main/geometry.py`; `measurement_worker.py`; `io/specimen_reader.py` |
| `modMotor` | Module | Stand-alone implemented | `updown_control` / `dc_motor_control` launcher path |
| `frmChangerSampOrder` | Form | Mapped | `rapid_main/rapid_main/queue_compiler.py`, sample-queue panel UI | Duplicate position and invalid sample-step validation are now tracked in the compiler contract (`QueueValidationResult`) with regression tests. |
| `modChanger` | Module | Stand-alone implemented | `changer_xy_control`/`updown_control` |
| `frmMeasure` | Form | Mapped | `panels.measurement` |
| `frmStats` | Form | Integrated foundation | `panels/measurement.py`, `measurement_worker.py`, `analysis.py`, `tests/test_measurement_panel.py` | Live cycle mean, X/Y/Z ranges, range-to-moment ratios, directional spread, and signal-to-drift are rendered from configured SQUID reads. Sig/Holder and Sig/Induced now come from the reduced bracketed block, and holder identity, magnitude, and age are shown; they stay `N/A` when no bracketed block or valid holder exists. |
| `modPrint` | Module | Mapped | `rapid_main/io/measurement_bundle.py`, `measurement_worker.py`, `measurement output writers` |
| `frmVacuum` | Form | Software-complete command path | Shared glass dialog plus fail-closed acknowledged legacy controller; independent pressure integration and physical acceptance remain open. |
| `frmDCMotors` | Form | Integrated + auditable live adapter | `rapid_main/dialogs/dc_motors.py`, `diagnostic_services.py`, `rapidpy_common/hardware.py`, `tests/test_motor_transport.py`, `tests/test_dcmotors_telemetry.py`, `tests/test_remaining_dialog_glass.py`; exact transport evidence, unsuccessful-result refusal, live/simulated/unavailable state, telemetry, and fail-closed controls are software-verified, while physical motion/interlock/encoder/torque acceptance remains open. |
| `frmAF_2G` | Form | Stand-alone implemented | AF calibration/treatment apps exist; main path not integrated |
| `frmSendMail` | Form | Not required | Retired automatic email dispatch. Modern replacement is operator-controlled sharing of measurement bundle artifacts written by `rapid_main/io/measurement_bundle.py` and `measurement_worker.py` (`.sample`, `.rmg`, `measurements.txt`, `specimens.txt`); no hardware-control or run-safety dependency remains. |
| `frmSquid` | Form | Stand-alone implemented | `gaussmeter`/meter abstraction pending in shell |
| `frmOptions` | Form | Mapped | `panels.settings_panel` |
| `modFlow` | Module | Mapped | state-machine draft and timer scaffolding in `app.py` |
| `frmStepMonitor` | Form | Integrated foundation | `dialogs/step_monitor.py`, `tests/test_remaining_dialog_glass.py`; queue-step values, normalized progress, written execution state, event log, compact actions, and screen-fit behavior are covered. Field restart/recovery acceptance remains open. |
| `modDataAnalysis` | Module | Mapped | `rapid_main/analysis.py` and `tests/test_analysis.py` provide deterministic vector extraction from `MeasurementStep.sdx/sdy/sdz`, centroid/PCA-style principal-axis evidence, orientation/variance/residual metrics, and moment statistics/decay summaries. Advanced plotting surfaces and hardware-run bundle review remain separate acceptance caveats. |
| `modMagnetometer` | Module | Integrated foundation | `rapid_main/magnetometer.py`, `measurement_worker.py`, `tests/test_magnetometer.py`, and `tests/test_measurement_worker.py` normalize raw SQUID axis voltages into calibrated/background-corrected moment vectors with quality flags, and the measurement loop now accepts `MagnetometerReading` payloads while surfacing quality flags as operator warnings. Live SQUID adapter rollout and hardware timeout/retry evidence remain acceptance gates. |
| `modAF_2G` | Module | Stand-alone implemented | AF treatment app remains external |
| `frmIRMARM` | Form | Integrated + auditable live adapter | `rapid_main/dialogs/irm_arm.py`, `diagnostic_services.py`, `hardware_contracts.py`, `panels/calibration.py`, `tests/test_diagnostic_services.py` | Integrated Diagnostics route, backend ownership, calibration artifacts, fail-closed ADwin readiness, exact ramp result/error evidence, and confirmed reset exist; physical execution, interruption recovery, and acceptance remain open. |
| `frmRockmagRoutine` | Form | Integrated foundation | `rapid_main/rockmag.py`, `rapid_main/panels/sequence.py`, `tests/test_rockmag_routine.py`, and `tests/test_sequence_rockmag.py` compile operator-level rockmag routine specs/templates into `RockmagSteps`, family-preserving `MeasurementBlock`s, runner-compatible labels, deterministic JSON planning artifacts, and the Sequence panel "Rockmag the Works" preset. Full hardware execution and reproducible rockmag run bundles remain acceptance gates. |
| `RockmagStep` | Class | Mapped | `rapid_main/data_model.py` defines `RockmagStep` with label family/value/unit parsing for NRM, AF, IRM, ARM, RRM, backfield, thermal, and susceptibility labels. |
| `RockmagSteps` | Class | Mapped | `rapid_main/data_model.py` defines `RockmagSteps` with ordered label preservation, family summary, and conversion to `MeasurementBlock`; tests cover routine flattening. |
| `frmSusceptibilityMeter` | Form | Auditable diagnostic transport; automated motion blocked | Dedicated accessible glass dialog plus exact Z/M serial transport, strict CR-response validation, migrated settings, and immutable communication evidence. Queue execution still requires coil motion and holder-correction integration. |
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
| `modSusceptibility` | Module | Diagnostic protocol integrated; queue route fail-closed | Explicit SUSC-only reads, abort-on-failure behavior, `.rmg`/`susceptibility.json` persistence, and a typed bridge boundary are implemented. Holder subtraction, coil positioning, calibration, and safe return remain open. |
| `Angular3D` | Class | Mapped | `RapidPy/rapid_main/rapid_main/geometry.py` |
| `MeasurementBlock` | Class | Mapped | `rapid_main/data_model.py` (`MeasurementBlock`) preserves named block labels, type, and metadata; `test_queue_and_bundle.py` verifies normalization and runner-label flattening. |
| `MeasurementBlocks` | Class | Mapped | `rapid_main/data_model.py` (`MeasurementBlocks`) preserves ordered block collections and flattens to current queue/runner label contracts with regression coverage. |
| `modStatusCode` | Module | Mapped | `rapid_main/status_codes.py` defines stable operator status/severity codes; `measurement_worker.py` emits structured `status_event` objects alongside existing UI strings; `tests/test_status_codes.py` verifies phase, warning, SQUID-error, and safety-return classification. |
| `frmSampleQueueMonitor` | Form | Mapped | queue monitor panel scaffolding |
| `frmVRM` | Form | Software-integrated; hardware-unvalidated | `_launch_vrm` refuses active/conflicting ownership, releases retained main SQUID clients, and holds measurement/SQUID leases for the child lifetime. The logger finalizes an immutable outcome/hash manifest after CSV closure and distinguishes validated handoff from true run association. Physical acquisition remains open. |
| `frmPlots` | Form | Stand-alone implemented | data review UI in `data_viewer` |
| `frmWebcam` | Form | Stand-alone implemented | `dialogs.webcam` launch path |
| `modConfig` | Module | Mapped | `config.py` migration started |
| `CIniFile` | Class | Mapped | `RapidPy/rapid_main/rapid_main/legacy_ini.py` + Settings import workflow (`Import VB6 INI…`) with mapping notes and validation. |
| `frmAFTuner` | Form | Stand-alone implemented | `RapidPy/af_tuner` |
| `frmCalibrateCoils` | Form | Mapped | `rapid_main/panels/calibration.py`, `rapid_main/calibration.py`, `rapid_main/app.py` stack/navigation wiring |
| `frmDAC_Comm` | Form | Integrated foundation + startup readiness contract | `rapid_main/dac.py`, `diagnostic_services.py`, `tests/test_dac.py`, and `tests/test_diagnostic_services.py` provide a stable DAC command surface, safe-zero startup reports, missing-adapter error surfacing, and IRM/ARM ADwin field-to-voltage planning through explicit DAC commands, blocking out-of-range requests without silent clamping. Live MCC/DAQ binding and hardware loopback acceptance remain open. |
| `frmFileSave` | Form | Mapped | output writers + measurement bundle |
| `ADWIN` | Module | Integrated live adapter foundation | `rapidpy_common/adwin_af.py` plus fail-closed AF/IRM adapters in `diagnostic_services.py`; physical board/coil acceptance remains open. |
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
| `frmDebug` | Form | Integrated foundation | `rapid_main/dialogs/debug_console.py`, `diagnostic_services.collect_diagnostic_status`, `tests/test_diagnostic_services.py`, `tests/test_remaining_dialog_glass.py` | The Debug Console refreshes structured readiness for Vacuum, AF Demag, IRM/ARM, SQUID, and DC Motors, uses the shared glass/accessibility contract, and safely renders backend text. Raw DAQ/ADwin troubleshooting remains hardware-acceptance-gated. |
| `frmIRM_VoltageCalibration` | Form | Integrated foundation + governed calibration record | `rapid_main/calibration.py` defines IRM field-to-voltage calibration points, linear fit results, voltage-limit request validation, and zero-intercept legacy-table support; `IrmArmConfig` and Settings retain slope/intercept/limit values; `diagnostic_services.py` uses those coefficients during IRM/ARM DAC planning; `calibration_registry.py` and `panels/calibration.py` provide immutable artifact snapshots, versioned operator approval, expiry/invalidation, event-based rollback, hash verification, and measurement provenance references. Live DAC binding, reference-standard acceptance, interruption recovery, and hardware execution remain acceptance gates. |
| `frmShutdownMsg` | Form | Mapped | `rapid_main/rapid_main/app.py` shutdown action path (`_confirm_shutdown`, `_request_shutdown`, menu/toolbar exit wiring) and `DeviceOwnershipManager`-safe halt flow on exit. |
| `frmCalRod` | Form | Stand-alone implemented | coil/rod calibration parity pending |
| `modAF_DAQ` | Module | Stand-alone implemented | AF DAQ bridge in external project |
| `frmINIConverter` | Form | Mapped | VB6 INI converter behavior is replaced by the `Import VB6 INI…` action in `SettingsPanel`, with mapping/reporting and preserved operator review of unresolved keys. |
| `modThermal` | Module | Integrated planning; live route fail-closed | `rapid_main/thermal.py`, `hardware_contracts.py`, `measurement_worker.py`, `panels/calibration.py`, `tests/test_thermal.py`, `tests/test_calibration_panel.py`, `tests/test_queue_and_bundle.py`, and `tests/test_measurement_worker.py` provide thermal parsing, safety validation, timing estimates, queue conversion, planning artifacts, and an optional `apply_thermal` route. Treatment-plan preflight blocks live execution without a production furnace/oven adapter, and unsupported labels cannot fall through to a generic acknowledgment. Live hardware integration remains open. |
| `frmXYHoming` | Form | Stand-alone implemented | changer homing logic in standalone app |
| `IRMData` | Class | Stand-alone implemented | IRM model migration pending |
| `IrmDataPoint` | Class | Stand-alone implemented | IRM model migration pending |
| `InterpolationRange` | Class | Mapped | `rapid_main/geometry.py` defines tested linear segment interpolation with validation, clamping, and descending-axis support. |
| `InterpolationRanges` | Class | Mapped | `rapid_main/geometry.py` defines ordered interpolation segment lookup and clamped boundary behavior; `tests/test_geometry.py` covers range lookup and failures. |
| `XYCup` | Class | Stand-alone implemented | XY changer model external |
| `XYCup_Positions` | Class | Stand-alone implemented | XY changer model external |
| `frmTransverseProbeAutoPosition` | Form | Mapped foundation | `rapid_main/transverse.py` and `tests/test_transverse.py` convert angle-vs-field scans into an auditable auto-position plan: optimized target angle, shortest signed move, tolerance/hold behavior, confidence blocking, and command summary. Live motor execution remains gated by hardware ownership and acceptance evidence. |
| `AngleVsField_Point` | Class | Mapped | `rapid_main/data_model.py` defines `AngleVsFieldPoint`; tests cover value normalization and use in ordered collections. |
| `AngleVsFieldCollection` | Class | Mapped | `rapid_main/data_model.py` defines `AngleVsFieldCollection`; tests cover ordered angle/field access, peak selection, and weighted center angle. |
| `ProbeAngleOptimizer` | Class | Mapped | `rapid_main/data_model.py` defines deterministic `ProbeAngleOptimizer` and `ProbeAngleOptimizationResult`; tests verify peak-field angle selection, confidence scoring, and empty-collection behavior. |
| `AdwinAfInputParameters` | Class | Stand-alone implemented | AF parameter model in external AF stack |
| `AdwinAfOutputParameters` | Class | Stand-alone implemented | AF parameter model in external AF stack |
| `AdwinAfParameter` | Class | Stand-alone implemented | AF parameter model in external AF stack |
| `AdwinAfPauseConstants` | Class | Stand-alone implemented | AF pause constants in external AF stack |
| `AdwinAfRampStatus` | Class | Stand-alone implemented | AF ramp-status model external |
| `modLogAFParameters` | Module | Stand-alone implemented | AF log parser logic external |
| `modListenAndLog` | Module | Integrated + core live-adapter evidence | `rapid_main/communication_log.py`, `squid_transport.py`, `diagnostic_services.py`, `hardware_contracts.py`, `measurement_worker.py`, `rapidpy_common/hardware.py`, `tests/test_communication_log.py`, `tests/test_squid_transport.py`, `tests/test_motor_transport.py`, `tests/test_diagnostic_services.py`, `tests/test_measurement_worker.py`, and `tests/test_queue_hardware_backend.py` provide deterministic per-run `communication.tsv`. SQUID raw traffic, vacuum commands/acknowledgments, structured AF/IRM ADwin requests/results/errors, and exact DC motor command/reply/error traffic merge immutably; current-run cursors and simulation exclusion prevent false live evidence. Retained auxiliary adapters still need equivalent adoption and physical transcript evidence. |

## Immediate action items for objective completion

1. Keep formerly "Not assessed" VB6 behaviors tied to owner + acceptance criteria until hardware/run-bundle closure is complete.
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
| SQUID comm + readings | Software-complete transport/recovery foundation | Diagnostic and fail-closed backend contracts exist; exact raw traffic is retained; stale/partial replies are rejected; transport faults trigger bounded safe-zero/reset/backoff and whole-block reacquisition; motion failures are not retried; recovery records and operator warnings are bundled | Physical SQUID timeout/recovery/safe-abort fault injection and transcript acceptance remain required. |
| Motors (changer/XY/up-down/turning) | Software-complete transport/state evidence | Shared Quicksilver logging records exact commands, terminated replies, and errors; malformed/partial replies fail; unsuccessful motion cannot mutate queue direction/hole/load state; safe return attempts all halts; telemetry/ownership/configuration remain integrated | Physical direction, encoder/torque, limit/stall, sample-transfer, safe-abort, and transcript acceptance against the hardware command set. |
| Vacuum control | Software-complete fail-closed command path | Exact legacy-controller commands, CR-terminated acknowledgments, and errors are recorded; no live pressure is fabricated; positive thresholds block command-only operation | Physical pump/valve round-trip plus an independent pressure adapter/interlock and controlled fault injection if pressure gating is retained. |
| IRM / ARM + calibrations | Integrated auditable ADwin adapter | Live readiness probe, complete structured per-ramp request/result/error evidence, confirmed reset, queue routing, calibration artifacts, and run-transcript merge | Physical calibration, coil/DAC loopback, interruption/safe-abort, and execution acceptance. |
| AF treatment/tuning | Integrated auditable ADwin adapter | AF planner/queue route plus readiness probe, complete structured request/result/error evidence, malformed-result rejection, confirmed reset, and run-transcript merge | Physical rig execution, telemetry/decay, interlock, abort, and transcript acceptance. |
| Webcam/DAQ/Susceptibility/VRM | Mixed integrated foundations | Stand-alone entry points remain available | Susceptibility has a dedicated in-shell diagnostic transport and truthful queue blocker. VRM adds exclusive process-lifetime ownership plus finalized immutable outcome/hash evidence; missing `run_id` remains explicitly unassociated. Automated susceptibility mechanics and physical VRM/output acceptance stay open. |

### Automated sample handling readiness

| VB6 behavior | RapidPy status | Remaining parity work |
|---|---|---|
| `frmSampleQueueMonitor` + run control loop | Integrated and simulator-driven | Queue state persists, interrupted `Running` rows restore as explicit `Interrupted` operator decisions, and pause/halt paths retain safe-state return semantics; hardware field replay acceptance remains open. |
| `frmChangerSampOrder` validation | Queue compiler validation in place (`compile_queue`, strict mode, tests) | Expand duplicate/invalid-position diagnostics in operator-facing diagnostics, include hard stop if mandatory prerequisites missing. |
| `frmRerunSamples` | Integrated | Failed samples can be reset and interrupted rows now expose explicit Resume/Re-run/Skip/Abort recovery actions; live restart semantics for unsafe machine states still need hardware acceptance. |
| `MeasurementBlock` / step mapping | Mapped in `data_model.py` and compiler output path | Continue integrating block metadata into rock-magnetic execution and VB6-equivalent report annotations in exported bundles. |

#### Current blockers to close before main-shell parity claim

- AF and IRM/ARM treatment now have fail-closed, auditable ADwin adapter routes;
  live rig execution, telemetry/decay, interlock, abort, and calibration evidence
  are still missing.
- SQUID, vacuum, ADwin treatment, and DC motor software paths are
  production-bound and simulation-isolated; retained auxiliary adapters still
  need equivalent run-level evidence, and every live path still needs physical
  acceptance.
- No full end-to-end hardware proof that queue-run execution can resume deterministically from persistent interrupted states under all operator abort paths.
- Measurement runs now write `workflow_summary.json` with emitted phase transitions (`preflight -> loading -> treating -> measuring -> validating -> saving -> returning -> complete/error` as applicable), final aborted/completed state, labels, status codes, and hardware-validation caveat. Worker exits also write `artifact_index.json` so completed, halted, unavailable-backend, and preflight-failed runs have deterministic bundle evidence; live hardware transition and recovery evidence remains required.

### Not Assessed Remediation Queue (owner + acceptance criteria)

The list below tracks capabilities that were originally "Not assessed" and now carry explicit ownership plus acceptance gates.

| VB6 item | Owner | Acceptance criteria |
|---|---|---|
| AF demagnetization workflow (`frmAF_2G`, AF-2G tune/clip helpers) | AF/Automation Owner | Software criterion met: AF labels route through the dedicated planner/ADwin adapter, require readiness, retain complete request/result/error evidence, reject malformed results, and confirm reset. Still need physical treatment/tuning, interlock, abort, and transcript evidence. |
| IRM/ARM (`frmIRMARM`, voltage calibration forms) | IRM/ARM Workflow Owner | Software criterion met for the ADwin treatment boundary: readiness, incremental ramp evidence, propagated failures, confirmed reset, queue routing, and transcript merge exist. Physical calibration, safe-abort, ownership-release, and execution evidence remain. |
| Thermal (`modThermal` + associated forms) | Thermal Migration Owner | Thermal setup/calibration/edit forms are either mapped into an integrated `rapid_main` service or explicitly approved as out-of-scope with rationale in this tracker. |
| `frmVRM` (VRM routines) | VRM/Logger Integration Owner | Software criterion met for exclusive launch ownership, retained-port release, atomic validated context, and immutable finalized `.vrm.json` evidence with outcome, rows, provenance, port, and CSV digest. A handoff without `run_id` is explicitly unassociated. Live logged acquisition with physical run association remains the hardware acceptance gate. |
| Plots/stats (`frmPlots*`, Zijderveld, stereonet, quicklooks) | Data Review Owner | Software criterion met: `rapid_main` exposes on-demand Zijderveld, equal-area, intensity, PCA, moment-statistics, and decay review backed by reusable analysis services, stable tests, and provenance-marked exports. Physical run-bundle review remains a separate acceptance gate. |
| `modVector3d` | Math Utilities Owner | Completed | Orientation/matrix math is ported to `RapidPy/rapid_main/rapid_main/geometry.py` with tests in `RapidPy/rapid_main/tests/test_geometry.py`; used in `measurement_worker` and specimen parsing. |
| `frmSendMail` | Reporting/Run-log Owner | Completed | Automatic email dispatch is explicitly retired. Reporting evidence is preserved through deterministic measurement bundle outputs (`.sample`, `.rmg`, `measurements.txt`, `specimens.txt`) produced by `MeasurementBundleWriter` and `MeasurementWorker`; operators can review/share the output directory without a hidden mail side effect. |
| `modDataAnalysis` | Data Quality Owner | Completed | Shared analysis primitives in `rapid_main/analysis.py` expose vectors, centroid/principal-axis fit, orientation/variance/residual metrics, and moment statistics/decay. `PlotsDialog` now renders and exports that evidence with regression coverage. Hardware-run bundle acceptance remains tracked separately. |
| `modMagnetometer` | SQUID/Measurement Owner | Full SQUID contract implementation includes preflight, safe stop/retry behavior, and measurement read quality flags. |
| `frmRockmagRoutine` | Rock-Magnetic Owner | Rock-magnetic routine behavior is mapped to sequence/measurement flow in `rapid_main` with at least one regression fixture. |
| `RockmagStep` | Rock-Magnetic Owner | Completed | Rock-magnetic step structures are represented by `rapid_main.data_model.RockmagStep` with tested parsing for common generated sequence labels. |
| `RockmagSteps` | Rock-Magnetic Owner | Completed | Rock-magnetic batch-step container structure is represented by `rapid_main.data_model.RockmagSteps`; tests verify ordered labels, family metadata, and conversion to the current `MeasurementBlock` runner contract. |
| `SampleIndexRegistration` | Sample-Index Owner | Completed | `rapid_main/data_model.py` defines `SampleIndexRegistration`; `rapid_main/io/sample_index.py` maps `.sam` registry entries with preserved specimen order metadata. |
| `SampleIndexRegistrations` | Sample-Index Owner | Completed | `rapid_main/data_model.py` defines `SampleIndexRegistrations`; `rapid_main/io/sample_index.py` and `rapid_main/tests/test_sample_index.py` verify order preservation and header handling. |
| `frmSampleIndexRegistry` | Sample-Index Owner | Completed | `.sam` registry loading flow is represented in `SampleSelectDialog` (`rapid_main/dialogs/sample_select.py`), with registry metadata used where present and fixture-backed parsing tests. |
| `Cartesian3D` | Math Utilities Owner | Completed | Vector math primitives and conversion helpers are covered in `RapidPy/rapid_main/geometry.py` and `RapidPy/rapid_main/tests/test_geometry.py`. |
| `frmRerunSamples` | Queue/Workflow Owner | Completed | queue rerun/skip controls are implemented in `SampleQueuePanel` (failure resets, queue-persistent status restore, and skip-aware run extraction) with regression coverage in `RapidPy/rapid_main/tests/test_sample_queue_helpers.py` (`_normalize_status`) and queue execution tests via `test_queue_and_bundle` coverage patterns. |
| `Angular3D` | Math Utilities Owner | Completed | Angular direction conversion helpers are covered in `RapidPy/rapid_main/geometry.py` and validated by `RapidPy/rapid_main/tests/test_geometry.py`. |
| `MeasurementBlock` | Sequence-Model Owner | Completed | Block model is represented by `rapid_main.data_model.MeasurementBlock`; tests verify name/type normalization, metadata preservation, and executable label flattening. |
| `MeasurementBlocks` | Sequence-Model Owner | Completed | Ordered block container behavior is represented by `rapid_main.data_model.MeasurementBlocks`; tests verify block-name preservation, step counting, and runner label flattening. |
| `modStatusCode` | Status-Taxonomy Owner | Completed | Error/status code mapping is harmonized through `rapid_main.status_codes.OperatorStatus`, `StatusCode`, and `StatusSeverity`; `MeasurementWorker.status_event` provides structured status evidence without changing current operator UI strings. |
| `frmVRM` | VRM/Logger Integration Owner | Duplicate: one owner-driven plan created for the above `frmVRM` behavior; evidence captured in single merged plan entry. |
| `frmDAC_Comm` | DAQ/Interface Owner | Integrated foundation + startup readiness contract | `rapid_main.dac` represents channel setup, write/safe-zero command validation, startup readiness reports, and missing-adapter error surfacing; `diagnostic_services` uses it for IRM/ARM DAC voltage planning. Production MCC/DAQ adapter binding and loopback evidence remain under the DAC/MCC integration gate. |
| `frmIRM_VoltageCalibration` | IRM/ARM Workflow Owner | Integrated foundation + operator artifact path | IRM voltage calibration math is represented in `rapid_main.calibration`; retained slope/intercept/limit settings feed `diagnostic_services` DAC planning; Calibration Center records fit artifacts with run context/operator metadata. Live hardware calibration execution, DAC binding, and interruption recovery remain open. |
| `frmShutdownMsg` | Safety/Shutdown Owner | Completed | `rapid_main/rapid_main/app.py` now blocks close when active automation is running until operator confirms; close/menu exit path funnels through `_request_shutdown`, halts worker+queue, and persists layout safely. |
| `modThermal` | Thermal Migration Owner | Integrated planning; live route fail-closed | `rapid_main.thermal` maps routine planning, safety validation, queue labels, and JSON evidence exposed through Calibration Center. `QueueHardwareBackend` calls a typed `apply_thermal` route when provided; otherwise `MeasurementWorker` blocks the hardware-mode plan before device preflight or step execution. Live furnace/oven adapter wiring and acceptance evidence remain open. |
| `InterpolationRange` | Fitting Utility Owner | Completed | Interpolation model behavior used by workflow computations is represented by `rapid_main.geometry.InterpolationRange` with regression tests for linear values, bounds, clamping, descending axes, and degenerate input rejection. |
| `InterpolationRanges` | Fitting Utility Owner | Completed | Array/range interpolation behavior is represented by `rapid_main.geometry.InterpolationRanges` with ordered segment lookup and out-of-range behavior regression-tested. |
| `frmTransverseProbeAutoPosition` | Transverse Mechanics Owner | Completed foundation | `rapid_main.transverse.plan_transverse_auto_position` maps scan results to a validated target/move plan with tolerance and confidence checks. Motor execution and physical probe acceptance remain open. |
| `AngleVsField_Point` | Transverse Mechanics Owner | Completed | Transverse angle-vs-field point model is represented in `rapid_main.data_model.AngleVsFieldPoint` with regression coverage through collection behavior. |
| `AngleVsFieldCollection` | Transverse Mechanics Owner | Completed | Transverse curve/collection model is represented in `rapid_main.data_model.AngleVsFieldCollection` with tests for ordered points, field arrays, peak point lookup, and weighted center angle. |
| `ProbeAngleOptimizer` | Transverse Mechanics Owner | Completed | Optimization behavior is represented by `rapid_main.data_model.ProbeAngleOptimizer`; it produces auditable peak-field angle results from `AngleVsFieldCollection` data. Hardware auto-positioning remains tracked by `frmTransverseProbeAutoPosition`. |
| `modListenAndLog` | Telemetry/Logging Owner | Integrated + core live adapters | `rapid_main.communication_log` maps deterministic transcript events and per-run export. Production SQUID raw traffic, vacuum command/acknowledgment/error traffic, structured AF/IRM ADwin evidence, and exact DC motor commands/replies/errors merge immutably into `communication.tsv`; current-run cursors and simulation exclusion prevent false live evidence. Remaining work is retained-auxiliary integration and physical transcript acceptance. |

## 2026-07-15 Parity Closure Tracking (evidence requirement)

| VB6 item | Hard evidence gate before production claim |
|---|---|
| AF in-shell execution | Software now proves the operator/queue path, readiness refusal, complete structured ramp request/result/error evidence, malformed-result rejection, and confirmed reset. Must still capture physical AF treatment start/stop/profile, interlock/safe-abort behavior, and completion evidence. |
| SQUID transport parity | Magnetometer read-quality normalization, readiness guards, exact raw transcripts, stale/partial reply rejection, and bounded safe whole-block recovery are software-verified in `acquisition.py`, `squid_transport.py`, `measurement_worker.py`, `tests/test_acquisition.py`, `tests/test_squid_transport.py`, and `tests/test_measurement_worker.py`. Recovery returns to zero, clears counters, backs off, and reacquires from zero-before; exhaustion/recovery failure returns no block and motion faults are not retried. Still must show logged recovery and run-safe halt on live hardware before production claim. |
| Vacuum control | Software proves fail-closed connection/acknowledgment behavior and never fabricates pressure. Must show physical pump/valve command round-trip and, if pressure gating is retained, independent pressure readback/interlock plus controlled fault-to-halt behavior. |
| IRM voltage calibration | Foundation mapped in `rapid_main/calibration.py` + `tests/test_calibration.py` for field-to-voltage fitting, voltage-limit validation, and JSON artifact persistence; settings/config and `diagnostic_services.py` consume slope/intercept/limit values in DAC planning (`tests/test_diagnostic_services.py`); Calibration Center exposes operator point entry and artifact recording (`tests/test_calibration_panel.py`). Still must show live DAC binding and non-blocking queue interaction with explicit ownership release before production claim. |
| DAC/MCC integration (`frmDAC_Comm`/`modMCC`) | Stable command surface, startup safe-zero reports, and adapter error surfacing are mapped in `rapid_main/dac.py` + `tests/test_dac.py`, and IRM/ARM field-to-voltage planning routes through explicit DAC commands in `diagnostic_services.py` + `tests/test_diagnostic_services.py`; still must show production adapter binding or explicit retirement rationale plus loopback evidence before production claim. |
| Thermal (`modThermal`) | Planning, safety limits, queue-label conversion, and JSON artifacts are covered by `rapid_main/thermal.py`, Calibration Center, and focused tests. `hardware_contracts.py` + `measurement_worker.py` now validate the complete treatment plan before hardware preflight, block thermal execution without a production adapter, reject malformed/unsupported live treatment labels, and preserve the last accepted treatment context after failures. Still must provide live in-app furnace/oven integration with readback/interlocks and acceptance evidence, or an explicit retirement rationale with risk controls, before production claim. |
| Rockmag (`frmRockmagRoutine`, `RockmagStep`, `RockmagSteps`) | Model mapping, compiler compatibility, operator preset integration, and deterministic routine planning artifacts are now represented by `rapid_main/data_model.py`, `rapid_main/rockmag.py`, `rapid_main/panels/sequence.py`, `tests/test_queue_and_bundle.py`, `tests/test_rockmag_routine.py`, and `tests/test_sequence_rockmag.py`. Still must show at least one reproducible hardware/run bundle for rockmag sequence types before production claim. |
| VRM (`frmVRM`) | `rapid_main` now blocks conflicting launch, closes retained main-app SQUID clients, reserves ownership until the child exits, and atomically writes a validated `rapidpy.vrm.launch_context.v1` containing available provenance. `vrm_logger` finalizes a unique immutable `rapidpy.vrm.session_manifest.v2` only after CSV closure with outcome/error, rows, write mode, timing, port, live/simulated provenance, baseline/calibration, byte size, SHA-256, and explicit handoff/run-association status (`tests/test_vrm_integration.py`, `tests/test_diagnostic_services.py`, `tests/test_queue_hardware_backend.py`). Still must show a live logged session, timing/drift/fault behavior, and physical production-run association before a production claim. |
| Plotting/analysis (`frmPlots*`, `modDataAnalysis`, `modListenAndLog`) | `MeasurementPanel` feeds completed steps to `PlotsDialog`, which renders Zijderveld/equal-area/intensity plus shared North/East/Down PCA and decay evidence, provides atomic provenance-marked JSON/CSV, and writes an isolated indexed `quicklook.json`; production SQUID, vacuum, AF/IRM ADwin, and DC motor paths contribute live evidence to per-run `communication.tsv`. Remaining gates are real hardware-run bundle review, any specifically identified legacy-only plot surfaces, and raw logging for retained auxiliary adapters. |
| Interrupted-run recovery | Operator-selectable Resume/Re-run/Skip/Abort recovery actions now exist for restored interrupted rows (`tests/test_queue_orchestration.py`, `tests/test_sample_queue_helpers.py`), and measurement runs write `workflow_summary.json` phase evidence (`tests/test_measurement_worker.py`); live safe-state checks after hardware interruption remain required. |

## Evidence index (first-party)

| Artifact | Purpose |
|---|---|
| `RapidPy/rapid_main/rapid_main/measurement_worker.py` | Sequence engine + error/abort behavior |
| `RapidPy/rapid_main/rapid_main/measurement_worker.py::artifact_index.json` | Per-run artifact manifest recording required, optional, present, and missing measurement bundle outputs |
| `RapidPy/rapid_main/rapid_main/panels/measurement.py` | Live measurement operator controls |
| `RapidPy/rapid_main/rapid_main/hardware_contracts.py` | Backend contract and preflight abstraction |
| `RapidPy/rapid_main/rapid_main/queue_compiler.py` | Queue command compiler from sample metadata |
| `RapidPy/rapid_main/tests/test_queue_and_bundle.py` | Contract test coverage for bundle + queue compile path |
| `RapidPy/rapid_main/rapid_main/io/measurement_bundle.py` | Multi-format output atomicity by step |
| `RapidPy/rapid_main/rapid_main/analysis.py` | `modDataAnalysis`-mapped vector extraction, PCA-style line-fit evidence, orientation/variance metrics, and moment statistics/decay helpers |
| `RapidPy/rapid_main/tests/test_analysis.py` | Focused regression coverage for analysis primitives |
| `RapidPy/rapid_main/rapid_main/communication_log.py` | `modListenAndLog`-mapped communication transcript, payload normalization, file export, and serial-like logging bridge |
| `RapidPy/rapid_main/rapid_main/squid_transport.py` | Exact production SQUID clear/reset, range, latch, counter, DVM, reply, and error transcript events with immutable snapshots |
| `RapidPy/rapid_main/tests/test_communication_log.py` | Focused regression coverage for communication transcript and bridge behavior |
| `RapidPy/rapid_main/rapid_main/calibration.py` | SQUID baseline calibration and `frmIRM_VoltageCalibration`-mapped IRM field-to-voltage fit/request validation primitives |
| `RapidPy/rapid_main/tests/test_calibration.py` | Focused regression coverage for IRM voltage calibration fit and validation behavior |
| `RapidPy/rapid_main/rapid_main/rockmag.py` | `frmRockmagRoutine`-mapped routine specs/templates and compiler to rockmag steps, measurement blocks, and runner labels |
| `RapidPy/rapid_main/tests/test_rockmag_routine.py` | Focused regression coverage for rockmag routine compilation and template behavior |
| `RapidPy/rapid_main/tests/test_sequence_rockmag.py` | Regression coverage that the Sequence panel "Rockmag the Works" preset uses the rockmag compiler and falls back to manual controls after edits |
| `RapidPy/rapid_main/rapid_main/transverse.py` | `frmTransverseProbeAutoPosition`-mapped auto-position planning from angle/field scans to validated move commands |
| `RapidPy/rapid_main/tests/test_transverse.py` | Focused regression coverage for transverse target selection, shortest move, tolerance, and confidence blocking |
| `RapidPy/rapid_main/rapid_main/dac.py` | `frmDAC_Comm`-mapped DAC channel setup, voltage command validation, safe-zero commands, and adapter execution shim |
| `RapidPy/rapid_main/tests/test_dac.py` | Focused regression coverage for DAC command planning, blocking, safe-zero, and adapter surface validation |
| `RapidPy/rapid_main/rapid_main/magnetometer.py` | `modMagnetometer`-mapped raw SQUID voltage calibration, background subtraction, moment magnitude, and read-quality flags |
| `RapidPy/rapid_main/tests/test_magnetometer.py` | Focused regression coverage for magnetometer calibration and quality-flag behavior |
| `RapidPy/rapid_main/rapid_main/thermal.py` | `modThermal`-mapped treatment planning, label parsing, safety limits, timing estimates, and queue block conversion |
| `RapidPy/rapid_main/tests/test_thermal.py` | Focused regression coverage for thermal routine planning and safety validation |
| `RapidPy/rapid_main/tests/test_calibration_panel.py` | Focused regression coverage for Calibration Center IRM voltage and thermal planning artifact paths |

## Ownership and status tracking

This matrix is designed to be updated as each cell transitions toward "Integrated":
- update `Migration Status`
- add evidence rows under `Evidence`
- track validation artifacts under `Notes` until operator review is complete.

Current roadmap target is to carry the existing rockmag compiler through
truthful run-bundle execution, then complete automated susceptibility mechanics
and collect physical acceptance evidence for the now-auditable SQUID, vacuum,
AF/IRM, DC motor, VRM, and retained auxiliary paths.

