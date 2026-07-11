# Roadmap Execution Notes — 10 July 2026

## Goal

Continue execution of the RAPID modernization roadmap through Phase 2 safety and control hardening.

## What changed this run

1. Added shared device-ownership support to diagnostic launchers in `rapid_main`:
   - `IrmArmDialog` launcher now requests `irm` ownership
   - `VacuumDialog` launcher now requests `vacuum` ownership
   - `SquidCommDialog` launcher now requests `squid` ownership

2. Expanded measurement worker execution semantics:
   - Added timeout-aware execution wrappers for preflight, treatment-step, and read operations
   - Added halt-aware cancellation check during those operations
   - Added distinct timeout errors on slow hardware operations

3. Added regression tests:
   - `RapidPy/rapid_main/tests/test_measurement_worker.py`
     - preflight timeout failure
     - halt-before-start skip
     - step command timeout
     - backend unavailable check
     - preflight warning propagation
   - Existing ownership tests remain active in `RapidPy/rapid_main/tests/test_device_ownership.py`
4. Updated VB6 parity inventory evidence rows (`docs/vb6_parity_inventory.md`) to convert `frmLogin` and `modPrint` from Not assessed to mapped where evidence exists in current `rapid_main` implementation.

## Validation

Executed command:

- `python -m py_compile RapidPy/rapid_main/rapid_main/measurement_worker.py RapidPy/rapid_main/rapid_main/app.py RapidPy/rapid_main/tests/test_measurement_worker.py RapidPy/rapid_main/rapid_main/hardware_contracts.py`
- `$env:PYTHONPATH='RapidPy/rapid_main'; python -m unittest discover -s RapidPy/rapid_main/tests -p 'test_measurement_worker.py' -v`
- `$env:PYTHONPATH='RapidPy/rapid_main'; python -m unittest discover -s RapidPy/rapid_main/tests -p 'test_device_ownership.py' -v`
- `$env:PYTHONPATH='RapidPy/rapid_main'; python -m unittest discover -s RapidPy/rapid_main/tests -p 'test_queue_and_bundle.py' -v`
- `$env:PYTHONPATH='RapidPy/rapid_main'; python -m unittest discover -s RapidPy/rapid_main/tests -p 'test_measurement_worker.py' -v`
- `$env:PYTHONPATH='RapidPy/rapid_main'; python -m unittest discover -s RapidPy/rapid_main/tests`

Outcome: all commands passed.

### Additional execution step (same date)

Executed next roadmap item under Phase 3 (responsive shell):

1. Added persisted shell layout support in `rapid_main/rapid_main/app.py`:
   - window geometry persistence (`QSettings`)
   - persisted active panel and sidebar width
   - restore on startup
   - explicit `Reset Layout` action under `View`

2. Added state persistence hooks:
   - `_restore_layout_state`
   - `_save_layout_state`
   - `closeEvent` saves UI state

Validation commands re-run:

- `python -m py_compile RapidPy/rapid_main/rapid_main/app.py`
- `$env:PYTHONPATH='RapidPy/rapid_main'; python -m unittest discover -s RapidPy/rapid_main/tests -p 'test_measurement_worker.py' -v`
- `$env:PYTHONPATH='RapidPy/rapid_main'; python -m unittest discover -s RapidPy/rapid_main/tests -p 'test_queue_and_bundle.py' -v`
- `$env:PYTHONPATH='RapidPy/rapid_main'; python -m unittest discover -s RapidPy/rapid_main/tests`

Outcome: all validation passed.

### Additional execution step (same date)

Resolved a regression in `measurement.py` introduced during previous patching work and completed the next micro-deliverable from Phase 2 / 3 execution:

1. repaired `MeasurementPanel._on_start` syntax and flow wiring so the start path is deterministic again,
2. connected measurement panel start/pause/halt controls to app-level flow state (`running` / `paused` / `halted`),
3. finalized `MeasurementPanel` ownership + backend wiring:
   - acquire/release measurement lease around a run,
   - resolve backend from `MainWindow.measurement_backend()` with fallback to `NoCommBackend`,
   - emit preflight warnings to status bar,
4. added specimen context synchronization from panel helpers (`set_specimen_context`) and release of ownership on completion/error.

Validation:

- `python -m py_compile RapidPy/rapid_main/rapid_main/panels/measurement.py RapidPy/rapid_main/rapid_main/app.py RapidPy/rapid_main/rapid_main/measurement_worker.py`
- `python -m py_compile RapidPy/rapid_main/rapid_main/panels/measurement.py RapidPy/rapid_main/rapid_main/app.py RapidPy/rapid_main/rapid_main/measurement_worker.py RapidPy/rapid_main/tests/test_measurement_panel.py RapidPy/rapid_main/tests/test_measurement_worker.py`
- `$env:PYTHONPATH='RapidPy/rapid_main'; python -m unittest discover -s RapidPy/rapid_main/tests -p 'test*.py' -v`

Outcome: compile checks and full test suite passed (49 tests).

### Additional execution step (same date)

Executed a small hardening item under Phase 4 (AF measurement workflow):

1. Normalized measurement workflow phase emissions to canonical names:
   - Removed dynamic phase strings such as `set_demag_step(AF20)` and `read_squid(NRM)` from worker signals.
   - Kept canonical `WorkflowPhase` names (`preflight`, `loading`, `treating`, `measuring`, `saving`, `returning`, `complete`, `error`, etc.) so the shell receives stable, map-ready signals.
2. Updated shell workflow styles:
   - Added dedicated `flowValidating` styling so future validating phase can render with distinct visual state.
   - Wired validating phase to the dedicated style class.
3. Extended worker tests:
   - Added canonical-phase validation for the `phase_changed` signal stream in successful runs (`phase` values must be members of `WorkflowPhase`).
   - Adjusted timeout assertion text after canonicalizing step-phase messages.

Validation:

- `python -m py_compile RapidPy/rapid_main/rapid_main/measurement_worker.py RapidPy/rapid_main/rapid_main/app.py RapidPy/rapid_main/tests/test_measurement_worker.py`
- `$env:PYTHONPATH='RapidPy/rapid_main'; python -m unittest discover -s RapidPy/rapid_main/tests -p 'test*.py' -v`

Outcome: all checks passed (52 tests).

### Additional execution step (same date)

Execution item under Phase 3 + Phase 1:

1. Responsive shell hardening:
   - Migrated `rapid_main` central layout from fixed sidebar+stack widgets to a horizontal `QSplitter`.
   - Persisted splitter state with `QSettings` (`ui/main_splitter_state`) and kept legacy width compatibility for reset/restore.
   - Added explicit default-state fallback when no splitter state is present.
2. Phase 1 evidence expansion:
   - Added a remediation queue table in `vb6_parity_inventory.md` that assigns every remaining `Not assessed` VB6 item an owner and acceptance criteria.

Validation:

- `python -m py_compile RapidPy/rapid_main/rapid_main/app.py`

Outcome: layout compile checks passed and the roadmap inventory gap was expanded.

### Additional execution step (same date)

Implemented initial workflow-phase contract hardening for AF/measurement runs:

1. Added shared phase primitives in `rapid_main/rapid_main/workflow.py`:
   - `WorkflowPhase` enum covering all required high-level workflow states,
   - `WorkflowStateMachine` with guarded transitions and transition-error behavior.
2. Emitted deterministic phase transitions from `MeasurementWorker` (`preflight`, `loading`, `treating`, `measuring`, `saving`, `returning`, `complete`, with failure-to-`error`/`halted` paths).
3. Wired `MeasurementPanel` to consume worker phase transitions and update the shell status label accordingly.
4. Extended shell style/label mapping in `app.py` to render richer workflow phases in the top flow indicator.

Validation commands:

- `python -m py_compile RapidPy/rapid_main/rapid_main/workflow.py RapidPy/rapid_main/rapid_main/measurement_worker.py RapidPy/rapid_main/rapid_main/app.py RapidPy/rapid_main/rapid_main/panels/measurement.py RapidPy/rapid_main/tests/test_workflow.py RapidPy/rapid_main/tests/test_measurement_worker.py`
- `$env:PYTHONPATH='RapidPy/rapid_main'; python -m unittest discover -s RapidPy/rapid_main/tests -p 'test*.py' -v`

Outcome: all compile checks and full test suite passed (52 tests).

### Additional execution step (same date)

Executed next roadmap item to unblock AF workflow start in `rapid_main`:

1. Fixed runtime wiring for measurement start by resolving backend via the
   `measurement_backend()` window contract in `rapid_main/rapid_main/panels/measurement.py`.
2. Added structural fallback logic so missing/invalid backend providers degrade safely to
   `NoCommBackend` instead of crashing.
3. Added deterministic success-path coverage:
   - `RapidPy/rapid_main/tests/test_measurement_worker.py`
     - `test_successful_run_generates_bundle_and_steps` (preflight + 2-step run + output bundle writes + file checks)
   - `RapidPy/rapid_main/tests/test_measurement_panel.py`
     - `_resolve_measurement_backend` behavior for callable/no provider/non-callable provider cases.

Validation:

- `python -m py_compile RapidPy/rapid_main/rapid_main/panels/measurement.py RapidPy/rapid_main/tests/test_measurement_worker.py RapidPy/rapid_main/tests/test_measurement_panel.py`
- `$env:PYTHONPATH='RapidPy/rapid_main'; python -m unittest discover -s RapidPy/rapid_main/tests -p 'test*.py' -v`

Outcome: all validation passed (49 tests).

### Additional execution step (same date)

Executed queue-parity hardening for sample-order workflows:

1. Added queue validation helpers in `rapid_main/rapid_main/queue_compiler.py`:
   - `QueueValidationResult` container now captures ordered validation errors and warnings.
   - `validate_queue_samples()` validates sample metadata (empty sample names, invalid hole indices, non-positive step counts).
   - duplicate sample-hole values are surfaced as warnings for operator visibility.
   - `compile_queue()` now supports strict mode, enabling deterministic rejection of invalid queues before workflow start.
2. Added coverage in `RapidPy/rapid_main/tests/test_queue_and_bundle.py` for:
   - warning capture on duplicate hole assignment,
   - strict-mode rejection for invalid queue rows.
3. Updated `docs/vb6_parity_inventory.md` so `frmChangerSampOrder` is now `Mapped` with evidence in the compiler contract.

Validation:

- `python -m py_compile RapidPy/rapid_main/rapid_main/queue_compiler.py RapidPy/rapid_main/tests/test_queue_and_bundle.py`
- `$env:PYTHONPATH='RapidPy/rapid_main'; python -m unittest discover -s RapidPy/rapid_main/tests -p 'test_queue_and_bundle.py' -v`
- `$env:PYTHONPATH='RapidPy/rapid_main'; python -m unittest discover -s RapidPy/rapid_main/tests -v`

Outcome: all checks passed (54 tests).

### Additional execution step (same date)

Executed queue runtime integration to wire the queue panel into `MainWindow` and measurement execution:

1. Added queue automation lifecycle in `rapid_main/rapid_main/app.py`:
   - New `start_queue_run(samples, options)` entry point with strict validation and queue-state bookkeeping.
   - `_run_next_queue_sample`, `cancel_queue_run`, and `_on_queue_sample_finished` to coordinate sequential sample runs and stop conditions.
   - Queue execution now runs each sample through `MeasurementPanel.start_measurement_for_sample(sample_name)` and advances automatically.
2. Finished queue panel run controls in `rapid_main/rapid_main/panels/sample_queue.py`:
   - Added `Run Queue` button path with strict/alert validation and warning continuation prompts.
   - Added queue row status updates for `Running` / `Done` / `Error`.
   - Added parse/read helpers used by both runtime launch and test coverage.
3. Connected measurement completion signaling in `rapid_main/rapid_main/panels/measurement.py`:
   - Added `sample_run_finished` signal.
   - Measurement start helper now accepts explicit sample names and cancels active queue runs only for manual start paths.
4. Added regression coverage for queue parse helpers:
   - `RapidPy/rapid_main/tests/test_sample_queue_helpers.py`
   - Covers position parsing and treatment-step count inference.

Validation commands:

- `python -m py_compile RapidPy/rapid_main/rapid_main/app.py RapidPy/rapid_main/rapid_main/panels/measurement.py RapidPy/rapid_main/rapid_main/panels/sample_queue.py RapidPy/rapid_main/tests/test_sample_queue_helpers.py`
- `$env:PYTHONPATH='RapidPy/rapid_main'; python -m py_compile RapidPy/rapid_main/rapid_main/queue_compiler.py RapidPy/rapid_main/rapid_main/app.py RapidPy/rapid_main/rapid_main/measurement_worker.py RapidPy/rapid_main/tests/test_queue_and_bundle.py RapidPy/rapid_main/tests/test_sample_queue_helpers.py`
- `$env:PYTHONPATH='RapidPy/rapid_main'; python -m unittest discover -s RapidPy/rapid_main/tests -p 'test*.py' -v`

Outcome: all checks passed (57 tests).

### Additional execution step (same date)

Executed control-path hardening for queue/manual run safety:

1. Added shared control handlers so queue toolbar and app header pause/halt buttons can drive active measurement execution:
   - App header Pause/Halt now calls `MainWindow.toggle_queue_pause()` and `MainWindow.halt_measurement()`.
   - `SampleQueuePanel` toolbar Pause/Halt now routes to window-level handlers when available.
2. Ensured queue cancellation is safe against active workers:
   - `MainWindow.cancel_queue_run()` now halts any active measurement worker before clearing queue state.
   - `MainWindow.halt_measurement()` now performs both worker halt and queue cancel with clear user-facing status.
3. Added small public measurement panel run controls to support cross-widget control (`toggle_pause`, `halt_run`, `is_active`).

Validation:

- `python -m py_compile RapidPy/rapid_main/rapid_main/app.py RapidPy/rapid_main/rapid_main/panels/measurement.py RapidPy/rapid_main/rapid_main/panels/sample_queue.py`
- `$env:PYTHONPATH='RapidPy/rapid_main'; python -m unittest discover -s RapidPy/rapid_main/tests -p 'test*.py' -v`

Outcome: all checks passed (57 tests).

### Additional execution step (same date)

Executed first major Phase 1 + Phase 7 bridge item by migrating VB6 vector math into a shared `rapid_main` geometry service:

1. Added `RapidPy/rapid_main/rapid_main/geometry.py`, implementing:
   - Legacy-style angle helpers (`atan`, `Atan2`, arccos/arcos equivalents),
   - `Angular3D`/`Cartesian3D` dataclasses and conversion math,
   - vector arithmetic and angle utilities (`dot product`, `diff angle`, `sum/diff/average`,
   scalar/scalar square helpers).
2. Wired worker and parser conversion paths to the new service:
   - `_build_step()` in `measurement_worker.py` now derives specimen directions via
     `cartesian3d_to_angular3d`.
   - `specimen_reader.py` now resolves `sdec/sinc` back to xyz through the shared vector utility
     (with legacy viewer z-sign convention preserved).
3. Added parity evidence for this conversion layer:
   - `RapidPy/rapid_main/tests/test_geometry.py` with round-trip and vector-math checks.
4. Updated the parity ledger:
   - `docs/vb6_parity_inventory.md` rows for `modVector3d`, `Cartesian3D`, and `Angular3D`
     moved from Not-assessed toward Mapped/Completed.

Validation performed:

- `python -m py_compile RapidPy/rapid_main/rapid_main/geometry.py RapidPy/rapid_main/rapid_main/measurement_worker.py RapidPy/rapid_main/rapid_main/io/specimen_reader.py RapidPy/rapid_main/tests/test_geometry.py`
- `$env:PYTHONPATH='RapidPy/rapid_main'; python -m unittest discover -s RapidPy/rapid_main/tests -p 'test_geometry.py' -v`
- `$env:PYTHONPATH='RapidPy/rapid_main'; python -m unittest discover -s RapidPy/rapid_main/tests -p 'test*.py' -v`

Outcome: all new checks passed.

### Additional execution step (same date)

Executed the roadmap Phase 1 config-parity item by implementing VB6 legacy INI migration in `rapid_main`:

1. Added a first-class VB6 INI migration utility:
   - New `RapidPy/rapid_main/rapid_main/legacy_ini.py` with conservative key mapping into `AppConfig`.
   - Captures mapped/unmapped fields and warning details in `LegacyIniImportReport`.
   - Normalizes COM ports, applies best-effort placeholder mappings where a direct target field is not available, and reports unsupported entries.
2. Wired a discoverable workflow path in Settings:
   - Added `Import VB6 INI…` action to `SettingsPanel`.
   - `_import_vb6_ini` opens a file picker, applies import, refreshes panel bindings, and surfaces concise user feedback.
3. Added regression coverage:
   - `RapidPy/rapid_main/tests/test_legacy_ini.py` imports `VB6/settings/Paleomag_v3.INI`,
     validates mapped settings (`Program`, `COMPorts`, `MagnetometerCalibration`, `AF`, `ARM/IRM`, `Vacuum`) and warning capture.
   - Added invalid-value test to confirm safe defaulting + warnings.
4. Closed VB6 inventory gaps:
   - `docs/vb6_parity_inventory.md` moved `CIniFile` and `frmINIConverter` from Not assessed to Mapped with implementation evidence.
   - `ROADMAP.md` Step 1 execution text updated to reflect completed mapping work for those items.

Validation:

- `python -m py_compile RapidPy/rapid_main/rapid_main/legacy_ini.py RapidPy/rapid_main/rapid_main/panels/settings_panel.py`
- `$env:PYTHONPATH='RapidPy/rapid_main'; python -m unittest RapidPy/rapid_main/tests/test_legacy_ini.py -v`

Outcome: checks passed for INI migration compile/test coverage.

### Additional execution step (same date)

Executed final Step-2 hardening for service-level ownership on shell launchers:

1. Added an owned-stub dialog path for non-modal actions so ownership remains explicit even when a launcher is still a placeholder.
2. Routed the DC Motors diagnostic launcher through ownership control (`resource=dc_motors`, `owner=dc_motors_panel`) so concurrent command attempts are now serialized with a user-visible busy error.
3. Confirmed `ROADMAP.md` and `vb6_parity_inventory.md` reflect the INI migration evidence and the ownership hardening completion for this run.

Validation:

- `python -m py_compile RapidPy/rapid_main/rapid_main/app.py`

Outcome: checks passed for ownership-safe launcher semantics and documentation updates for this step.

### Additional execution step (same date)

Captured the missing Step-5 checkpoint evidence before progressing deeper into calibration automation:

1. Added `docs/operator_review_notes_2026-07-10.md` as the required operator-review artifact for this phase.
2. Recorded what is working well, outstanding risks, and explicit follow-up points for next operator validation session.
3. Updated the roadmap checklist status to mark Step 5 as implemented at the checkpoint artifact level, while keeping operator sign-off as pending.

Validation:

- Repository documentation review to confirm artifact path and roadmap references.

Outcome: Step 5 evidence capture is now present and linked for traceability before broader calibration automation work.

### Additional execution step (same date)

Hardening phase-state semantics for AF/measurement parity:

1. Integrated `WorkflowStateMachine` into `MeasurementWorker` to enforce canonical phase transitions at runtime and emit transition-controlled phase updates.
2. Added phase hooks for `positioning` and `validating` so the AF/measurement workflow now emits the broader treatment and quality pipeline phases (`treating -> positioning -> measuring -> validating -> saving`).
3. Added runtime step validation (`_validate_position`, `_validate_step`) as a contract extension point for future device checks.
4. Expanded worker tests with ordered phase-coverage checks in
   `test_phase_sequence_covers_treatment_position_and_validation`.

Validation:

- `python -m py_compile RapidPy/rapid_main/rapid_main/measurement_worker.py RapidPy/rapid_main/tests/test_measurement_worker.py`
- `python -m unittest discover -s RapidPy/rapid_main/tests -p 'test_measurement_worker.py' -v`

Outcome: stronger workflow-state contract alignment for AF/measurement run semantics, with regression coverage for phase ordering.

### Additional execution step (same date)

Executed the next roadmap-controlled hardening item in the AF/measurement contract:

1. Added backend-driven positioning validation support in `MeasurementWorker`:
   - `_validate_position()` now calls `backend.validate_position()` when present.
   - Validation supports timeout control via backend-provided `position_timeout`.
   - Backend-returned invalid position results now abort the step and emit worker-level error (`Hardware error at step ...`).
2. Added regression coverage for the new contract behavior in
   `test_measurement_worker.py`:
   - `test_position_validation_failure_prevents_run` proves an invalid position blocks the run and keeps the sample in a failed state.
3. Verified that position checks are now part of the canonical phase stream between `treating -> validating -> saving`.

Validation:

- `python -m py_compile RapidPy/rapid_main/rapid_main/measurement_worker.py RapidPy/rapid_main/tests/test_measurement_worker.py`
- `$env:PYTHONPATH='RapidPy/rapid_main'; python -m unittest RapidPy/rapid_main/tests/test_measurement_worker.py -v`
- `$env:PYTHONPATH='RapidPy/rapid_main'; python -m unittest discover -s RapidPy/rapid_main/tests -p 'test*.py' -v`

Outcome: full regression suite remains green (64 tests).

### Additional execution step (same date)

Closed the roadmap control-safety gap for VB6 shutdown parity by implementing a controlled shutdown path in `rapid_main`:

1. Added a centralized exit flow in `MainWindow`:
   - Replaced direct `QApplication.quit` calls in the file menu and header `Exit` button with `_request_shutdown`.
   - Added `_confirm_shutdown()` to coordinate operator prompts for active automation, safe-halting behavior, and final confirmation.
   - Added `_has_active_automation()` guard and `closeEvent` integration so close requests go through the same confirmation/halt path.
2. Mapped roadmap inventory evidence:
   - Updated `docs/vb6_parity_inventory.md` status for `frmShutdownMsg` to `Mapped`.
   - Added completion note in `ROADMAP.md` execution status and evidence list.

Validation:

- `python -m py_compile RapidPy/rapid_main/rapid_main/app.py`
- `$env:PYTHONPATH='RapidPy/rapid_main'; python -m unittest discover -s RapidPy/rapid_main/tests -p 'test*.py' -v`

Outcome: shutdown control path is implemented and existing full regression remains green (64 tests).

### Additional execution step (same date)

Executed the roadmap item to close the VB6 sample-index migration gap in `rapid_main`:

1. Added VB6 `.sam` registry data models in `rapid_main/rapid_main/data_model.py`:
   - `SampleIndexRegistration` (one registry row),
   - `SampleIndexRegistrations` (ordered collection).
2. Added a `.sam` parser in `rapid_main/rapid_main/io/sample_index.py` to support:
   - legacy two-line header interpretation (`sample_set` + coordinate/location line),
   - ordered specimen extraction,
   - depth/formation/location metadata where available.
3. Wired sample selection loading (`SampleSelectDialog`) to the new registry reader:
   - row view now binds depth/formation/location fields from the parsed registry when available,
   - retains existing specimen-file fallback for volume when present.
4. Added regression fixtures in `rapid_main/tests/test_sample_index.py` for:
   - header-style `.sam` files,
   - plain specimen list `.sam` files and comment/whitespace handling.

Validation:

- `python -m py_compile RapidPy/rapid_main/rapid_main/data_model.py RapidPy/rapid_main/rapid_main/io/sample_index.py RapidPy/rapid_main/rapid_main/dialogs/sample_select.py RapidPy/rapid_main/tests/test_sample_index.py`
- `$env:PYTHONPATH='RapidPy/rapid_main'; python -m unittest discover -s RapidPy/rapid_main/tests -p 'test_sample_index.py' -v`
- `$env:PYTHONPATH='RapidPy/rapid_main'; python -m unittest discover -s RapidPy/rapid_main/tests -p 'test*.py' -v`

Outcome: sample-index migration evidence is now captured (66 tests pass); the registry gap in `vb6_parity_inventory.md` was closed with a mapped status and completed remediation criteria.
### Additional execution step (same date)

Closed an AF safety-hardening gap in the core measurement contract:

1. Added a backend safe-return callback path in `MeasurementWorker`:
   - `return_to_safe_state()` is now called in a bounded path (`return_timeout`) after execution setup when output is active.
   - The worker now emits `RETURNING` on both successful and interrupted workflow exits before `COMPLETE` is emitted.
   - Any safe-return failures are surfaced to the operator and mapped to an error state while preserving run-finish signaling.
2. Extended backend defaults to keep simulator-safe behavior explicit:
   - `NoCommBackend.return_to_safe_state()` added as a no-op.
   - `return_timeout` added alongside existing timeout knobs for consistent backend configuration.
3. Added regression coverage in `RapidPy/rapid_main/tests/test_measurement_worker.py`:
   - `test_successful_run_invokes_return_to_safe_state` verifies the callback executes on success.
   - `test_return_to_safe_state_runs_after_run_failure` verifies the callback still executes after treatment/failure paths.
4. Updated execution tracker wording in `ROADMAP.md` Step 4 to capture this completion status.

Validation:

- `python -m py_compile RapidPy/rapid_main/rapid_main/measurement_worker.py RapidPy/rapid_main/rapid_main/hardware_contracts.py RapidPy/rapid_main/tests/test_measurement_worker.py`
- `$env:PYTHONPATH='RapidPy/rapid_main'; python -m unittest discover -s RapidPy/rapid_main/tests -p 'test*.py' -v`

Outcome: safe-return behavior is now verified with updated regression coverage (68 tests pass).

### Additional execution step (same date)

Executed the roadmap AF-launch completion item for operator workflow:

1. Updated the diagnostics AF menu and launcher behavior to support both planning and executable workflows:
   - Added a canonical AF preset helper, `_af_demo_labels()`, so AF sequence values are centrally defined.
   - Added `Run AF Demo Sequence` under `Diagnostics → AF Demagnetizer` for direct launch/start of an AF specimen run.
   - Kept `AF Demag Window` as the existing manual setup path for users who want to inspect or start manually.
2. Added deterministic workflow preparation helper in `rapid_main.app`:
   - `_prepare_af_workflow(auto_start: bool)` loads sequence context, sets specimen context (`AF_DEMO`), navigates to Live Measurement, and optionally starts automatically.
   - Added clear status messaging for both loaded and auto-started run states.
3. Added a roadmap regression test target for AF preset determinism:
   - `RapidPy/rapid_main/tests/test_app_af_workflow.py` validates that AF preset labels remain stable and canonical.

Validation:

- `python -m py_compile RapidPy/rapid_main/rapid_main/app.py`
- `python -m py_compile RapidPy/rapid_main/tests/test_app_af_workflow.py`
- `$env:PYTHONPATH='RapidPy/rapid_main'; python -m unittest discover -s RapidPy/rapid_main/tests -p 'test_app_af_workflow.py' -v`

Outcome: AF workflow entrypoint can now be started directly from diagnostics menu for the demo path, while preserving manual AF access.

### Additional execution step (same date)

Executed a queue-recovery hardening item for crash-safe continuation:

1. Added queue-status recovery behavior in `rapid_main/rapid_main/panels/sample_queue.py`:
   - New `recover_interrupted_samples()` helper scans persisted queue rows and resets lingering `Running` states to `Pending` before next startup.
   - This prevents stale in-flight rows from appearing as permanently active after an abrupt session stop.
2. Added queue progress persistence hooks in `rapid_main/rapid_main/app.py`:
   - Persisted queue snapshot, queue position/progress, active-row marker, and active-queue flag in `QSettings`.
   - Persisted state is updated at queue transitions (start, each sample dispatch, done/error/cancel).
   - `closeEvent` and reset layout paths now clear queue recovery keys consistently.
3. Added recovery messaging and bootstrap behavior:
   - On restore, interrupted rows are auto-normalized to pending.
   - If the prior queue session was active without stale rows, status explicitly advises re-run to continue safely.

Validation:

- `python -m py_compile RapidPy/rapid_main/rapid_main/app.py RapidPy/rapid_main/rapid_main/panels/sample_queue.py`
- `$env:PYTHONPATH='RapidPy/rapid_main'; python -m unittest discover -s RapidPy/rapid_main/tests -p 'test*.py' -v`

Outcome: queue persistence/recovery hooks are added with recovery status handling; existing test suite remains green.

### Additional execution step (same date)

Moved data review from static placeholder into an actionable workflow bridge:

1. Connected the Measurement panel "Show Stats Window" button to its existing `PlotsDialog` launch path so the review control is now usable in the UI.
2. Added a main menu **View → Data Review** action that launches the `RapidPy/data_viewer/main.py` utility in a dedicated subprocess (with running-instance guard and missing-script/error handling).

Validation:

- `python -m py_compile RapidPy/rapid_main/rapid_main/app.py`
- `python -m py_compile RapidPy/rapid_main/rapid_main/panels/measurement.py`
- `$env:PYTHONPATH='RapidPy/rapid_main'; python -m unittest discover -s RapidPy/rapid_main/tests -p 'test*.py' -v`

Outcome: review entrypoints are now active and user-launchable; compile/test checks remain green (`70` tests previously passing, expected unchanged by this wiring).

### Additional execution step (same date)

Completed the requested UI consistency sweep across standalone apps:

1. Added launcher icon assets for apps that lacked any icon resources:
   - `RapidPy/dc_motor_control/assets/dc_motor_control_icon.png` and `.ico`
   - `RapidPy/system_shell/assets/system_shell_icon.png` and `.ico`
   - `RapidPy/rapid_main/assets/rapid_icon.png` and `.ico` (aligning icon bootstrap path for `rapid_main`)
2. Hooked remaining app bootstraps into the shared icon contract:
   - `RapidPy/dc_motor_control/dc_motor_control/app.py`
   - `RapidPy/system_shell/system_shell/app.py`
3. Synced `vrm_logger` startup styling path to the shared baseline:
   - Bootstrapped common styling in `vrm_logger/app.py` main loop (`apply_liquid_glass_theme`)
   - Kept custom VRM chart/sidebar polish in `_apply_style()` as an additive overlay so existing visual behavior remains unchanged.
4. Added explicit icon bootstrap and shared theme entry in `system_shell`, `dc_motor_control`, and `vrm_logger` so standalone launch UX is visually uniform.

Validation:

- `python -m py_compile RapidPy/dc_motor_control/dc_motor_control/app.py RapidPy/system_shell/system_shell/app.py RapidPy/vrm_logger/vrm_logger/app.py RapidPy/rapid_main/rapid_main/app.py`

Outcome: all targeted standalone apps now initialize with the shared Liquid-Glass look-and-feel and icon bootstrap path; no syntax errors in edited app entrypoints.

### Additional execution step (same date)

Unblocked `rapid_main` launch import pathing and completed the last standalone app icon gap:

1. Fixed `RapidPy/rapid_main/main.py` package bootstrap so the launcher now inserts:
   - `RapidPy/rapid_main` (for `rapid_main` package imports)
   - `RapidPy` (for shared package imports like `rapidpy_common`)
   so `python RapidPy/rapid_main/main.py` imports cleanly.
2. Added a `webcam_viewer` launcher icon asset:
   - `RapidPy/webcam_viewer/assets/webcam_viewer_icon.png`
3. Normalized `webcam_viewer` theme/icon bootstrap in `RapidPy/webcam_viewer/webcam_viewer/app.py`:
   - explicit shared theme initialization path with `set_app_icon(app, "webcam_viewer_icon.png", ...)`
   - removed the previous argument-mismatch call path that skipped shared styling on import errors.

Validation:

- `python -m py_compile RapidPy/rapid_main/main.py RapidPy/rapid_main/rapid_main/app.py RapidPy/webcam_viewer/webcam_viewer/app.py`
- `python -c "import sys; from pathlib import Path; sys.path.insert(0, str(Path('RapidPy/rapid_main').resolve())); sys.path.insert(0, str(Path('RapidPy').resolve())); import rapid_main.app as a; print('rapid_main import ok')"`

Outcome: launcher/import compatibility and the final missing app icon bootstrap item are addressed while preserving UI sync direction.
