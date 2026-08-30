# RAPID Main Completion Patch Plan

Date: 2026-07-13

## Goal (active objective)

Deliver `rapid_main` as the practical VB6 replacement for:

- stable automated queue execution and recovery,
- complete rapid sample handling workflows,
- reliable UI at baseline screen sizes,
- explicit app identity (non-default window/taskbar icons),
- traceable VB6 parity evidence.

## Current phase status (actual, as of 2026-07-13)

- Overall: **~85%**
- Phase 1 (sequence/layout stabilization): **completed**
- Phase 1a (UI evidence checks): **completed**
- Phase 2 (shared styling baseline): **completed**
- Phase 3 (hardware adapters and non-placeholder modules): **completed**
- Phase 4 (queue safety + recovery): **completed**
- Phase 5 (icon identity): **queued**
- Phase 6 (evidence lock and parity tracker): **queued**

## Concrete patch plan (next implementation run)

### Blocked-by-gap summary before we continue

- Queue safety: deterministic terminal state and lease release are in place; remaining work is final visual/icon parity and end-to-end hardware evidence.
- Diagnostics: dialog-level adapter contracts are wired and covered by tests; remaining work is deeper bench acceptance evidence for live controllers.
- Icons: many companion entry points still showed default Python/taskbar behavior in manual checks and need explicit visual verification after icon wiring updates.

### Plan by commit (recommended)

1. `fix: harden queue safety and recovery`
   - `RapidPy/rapid_main/rapid_main/app.py`
   - `RapidPy/rapid_main/rapid_main/measurement_worker.py`
   - `RapidPy/rapid_main/tests/test_queue_orchestration.py`
   - `RapidPy/rapid_main/tests/test_queue_and_bundle.py`
   - Scope: finalize queue terminal states, halt/restart transitions, and ownership lease release in all stop/abort paths.
   - Checks:
     - `cd RapidPy/rapid_main`
     - `cd RapidPy/rapid_main; $env:PYTHONPATH='E:/Github/RAPID/RapidPy;E:/Github/RAPID'; python -m unittest tests/test_queue_orchestration.py -v`
     - `cd RapidPy/rapid_main; $env:PYTHONPATH='E:/Github/RAPID/RapidPy;E:/Github/RAPID'; python -m unittest tests/test_measurement_worker.py -v`
     - `cd RapidPy/rapid_main; $env:PYTHONPATH='E:/Github/RAPID/RapidPy;E:/Github/RAPID'; python -m unittest tests/test_queue_and_bundle.py -v`

2. `feat: complete diagnostic adapters and error surfacing`
   - `RapidPy/rapid_main/rapid_main/hardware_contracts.py`
   - `RapidPy/rapid_main/rapid_main/diagnostic_services.py`
   - `RapidPy/rapid_main/rapid_main/dialogs/{dc_motors.py,irm_arm.py,vacuum.py,squid_comm.py}`
   - Scope: close remaining production transport/ownership gaps and align diagnostics with queue/measurement preflight failures.
   - Checks:
     - queue-safe preflight tests and negative-path adapter tests in `RapidPy/rapid_main/tests/`

3. `ui: enforce app icon identity`
   - `RapidPy/rapid_main/main.py`, `RapidPy/rapid_main/rapid_main/app.py`
   - Companion shell entry points listed in this file’s phase-5 scope.
   - Scope: add unique rapid_main shell icon, ensure top-frame + taskbar identity for rapid_main and remaining apps.
   - Checks:
     - launch checklist in staging environment for all edited entry points (manual QA), verifying no default Python icon.

4. `docs: lock phase evidence and parity`
   - `docs/vb6_parity_inventory.md`
   - `docs/rapid_main_completion_patch_plan_2026-07-13.md`
   - `docs/roadmap_execution_notes_2026-07-13.md`
   - Scope: convert remaining rows from “not assessed/in-progress” to pass/fail evidence or explicit retirement notes.

## Concrete execution board (ordered by commit)

- **Goal:** close `rapid_main` and related module gaps against VB6 for automated sample handling and hardware control.
- **Constraint:** one commit per phase, with test + evidence before proceeding.

1. `feat: rapid_main sequence and sidebar layout stabilization`
   - Scope: `RapidPy/rapid_main/rapid_main/panels/sequence.py`, `RapidPy/rapid_main/rapid_main/app.py`, `RapidPy/rapid_main/tests/test_sequence_ui_smoke.py`
   - Checks: `python -m py_compile RapidPy/rapid_main/rapid_main/app.py RapidPy/rapid_main/rapid_main/panels/sequence.py`; `python -m unittest RapidPy/rapid_main/tests/test_sequence_ui_smoke.py -v`
   - Exit: sequence values visible and side menu icon/text no longer clipped.

2. `ui: apply rapid_main spacing baseline`
   - Scope: shell/panel/dialog root spacing + consistency pass.
   - Checks: `python -m unittest RapidPy/rapid_main/tests/test_sequence_ui_smoke.py -v`
   - Exit: no visual clipping regressions from spacing changes.

3. `feat: complete rapid_main diagnostic adapter paths`
   - Scope: `RapidPy/rapid_main/rapid_main/hardware_contracts.py`, `RapidPy/rapid_main/rapid_main/diagnostic_services.py`, `RapidPy/rapid_main/rapid_main/dialogs/{dc_motors.py,irm_arm.py,vacuum.py,squid_comm.py}`
   - Checks: `python -m unittest RapidPy/rapid_main/tests/test_queue_and_bundle.py -v`; `python -m unittest RapidPy/rapid_main/tests/test_measurement_worker.py -v`
   - Exit: adapter contract paths wired, ownership/error surfacing defined.

4. `fix: harden queue safety and recovery`
   - Scope: `RapidPy/rapid_main/rapid_main/app.py`, `RapidPy/rapid_main/rapid_main/panels/sample_queue.py`, `RapidPy/rapid_main/rapid_main/measurement_worker.py`, `RapidPy/rapid_main/rapid_main/queue_compiler.py`
   - Checks: `python -m unittest RapidPy/rapid_main/tests/test_queue_orchestration.py -v`; `python -m unittest RapidPy/rapid_main/tests/test_queue_and_bundle.py -v`
   - Exit: deterministic `running/paused/halted/error/complete` progression with lease release guarantees.

5. `ui: enforce app icon identity`
   - Scope: `RapidPy/rapid_main/main.py`, `RapidPy/rapid_main/rapid_main/app.py`, and listed companion app entrypoints.
   - Checks: launch list app windows and confirm non-default taskbar icon (manual).
   - Exit: no listed app uses default Python icon.

6. `docs: lock phase evidence and parity status`
   - Scope: `docs/vb6_parity_inventory.md`, this patch plan, and execution notes.
   - Checks: date-stamped evidence entries for phases 1-5 and status closure.
   - Exit: parity rows are no longer ambiguous for touched VB6-mapped features.

## Execution rules

- One independent commit per phase.
- No next phase starts until pass/fail checks for the previous are added.
- Every phase must include:
  - test command + result,
  - date-stamped evidence note,
  - parity status update for touched capabilities.
- No placeholder behavior in queue/measurement/hardware paths is allowed unless explicitly marked as deferred.

## Phase 1 — sequence panel + side menu stabilization

### Why now

Operator flow is currently blocked by clipped values and visible overflow.

### Files

- `RapidPy/rapid_main/rapid_main/panels/sequence.py`
- `RapidPy/rapid_main/rapid_main/app.py`
- `RapidPy/rapid_main/tests/test_sequence_ui_smoke.py`

### Tasks

1. Reflow sequence presets and measurement card rows so values are readable at 1280x720.
2. Ensure all log-step-factor/toggle values are visible at normal text scales.
3. Expand side-menu width (default + minimum) for icon+label readability.
4. Persist/resurrect side menu width in existing layout settings flow.
5. Remove unsupported `QPushButton.setWordWrap` usage.
6. Add/extend smoke check for clipping and scrollbar expectations.

### Exit checks

- `python -m unittest RapidPy/rapid_main/tests/test_sequence_ui_smoke.py`
- Manual checks at 1280x720 and 1920x1080: no mandatory sequence horizontal scrollbar.

### Commit

`feat: rapid_main sequence and sidebar layout stabilization`

## Phase 2 — rapid_main UI spacing baseline

### Why now

Current spacing variance across panels can reintroduce clipping after phase-1 layout changes.

### Files

- `RapidPy/rapid_main/rapid_main/app.py`
- `RapidPy/rapid_main/rapid_main/panels/*.py`
- `RapidPy/rapid_main/rapid_main/dialogs/*.py`

### Tasks

1. Normalize margins, spacing, and control density on shell + panel roots.
2. Keep behavior unchanged; only layout ergonomics/consistency.
3. Re-run smoke checks after styling changes.

### Exit checks

- `python -m unittest RapidPy/rapid_main/tests/test_sequence_ui_smoke.py`
- Baseline visual verification: no clipping, no overlap.

### Commit

`ui: apply rapid_main spacing baseline`

## Phase 3 — hardware adapter closure (no placeholder controls)

### Why now

DC motors, IRM/ARM, vacuum, and SQUID diagnostics should run through explicit adapter contracts.

### Files

- `RapidPy/rapid_main/rapid_main/hardware_contracts.py`
- `RapidPy/rapid_main/rapid_main/diagnostic_services.py`
- `RapidPy/rapid_main/rapid_main/dialogs/dc_motors.py`
- `RapidPy/rapid_main/rapid_main/dialogs/irm_arm.py`
- `RapidPy/rapid_main/rapid_main/dialogs/vacuum.py`
- `RapidPy/rapid_main/rapid_main/dialogs/squid_comm.py`
- `RapidPy/rapid_main/tests/test_diagnostic_services.py`
- `RapidPy/rapid_main/tests/test_measurement_worker.py`

### Tasks

1. Confirm all four modules are using adapter paths when configured and explicit no-comm fallback otherwise.
2. Ensure ownership-safe command routing and explicit error surfacing.
3. Tighten hardware availability checks so preflight can run before hard availability rejection.

### Exit checks

- `python -m unittest RapidPy/rapid_main/tests/test_diagnostic_services.py`
- `python -m unittest RapidPy/rapid_main/tests/test_measurement_worker.py`

### Commit

`feat: complete rapid_main diagnostic adapter paths`

## Phase 4 — queue safety and recovery hardening

### Why now

Automated runs need deterministic halt/abort, phase visibility, and safe restart behavior.

### Files

- `RapidPy/rapid_main/rapid_main/app.py`
- `RapidPy/rapid_main/rapid_main/panels/sample_queue.py`
- `RapidPy/rapid_main/rapid_main/measurement_worker.py`
- `RapidPy/rapid_main/rapid_main/queue_compiler.py`
- `RapidPy/rapid_main/tests/test_queue_orchestration.py`
- `RapidPy/rapid_main/tests/test_queue_and_bundle.py`

### Tasks

1. Validate full terminal state model including `halted`, `error`, `complete`.
2. Harden pause/halt behavior so queue state is consistent after interruptions.
3. Ensure queue recovery can continue safely from stable point.
4. Improve queue compiler guards for invalid transitions.

### Exit checks

- `python -m unittest RapidPy/rapid_main/tests/test_queue_orchestration.py`
- `python -m unittest RapidPy/rapid_main/tests/test_queue_and_bundle.py`

### Commit

`fix: harden queue safety and recovery`

## Phase 5 — app/taskbar icon identity

### Why now

Operators must identify windows and avoid default Python icon ambiguity.

### Files

- `RapidPy/rapid_main/main.py`
- `RapidPy/rapid_main/rapid_main/app.py`
- `RapidPy/rapid_main/assets/rapid_main_window_icon.png`
- `RapidPy/dc_motor_control/main.py`
- `RapidPy/system_shell/main.py`
- `RapidPy/updown_control/main.py`
- `RapidPy/vrm_logger/main.py`
- `RapidPy/webcam_viewer/main.py`
- `RapidPy/data_viewer/main.py`

### Tasks

1. Ensure rapid_main icon is unique (not the COM-port icon).
2. Apply icon in both top frame and taskbar paths for each listed app.
3. Confirm startup entry points reference explicit icons.

### Exit checks

- Launch each listed app and verify non-default icon in window chrome and taskbar.

### Commit

`ui: enforce app icon identity`

## Phase 6 — evidence + parity tracker closure

### Why now

Work only counts when parity and completion evidence are captured.

### Files

- `docs/vb6_parity_inventory.md`
- `docs/rapid_main_completion_patch_plan_2026-07-13.md`
- `docs/roadmap_execution_notes_2026-07-13.md`

### Tasks

1. Update parity status for touched features and sample handling logic.
2. Add a dated evidence note per completed phase.
3. Record remaining blockers with owners and next owner action.

### Exit checks

- No placeholder/unknown statuses for feature rows touched in phases 1-5.

### Commit

`docs: lock phase evidence and parity status`

## Active goal commit order

1. `feat: rapid_main sequence and sidebar layout stabilization`
2. `ui: apply rapid_main spacing baseline`
3. `feat: complete rapid_main diagnostic adapter paths`
4. `fix: harden queue safety and recovery`
5. `ui: enforce app icon identity`
6. `docs: lock phase evidence and parity status`

## Completion criteria for this overall goal

- Baseline sequence UI no longer requires scrolling for core control values.
- Automated sample queue run/recover/interrupt behavior is validated.
- Hardware diagnostics are non-placeholder in operator paths.
- App icons are consistently explicit across rapid_main and companion shipped apps.
- Parity tracker reflects the current implementation state with evidence.

## Concrete execution goal (active)

### Objective (set as active goal)

Deliver the VB6-equivalent `rapid_main` control loop for:

1. stable, readable sequence and diagnostics UI at 1280×720+,
2. non-placeholder hardware adapter paths for motors, vacuum, IRM/ARM, and SQUID,
3. deterministic queue preflight/load/treat/measure/save/halt/recovery behavior equivalent to VB6 automation,
4. explicit application identity in both title bar and taskbar.

### Patch sequence and acceptance gates

All changes are expected to land as clean, bounded commits in this exact order:

#### Commit 1 — `feat: rapid_main sequence and sidebar layout stabilization`

- Scope:
  - `RapidPy/rapid_main/rapid_main/panels/sequence.py`
  - `RapidPy/rapid_main/rapid_main/app.py`
  - `RapidPy/rapid_main/tests/test_sequence_ui_smoke.py`
- Required outcomes:
  - sequence presets and measurement card values fully visible at 1280×720 and 1920×1080,
  - side menu expanded so all expanded icon+label items fit without clipping,
  - no runtime AttributeError from unsupported Qt methods (`setWordWrap` removed from unsupported widget types),
  - no mandatory horizontal scrollbars in sequence panel at default target sizes.
- Local checks:
  - `python -m py_compile RapidPy/rapid_main/rapid_main/app.py RapidPy/rapid_main/rapid_main/panels/sequence.py`
  - sequence visual smoke test command used in this repo (same command currently documented in roadmap notes).
- Status target: done before any Phase 2 changes.

#### Commit 2 — `ui: apply rapid_main spacing baseline`

- Scope:
  - `RapidPy/rapid_main/rapid_main/app.py`
  - `RapidPy/rapid_main/rapid_main/panels/*.py`
  - `RapidPy/rapid_main/rapid_main/dialogs/*.py`
- Required outcomes:
  - consistent spacing, margins, and control density across main shell and detail panels,
  - no new clipping introduced while retaining current behavior.
- Local checks:
  - targeted manual visual checks on 1280×720 and 1920×1080.

#### Commit 3 — `feat: complete rapid_main diagnostic adapter paths`

- Scope:
  - `RapidPy/rapid_main/rapid_main/hardware_contracts.py`
  - `RapidPy/rapid_main/rapid_main/diagnostic_services.py`
  - `RapidPy/rapid_main/rapid_main/dialogs/dc_motors.py`
  - `RapidPy/rapid_main/rapid_main/dialogs/irm_arm.py`
  - `RapidPy/rapid_main/rapid_main/dialogs/vacuum.py`
  - `RapidPy/rapid_main/rapid_main/dialogs/squid_comm.py`
- Required outcomes:
  - each dialog uses adapter path when available and explicit no-comm path when not,
  - ownership lock applied consistently; ownership errors are surfaced in UI,
  - preflight guard semantics are deterministic before starting motions/treatments.
- Local checks:
  - `python -m unittest RapidPy/rapid_main/tests/test_diagnostic_services.py`
  - `python -m unittest RapidPy/rapid_main/tests/test_measurement_worker.py`

#### Commit 4 — `fix: harden queue safety and recovery`

- Scope:
  - `RapidPy/rapid_main/rapid_main/app.py`
  - `RapidPy/rapid_main/rapid_main/panels/sample_queue.py`
  - `RapidPy/rapid_main/rapid_main/measurement_worker.py`
  - `RapidPy/rapid_main/rapid_main/queue_compiler.py`
  - queue tests in `RapidPy/rapid_main/tests/`
- Required outcomes:
  - deterministic state transitions for run lifecycle (`preflight`, `running`, `halted`, `error`, `complete`),
  - safe pause/halt behavior that leaves queue/sample status machine consistent,
  - recovery path validated for interrupted runs and invalid transitions are rejected early.
- Local checks:
  - `python -m unittest RapidPy/rapid_main/tests/test_queue_orchestration.py`
  - `python -m unittest RapidPy/rapid_main/tests/test_queue_and_bundle.py`

#### Commit 5 — `ui: enforce app icon identity`

- Scope:
  - `RapidPy/rapid_main/main.py`
  - `RapidPy/rapid_main/rapid_main/app.py`
  - `RapidPy/rapid_main/assets/rapid_main_window_icon.png`
  - `RapidPy/dc_motor_control/main.py`
  - `RapidPy/system_shell/main.py`
  - `RapidPy/updown_control/main.py`
  - `RapidPy/vrm_logger/main.py`
  - `RapidPy/webcam_viewer/main.py`
  - `RapidPy/data_viewer/main.py`
- Required outcomes:
  - custom icon is loaded by app entrypoints and applied to top-level window and taskbar,
  - rapid_main icon is not the COM-port default.
- Local checks:
  - launch checklist across all listed executables and verify non-default icon visibility.

#### Commit 6 — `docs: lock phase evidence and parity status`

- Scope:
  - `docs/vb6_parity_inventory.md`
  - `docs/rapid_main_completion_patch_plan_2026-07-13.md`
  - `docs/roadmap_execution_notes_2026-07-13.md`
- Required outcomes:
  - no gap left as “placeholder” for any item touched in commits 1–5,
  - dated evidence entries attached to each completed phase,
  - completion status and remaining blockers explicit for rapid_main and related modules.
