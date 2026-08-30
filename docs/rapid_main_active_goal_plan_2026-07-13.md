# Rapid Main Concrete Patch Plan (2026-07-13)

## Active goal (set for this thread)

Deliver `RapidPy/rapid_main` to practical VB6 replacement level by completing:

- deterministic queue execution and recovery,
- non-placeholder diagnostic/adaption hardware paths,
- readable sequence UI at 1280×720+ without forced horizontal scrolling,
- distinct window + taskbar app identity across first-party apps,
- explicit VB6 parity evidence in `docs/vb6_parity_inventory.md`.

This is the current thread goal.

## Concrete execution patch plan (ordered, commit-ready)

Target outcome: complete this goal in **6 clean commits**, one commit per phase, with evidence logged in
`docs/roadmap_execution_notes_2026-07-13.md` and parity updates in `docs/vb6_parity_inventory.md`.

### Commit 1 — sequence + sidebar stabilization

- Message: `feat: rapid_main sequence and sidebar layout stabilization`
- Files:
  - `RapidPy/rapid_main/rapid_main/panels/sequence.py`
  - `RapidPy/rapid_main/rapid_main/app.py`
  - `RapidPy/rapid_main/tests/test_sequence_ui_smoke.py`
- Exit criteria:
  - presets and measurement card values fit at 1280×720 and 1920×1080,
  - `QPushButton.setWordWrap` removed (no runtime `AttributeError`),
  - no hard horizontal scrollbar in normal sequence canvas width.
- Checks:
  - `python -m py_compile RapidPy/rapid_main/rapid_main/app.py RapidPy/rapid_main/rapid_main/panels/sequence.py`
  - `cd RapidPy/rapid_main && $env:PYTHONPATH='E:/Github/RAPID/RapidPy;E:/Github/RAPID'; python -m unittest tests.test_sequence_ui_smoke -v`

### Commit 2 — spacing baseline

- Message: `ui: rapid_main spacing baseline`
- Files:
  - `RapidPy/rapid_main/rapid_main/app.py`
  - `RapidPy/rapid_main/rapid_main/panels/*.py`
  - `RapidPy/rapid_main/rapid_main/dialogs/*.py`
- Exit criteria:
  - consistent spacing across main/shell and dialogs,
  - no clipping from spacing regressions.
- Checks:
  - `cd RapidPy/rapid_main && $env:PYTHONPATH='E:/Github/RAPID/RapidPy;E:/Github/RAPID'; python -m unittest tests.test_sequence_ui_smoke -v`

### Commit 3 — hardware adapter closure

- Message: `feat: complete rapid_main diagnostic adapter paths`
- Files:
  - `RapidPy/rapid_main/rapid_main/hardware_contracts.py`
  - `RapidPy/rapid_main/rapid_main/diagnostic_services.py`
  - `RapidPy/rapid_main/rapid_main/dialogs/dc_motors.py`
  - `RapidPy/rapid_main/rapid_main/dialogs/irm_arm.py`
  - `RapidPy/rapid_main/rapid_main/dialogs/vacuum.py`
  - `RapidPy/rapid_main/rapid_main/dialogs/squid_comm.py`
- Exit criteria:
  - all four diagnostics use adapter contracts when configured,
  - no placeholder-only DC/IRM/SQUID/vacuum command path.
- Checks:
  - `cd RapidPy/rapid_main && $env:PYTHONPATH='E:/Github/RAPID/RapidPy;E:/Github/RAPID'; python -m unittest tests.test_diagnostic_services tests.test_measurement_worker -v`

### Commit 4 — queue recovery hardening

- Message: `fix: harden queue safety and recovery`
- Files:
  - `RapidPy/rapid_main/rapid_main/app.py`
  - `RapidPy/rapid_main/rapid_main/measurement_worker.py`
  - `RapidPy/rapid_main/rapid_main/panels/sample_queue.py`
  - `RapidPy/rapid_main/rapid_main/queue_compiler.py`
- Exit criteria:
  - deterministic lifecycle (`preflight`, `running`, `paused`, `halted`, `error`, `complete`)
  - safe ownership/lease release on halt/cancel/error.
- Checks:
  - `cd RapidPy/rapid_main && $env:PYTHONPATH='E:/Github/RAPID/RapidPy;E:/Github/RAPID'; python -m unittest tests.test_queue_orchestration tests.test_queue_and_bundle -v`

### Commit 5 — app/taskbar icon identity

- Message: `ui: enforce app icon identity`
- Files:
  - `RapidPy/rapid_main/main.py`
  - `RapidPy/rapid_main/rapid_main/app.py`
  - `RapidPy/system_shell/main.py`
  - `RapidPy/data_viewer/main.py`
  - `RapidPy/updown_control/main.py`
  - `RapidPy/vrm_logger/main.py`
  - `RapidPy/webcam_viewer/main.py`
  - `RapidPy/dc_motor_control/main.py`
  - `RapidPy/com_port_mapper/main.py`
  - `RapidPy/adwin_comms/main.py`
  - `RapidPy/gaussmeter_control/main.py`
  - `RapidPy/af_clip_test/main.py`
- Exit criteria:
  - custom icon on window title bar and taskbar for all listed apps,
  - no first-party app ships with default Python taskbar icon.
- Checks:
  - manual Windows launch verification for each entrypoint (taskbar + frame).

### Commit 6 — parity evidence lock

- Message: `docs: lock phase evidence and parity status`
- Files:
  - `docs/vb6_parity_inventory.md`
  - `docs/rapid_main_completion_patch_plan_2026-07-13.md`
  - `docs/roadmap_execution_notes_2026-07-13.md`
- Exit criteria:
  - no "placeholder" ambiguity on touched VB6 rows,
  - explicit date-stamped evidence notes for each completed commit.

## Remaining parity closure sub-phases (to follow Commit 6 if we continue this stream)

- 7a AF in-shell execution parity
- 7b Thermal parity
- 7c VRM integration parity
- 7d Rockmag parity
- 7e Status/logging/maths parity

## Execution contract

- One independent commit per phase.
- A phase starts only after the prior phase test checks and evidence are logged.
- Each phase requires a dated note in `docs/roadmap_execution_notes_2026-07-13.md`.
- Keep behavior stable unless the phase objective explicitly changes it.

## Patch board (strict order)

### 1) `feat: rapid_main sequence and sidebar layout stabilization` — completed

Scope:

- `RapidPy/rapid_main/rapid_main/panels/sequence.py`
- `RapidPy/rapid_main/rapid_main/app.py`
- `RapidPy/rapid_main/tests/test_sequence_ui_smoke.py`

Deliverables:

- remove `QPushButton.setWordWrap` (caused `AttributeError`),
- reorder/reflow preset + measurement controls so full toggle values are visible,
- expand side panel width to avoid icon/text clipping,
- eliminate mandatory horizontal scrollbar for normal sequence views.

Checks:

- `python -m py_compile RapidPy/rapid_main/rapid_main/app.py RapidPy/rapid_main/rapid_main/panels/sequence.py`
- `python -m unittest RapidPy/rapid_main/tests/test_sequence_ui_smoke.py -v`

### 2) `ui: rapid_main spacing baseline` — completed

Scope:

- `RapidPy/rapid_main/rapid_main/app.py`
- `RapidPy/rapid_main/rapid_main/panels/*.py`
- `RapidPy/rapid_main/rapid_main/dialogs/*.py`

Deliverables:

- normalize spacing/margins/control heights after sequence update,
- preserve behavior while preventing reintroduced clipping.

Checks:

- `python -m unittest RapidPy/rapid_main/tests/test_sequence_ui_smoke.py -v`
- manual checks at 1280×720 and 1920×1080.

### 3) `feat: complete rapid_main diagnostic adapter paths` — completed

Scope:

- `RapidPy/rapid_main/rapid_main/hardware_contracts.py`
- `RapidPy/rapid_main/rapid_main/diagnostic_services.py`
- `RapidPy/rapid_main/rapid_main/dialogs/{dc_motors.py,irm_arm.py,vacuum.py,squid_comm.py}`
- `RapidPy/rapid_main/tests/test_diagnostic_services.py`

Deliverables:

- adapter-backed SQUID / vacuum / IRM-ARM / DC diagnostic execution,
- explicit no-comm fallback only when configured,
- ownership-safe launcher lifecycle and clear failure surfacing.

Checks:

- `python -m unittest RapidPy/rapid_main/tests/test_diagnostic_services.py -v`
- `python -m unittest RapidPy/rapid_main/tests/test_measurement_worker.py -v`

### 4) `fix: harden queue safety and recovery` — completed

Scope:

- `RapidPy/rapid_main/rapid_main/app.py`
- `RapidPy/rapid_main/rapid_main/measurement_worker.py`
- `RapidPy/rapid_main/rapid_main/panels/sample_queue.py`
- `RapidPy/rapid_main/rapid_main/queue_compiler.py`
- `RapidPy/rapid_main/tests/test_queue_orchestration.py`
- `RapidPy/rapid_main/tests/test_queue_and_bundle.py`

Deliverables:

- deterministic lifecycle (`preflight`, `running`, `paused`, `halted`, `error`, `complete`, `canceled`),
- safe halt/cancel behavior for queue + active worker + lease state,
- persistence/recovery behavior for interrupted sessions.

Checks:

- `python -m unittest RapidPy/rapid_main/tests/test_queue_orchestration.py -v`
- `python -m unittest RapidPy/rapid_main/tests/test_queue_and_bundle.py -v`

### 5) `ui: enforce app icon identity` — in progress

Scope:

- `RapidPy/rapid_main/main.py`
- `RapidPy/rapid_main/rapid_main/app.py`
- `RapidPy/system_shell/main.py`
- `RapidPy/data_viewer/main.py`
- `RapidPy/updown_control/main.py`
- `RapidPy/vrm_logger/main.py`
- `RapidPy/webcam_viewer/main.py`
- `RapidPy/dc_motor_control/main.py`
- `RapidPy/com_port_mapper/main.py`
- `RapidPy/adwin_comms/main.py`
- `RapidPy/gaussmeter_control/main.py`
- `RapidPy/af_clip_test/main.py`
- (and any newly added first-party app entrypoints)

Deliverables:

- dedicated `rapid_main` icon, distinct from COM-port/app defaults,
- top-frame + taskbar icon applied at startup,
- quick pass/fail audit for default Python icon fallback.

Checks:

- `python -m py_compile` for touched entrypoint modules,
- manual Windows launch checklist confirms branded icon for title bar and taskbar.

### 6) `feat: replace DC motors module placeholder in rapid_main`

Scope:

- `RapidPy/rapid_main/rapid_main/dialogs/dc_motors.py`
- `RapidPy/rapid_main/rapid_main/app.py`
- related adapter/ownership tests in `RapidPy/rapid_main/tests/`

Deliverables:

- remove remaining placeholder-only behavior for DC module,
- route DC actions through adapter contract with ownership locks and clear busy/error states.

Checks:

- `python -m unittest RapidPy/rapid_main/tests/test_diagnostic_services.py -v`
- ownership test proving launcher path rejects conflicts and recovers cleanly.

### 7) `feat: remaining VB6 parity closures` (sub-phases)

This phase is split to keep commits small and reviewable.

- **7a AF execution parity:** in-shell AF queue workflow/state execution replacing demo-only launch.
  - add/extend targeted tests + queue runtime coverage.
- **7b Thermal parity:** port `modThermal` flow or formally retire it with approved rationale.
- **7c VRM parity:** deterministic shell-triggered VRM invocation + output/log handoff.
- **7d Rockmag parity:** add `RockmagStep`/`RockmagSteps` execution model into queue workflow.
- **7e Status/logging/maths parity:** close `modStatusCode`, `modListenAndLog`, `modDataAnalysis`, and status contract gaps.

Each sub-phase is one commit + regression tests + parity-table update.

### 8) `docs: lock phase evidence and parity status` — queued

Scope:

- `docs/vb6_parity_inventory.md`
- `docs/rapid_main_completion_patch_plan_2026-07-13.md`
- `docs/roadmap_execution_notes_2026-07-13.md`

Deliverables:

- remove ambiguity in touched VB6 migration rows,
- every completed sub-task has explicit evidence references,
- remaining blockers show owner + required next action.

Checks:

- doc review against command outputs and code diffs for each touched area.

## Immediate next execution item

Proceed with Phase 5 (`ui: enforce app icon identity`), then Phase 6, then the 7a→7e parity sub-phases, then final docs evidence lock.
