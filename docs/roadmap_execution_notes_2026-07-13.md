## RAPID Roadmap Execution Notes — 2026-07-13

- Active goal source: `docs/rapid_main_active_goal_plan_2026-07-13.md`
- Overall objective: close remaining rapid_main gaps for stable automated sample handling, UI readability, recovery safety, and identity parity.

### Current execution state

- Active objective: `RapidPy/rapid_main` plus parity-critical control modules (AF, thermal, VRM, rockmag) closure for hardware/sample handling  
  (active execution board: `docs/rapid_main_active_goal_plan_2026-07-13.md`).
- Phase 1 (sequence/layout): **completed**
- Phase 1a (evidence validation): **completed**
- Phase 2 (shared spacing baseline): **completed**
- Phase 3 (diagnostic adapter completion): **completed**
- Phase 4 (queue safety/recovery): **completed**
- Phase 5 (icon identity): **in progress**
- Phase 6 (parity evidence lock): **queued**
- Phase 7 (remaining parity gaps): **queued**

### Next actions (ordered)

1. `fix: harden queue safety and recovery`
   - scope: queue/measurement lifecycle and recovery semantics.
   - status: complete.
2. `ui: enforce app icon identity`
   - scope: rapid_main + listed companion apps (window icon + taskbar icon, non-default Python icon).
   - status: in progress.
3. `docs: lock phase evidence and parity status`
   - scope: fill all touched rows in parity tracker with phase-complete evidence references.
   - status: queued.
4. `feat: close remaining parity gaps (AF/thermal/VRM/rockmag)`
   - scope: in-shell AF execution + thermal/VRM/rockmag workflow execution and output wiring.
   - status: queued.

### Completion rule for each phase

- File diffs and tests are run for touched paths.
- Evidence note (date, command, result) added to this file.
- Parity tracker updated where applicable.

### Evidence log

- 2026-07-13: Sequence smoke check passing:
  - `cd RapidPy/rapid_main; python -m unittest tests/test_sequence_ui_smoke.py -v`
  - Result: `OK` (1 test, 2 subtests at 1280x720 and 1920x1080).
- 2026-07-13: Current sequence/layout verification:
  - `python -m py_compile RapidPy/rapid_main/rapid_main/app.py RapidPy/rapid_main/rapid_main/panels/sequence.py`
  - Result: `OK`
- 2026-07-13: Queue + worker checks after latest lock changes:
  - `cd RapidPy/rapid_main; python -m unittest tests/test_measurement_worker.py -v`
  - `cd RapidPy/rapid_main; python -m unittest tests/test_queue_orchestration.py -v`
  - `cd RapidPy/rapid_main; python -m unittest tests/test_queue_and_bundle.py -v`
  - Result: `OK` on all suites.
- 2026-07-13: Executable patch plan rewritten into concrete phase board:
  - Updated: `docs/rapid_main_active_goal_plan_2026-07-13.md`
  - Result: strict six-phase commit order defined and linked to checks/docs evidence.
- 2026-07-13: Rapid-main adapter and queue diagnostics test suite now green under app-oriented unittest baseline:
  - `cd RapidPy/rapid_main; $env:PYTHONPATH='E:/Github/RAPID/RapidPy'; python -m unittest tests/test_diagnostic_services.py -v`
  - `cd RapidPy/rapid_main; $env:PYTHONPATH='E:/Github/RAPID/RapidPy'; python -m unittest tests/test_queue_orchestration.py -v`
  - `cd RapidPy/rapid_main; $env:PYTHONPATH='E:/Github/RAPID/RapidPy'; python -m unittest tests/test_queue_and_bundle.py -v`
  - `cd RapidPy/rapid_main; $env:PYTHONPATH='E:/Github/RAPID/RapidPy'; python -m unittest tests/test_measurement_worker.py -v`
  - `cd RapidPy/rapid_main; $env:PYTHONPATH='E:/Github/RAPID/RapidPy'; python -m unittest tests/test_sequence_ui_smoke.py -v`
  - Result: `OK` (full targeted coverage; 35 tests + sequence smoke).
- 2026-07-13: app icon identity plumbing for `data_viewer` unified with shared helper:
  - `cd RapidPy; python -m py_compile data_viewer/data_viewer/app.py RapidPy/rapid_main/rapid_main/app.py` (invoked from repo root via explicit paths)
  - Result: `OK`
- 2026-07-13: App icon wiring audit confirms every entry point binds shared icon helper and referenced assets exist:
  - `python -c "<icon audit script using regex set_app_icon(...) in RapidPy/*/app.py>"`
  - Result: `OK` (`missing_icon_entrypoints=0`, `missing_icon_files=0`)
- 2026-07-13: queue safety/recovery + adapter-backed run-path tests re-ran clean with repo imports:
  - `cd RapidPy/rapid_main; $env:PYTHONPATH='E:/Github/RAPID/RapidPy'; python -m unittest tests.test_queue_orchestration tests.test_measurement_worker tests.test_queue_and_bundle tests.test_diagnostic_services tests.test_sequence_ui_smoke -v`
  - Result: `OK` (39 tests).
- 2026-07-13: diagnostics and workflow launch surface is wired in `rapid_main` menu paths for AF tuner, AF calibration, IRM calibrations, Gaussmeter, VRM, and Susceptibility/Calibrate Rod placeholders with safety guard.
  - `python -m py_compile RapidPy/rapid_main/rapid_main/app.py`
  - `cd RapidPy/rapid_main; $env:PYTHONPATH='E:/Github/RAPID/RapidPy'; python -m unittest tests.test_queue_and_bundle tests.test_diagnostic_services tests.test_queue_orchestration tests.test_measurement_worker tests.test_sequence_ui_smoke -v`
  - Result: `OK` (41->42 tests now green for full smoke/unit subset after launcher wiring).
- 2026-07-15: layout portability hardening pass applied:
  - Updated shared geometry clamping in `RapidPy/rapidpy_common/ui.py` (window-size caps).
  - Reduced rapid_main startup defaults and sidebar caps in `RapidPy/rapid_main/rapid_main/app.py`.
  - Added restore-normalization safeguards to geometry-restoring windows (`RapidPy/adwin_comms/adwin_comms/app.py`, `RapidPy/vrm_logger/vrm_logger/app.py`).
  - Extended `RapidPy/rapid_main/tests/test_window_layout.py` to cover representative display profiles and updated expected clamp outputs.
  - Result: code changes applied and reviewed; runtime validation pending (manual/GUI check requested).
- 2026-07-15: rapid_main stacked-panel min-size inflation control implemented:
  - Constrained `rapid_main` panel hinting in `_AdaptivePanelStack` to prevent oversized minimums from inflating startup width/height.
  - Added regression test `test_stacked_panel_hints_are_clamped_for_startup_safety` in `RapidPy/rapid_main/tests/test_sequence_ui_smoke.py`.
  - Result: code updated; runtime and unit verification still needed before closure claims.
- 2026-07-15: expanded window clamp/representation coverage for representative/resolution behavior:
  - Added small-display and high-resolution profile assertions to `RapidPy/rapid_main/tests/test_window_layout.py`
    (`test_app_window_minimum_constraints_for_small_displays`, `test_window_clamp_scales_with_high_resolution_profiles`).
  - Result: tests are added and require execution on an environment with PySide installed and offscreen-capable Qt.
- 2026-07-15: rapid_main startup compactness refinement started:
  - Reduced compact sidebar profile and startup width ratio in `RapidPy/rapid_main/rapid_main/app.py`
    (`_DEFAULT_SIDEBAR_WIDTH`, `_MIN_SIDEBAR_WIDTH`, `_MAX_SIDEBAR_WIDTH`, `_MAIN_MAX_WIDTH_RATIO`)
  - Updated `RapidPy/rapid_main/tests/test_window_layout.py` to keep sidebar/compact-cap assertions aligned with constants.
  - Result: code changes are in place; full test pass remains to be confirmed after runtime window manager smoke checks.

### Remaining phase 5 manual verification

- Confirmed by static code scan: every `RapidPy/*/app.py` now uses either `set_app_icon(...)` or equivalent explicit icon wiring.
- Manual runtime check still required to confirm all Windows taskbar buttons show the branded icon on first-party launch (especially multi-instance/packaged scenarios).

### Checklist snapshot

- [x] Sequence panel shows full values at 1280x720 and 1920x1080 in smoke-check script.
- [x] Side menu default and minimum widths prevent icon/text clipping.
- [x] No runtime `AttributeError` from unsupported widget calls (validated in smoke/compile checks).
- [x] DC motors / IRM-ARM / vacuum / SQUID paths are adapter-backed where configured and deterministic when no hardware path exists (plus deterministic fallback coverage in `test_diagnostic_services`).
- [x] Queue halt/recovery behavior is deterministic with safe fallback for covered paths (unit coverage and target-state assertions now green).
- [x] Queue state machine + adapter safety checks are covered by queue + measurement test pack (including pause/halt, lease lifecycle, and safe-state fallbacks).
- [ ] All listed app entry points show branded icon in taskbar at runtime (manual validation pending).
- [ ] `docs/vb6_parity_inventory.md` reflects completed milestones.

### Objective-critical remaining parity work (hardware + sample handling)

| VB6 workflow gap | Why it remains | Ownership-ready next condition |
|---|---|---|
| AF demagnetizer in-shell execution | `rapid_main` currently exposes AF demo/load helpers, but AF execution remains launch-based/demo-oriented with no full in-shell treatment/run-time state adapter contract | Add adapter-backed AF DAQ/AF control in queue/measurement path with preflight + ownership + safe-state integration |
| Thermal (`modThermal`, AF/Thermal probe workflows) | Not yet represented in `rapid_main` queue/measurement control graph | Implement thermal workflow adapters and queue step model or explicitly retire with migration rationale |
| VRM (`frmVRM`) | Remains in `vrm_logger` standalone app | Add a deterministic invocation and logging handoff path from rapid_main if it is in active scope |
| Rockmag routines (`frmRockmagRoutine`, `RockmagStep`, `RockmagSteps`) | Domain structure and run mapping still unimplemented in queue model | Add explicit step model + queue execution mapping with output artifacts |
| Advanced math modules (`modListenAndLog`, `modDataAnalysis`, `modStatusCode`) | Legacy error/status pipeline and diagnostic logging semantics still legacy/incomplete | Align status/error taxonomies and structured diagnostics output with rapid_main flow |
