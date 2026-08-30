# Operator Review Notes — 10 July 2026

## Scope reviewed

- `rapid_main` roadmap execution checkpoint through AF/queue control hardening.
- Files inspected:
  - `RapidPy/rapid_main/rapid_main/app.py`
  - `RapidPy/rapid_main/rapid_main/panels/measurement.py`
  - `RapidPy/rapid_main/rapid_main/panels/sample_queue.py`
  - `RapidPy/rapid_main/rapid_main/measurement_worker.py`
  - `RapidPy/rapid_main/rapid_main/workflow.py`
  - Related tests in `RapidPy/rapid_main/tests`

## Findings (as of this checkpoint)

### What is working well

- Device ownership is enforced for measurement and shared diagnostic launchers.
- Preflight, step execution, and measurement phases are now bound to timeout/cancellation controls.
- Queue execution path is connected from queue panel → app window → measurement panel.
- Pause/Halt ergonomics are routed through a shared runtime path.
- Queue/measurement state now surfaces through a canonical workflow phase channel.
- Shell layout persistence and reset actions are implemented.
- VB6 INI import path has explicit mapping/warning feedback.

### Observed gaps / risks

- No signed VB6 operator sign-off yet for the new AF/queue behavior; remaining human validation is still needed.
- Full AF + SQUID + device behavior is still simulation-oriented in this execution slice.
- Some VB6 migration rows remain `Not assessed` in `vb6_parity_inventory.md` (especially specialized routines and reporting paths).

### Approval and follow-up

- Use this checkpoint note as the required execution artifact before broader calibration automation work starts.
- Before automating additional calibration routines, schedule an operator walkthrough of:
  - AF demo run from queue
  - Pause/Halt response
  - VB6 INI import workflow
  - Startup/layout recovery behavior
- Keep this review record linked in the active roadmap execution notes.

## Status

- Checkpoint document created to satisfy roadmap Step 5 evidence.
- Human/operator sign-off status: **pending** (to be completed in next review session with lab operator).
