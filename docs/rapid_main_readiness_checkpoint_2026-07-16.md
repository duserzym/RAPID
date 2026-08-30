# RapidPy Production Readiness Checkpoint - 2026-07-16

## Immediate assessment

No. `rapid_main` is not yet fully production-ready against the full VB6 scope, and the repository is not fully remediated for all migration paths.

Current evidence says:

- Startup geometry and sidebar width are materially reduced in `rapid_main` (tightened further in this pass).
- Shared geometry guarding is stronger across app entrypoints.
- VB6 parity is still partial with explicit mapped/integrated components that retain production-readiness gates.
- DC motor diagnostics now provide live command/output/velocity feedback and best-effort torque display when available.

## What changed in this pass

1. `RapidPy/rapidpy_common/ui.py`
   - `clamp_window_geometry` currently uses:
     - `max_w = min(int(area_width * 0.35), area_width, 960)`
     - `min_w = min(int(area_width * 0.26), max_w)` with `min_w >= MIN_WINDOW_WIDTH`
     - `min_h = min(int(area_height * 0.24), max_h)`
   - This impacts all apps using shared geometry helpers.
   - Added nearest-screen fallback when a top-level widget is off-screen and has no overlap or `contains(topLeft)` hit.
     - `_screen_area_for_widget` now chooses the nearest `QScreen` center before defaulting to primary.

2. `RapidPy/rapid_main/rapid_main/app.py`
   - Startup geometry remains compact with `_DEFAULT_WINDOW_SIZE = (560, 460)`.
   - Rapid main width safety is capped by `_MAIN_MAX_WIDTH_RATIO` and a 960 px large-display ceiling in `_clamp_main_window_size`; very narrow profiles below 1024 px relax the fixed width floor so startup can remain compact on tiny/offscreen test monitors.
   - In this pass `_MAIN_MAX_WIDTH_RATIO` is now `0.20` and `_MAIN_MIN_WIDTH` is `240`, with a tighter `_clamp_main_window_size` footprint retained on small and off-screen profiles.
   - Sidebar envelope is now a compact icon rail:
     - `_DEFAULT_SIDEBAR_WIDTH = 52`
     - `_MIN_SIDEBAR_WIDTH = 42`
     - `_MAX_SIDEBAR_WIDTH = 58`
     - `_SIDEBAR_RESTORE_RATIO = 0.075`
   - Boot-time clamp pass now runs before window presentation and after restore paths.

3. `RapidPy/rapid_main/rapid_main/dialogs/dc_motors.py`
   - Added live plotting path with `pyqtgraph` for:
     - target/input command
     - actual/output feedback
     - error
     - velocity1/velocity2
     - input-output delta
     - torque feedback when available
   - Telemetry thread is continuous and updates are rendered efficiently:
     - `_refresh_plots` redraws only when new samples exist and chart is not paused.
     - logging cadence is throttled for responsiveness.
   - Torque is extracted from aliases:
     - `actual_torque`, `feedback_torque`, `torque`, `torque_feedback`.
     - Missing torque gracefully shows `N/A`.

4. `RapidPy/rapid_main/rapid_main/dialogs/squid_comm.py`
   - `Test Connection` executes backend `test_connection()` and updates user status with success/fail color.
   - Dialog uses injected backend with simulation-awareness in status text.

5. Test and contract additions
   - `RapidPy/rapid_main/tests/test_window_layout.py`
     - Added and updated compacting/clamping expectations for current ratio behavior.
     - Added restore regression checks to prevent oversized reopen on small/offset displays.
     - Added tiny-profile expectations so rapid_main can clamp below the legacy 360 px floor instead of reopening too wide on 800 px-class available geometries.
     - Added shared-guard regression coverage proving oversized top-level minimum sizes are lowered before the window is clamped back into the active work area.
   - `RapidPy/rapid_main/tests/test_app_bootstrap_window_contracts.py`
     - Added startup contract that all 16 app entrypoints import/call `apply_window_bounds_guard(app)`, all 16 `main.py` wrappers dispatch through packaged `app.main()`, and no entrypoint forces startup maximization.
     - Corrected the contract root so it scans `RapidPy/*/*/app.py` and `RapidPy/*/main.py` instead of the wrong repository depth.
   - `RapidPy/rapid_main/tests/test_sequence_ui_smoke.py`
     - Adjusted stacked panel minimum-size expectations for compact startup behavior.
   - `RapidPy/rapid_main/tests/test_queue_orchestration.py`
     - Added run-recovery assertion for `Running -> Pending` on restart.

## Current readiness against your questions

### "Does it fully port the functions of the VB6 code?"
- No.
- `docs/vb6_parity_inventory.md` still marks several areas as `Integrated`/`Mapped` with open production-readiness gates (notably thermal, SQUID/magnetometer transport maturity, rockmag reproducibility, VRM live run association, DAC/MCC loopback, and analysis/plot hardware-review surfaces).

### "Will all apps display and initialize nicely on any monitor configuration?"
- Improved but not fully proven.
- Shared clamps reduce oversized startup risk, new bootstrap tests enforce guard usage for all 16 app entrypoints, and window-layout tests now prove oversized minimums are normalized by the shared guard.
- Full proof across mixed DPI, taskbar insets, and hotplug permutations remains incomplete.

### "Does it auto resize and layout on any windows machine?"
- Better than before, but not guaranteed for all machines yet.
- Needs broader practical validation with per-machine DPI/taskbar/multi-monitor regression.

### "Is the main app fully functional on every aspect?"
- Not yet. Layout is improved and startup compacting is working, but VB6 parity and hardware-acceptance completeness are still open.

## Hard requirement status

- [x] Compact startup and left panel width pressure reduced.
- [x] Addressed `main.py` complaint for full-monitor startup growth by explicit startup clamping and shared guard behavior.
- [x] Static software contract confirms all 16 app entrypoints install the shared bounds guard and all 16 wrappers dispatch through packaged `app.main()`.
- [ ] Full VB6 function port complete.
- [ ] Full "any monitor" auto-layout proof for every packaged app in practice.
- [ ] End-to-end acceptance evidence for all active hardware paths.

## Next handoff objective

1. Continue proving startup and hotplug resize safety on real Windows monitors.
2. Close remaining VB6-not-assessed rows with acceptance criteria and evidence.
3. Keep DC motor panel telemetry efficient:
   - verify torque display under real QuickSilver outputs,
   - keep plot update budget predictable under slow links.
