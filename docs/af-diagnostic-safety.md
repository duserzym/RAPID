# AF Tuner ownership, capture and recovery

The main app and AF Tuner now use the same `rapidpy_common.hardware_safety`
journal and OS operation lease. Main-app imports remain available through the
compatibility module `rapid_main.hardware_safety`. An unfinished treatment blocks
AF Tuner output operations; an unfinished AF Tuner diagnostic blocks main-app
motion and treatments. Neither application may recover an operation owned by
another running process. Failed helper operations are recovered in AF Tuner with
the original board and limit profile; main-app operations use main-app recovery.

Connecting AF Tuner probes board/version/readback without booting or selecting a
coil. Changing the selected coil only updates the UI. The explicit **Test Active
Relay** action journals the test, checks the selected digital word and returns
outputs off. **Recover Outputs Off** checks process stops, zeros both ADwin-light-16
DAC channels and checks relay clear without booting or starting a process.

An entire frequency sweep owns one operation lease. Both sweeps and dense captures
journal their request before field I/O, retain measured results, and verify output
cleanup before success/data signals. Output failures and record-publication
failures keep the journal pending across restart. Diagnostic and recovery evidence
is retained under the safety journal's sibling `hardware_diagnostics/af-UUID/`
directory with a required `record.json` and SHA-256 artifact index. This safety
evidence does not automatically accept magnetic calibration into the main app.

Cancellation is checked before initialization and cooperatively during ramps or
captures. Closing an active helper asks its worker to stop and retains the window
until cleanup and worker completion. Closing suppresses a queued preview capture;
threads are never forcibly terminated.

Dense captures now carry an explicit coil in `AdwinDenseCaptureRequest`. The driver
selects/readbacks the coil after initialization so a subsequent board boot cannot
erase the selection. Loopback requests default to relays off. The clipping-test
helper supplies its selected coil as well. Unsupported amplitudes, rates, timing,
channels and memory budgets are rejected before output initialization rather than
silently clamped. Invalid or incomplete returned arrays cannot become captures.

Source limits come from `VB6/ADwin/sineout.bas` (minimum process delay 800 at 25 ns,
50 kHz) and `VB6/ADwin/globals.inc` (1,000,000 monitor/output points). The main
calibrated AF planner and native ramp path also reject IO rates beyond this
shipped process limit. These limits are for the shipped legacy process; a different
compiled process requires a deliberate, validated implementation change.

AF Tuner has a screen-aware fit handler. Wide windows show controls beside plots;
compact windows offer Controls/Plots tabs. Recovery buttons remain accessible at
736×720, and switching layouts preserves both widgets. Narrower viewports may
scroll instead of silently clipping controls.

Clipping scans hold one guarded operation across both scan directions. The
operator's scan maximum is an explicit diagnostic ceiling (at most 10 V), separate
from accepted treatment calibration. Zero-voltage baseline measurements are
supported only with that explicit trial ceiling. Scan evidence preserves each
request and raw returned capture; completing a scan does not accept new treatment
limits. Invalid scan ranges fail before instrument initialization.

ADwin communications startup and ordinary Connect only probe. Force reboot boots
the configured firmware under a guarded operation and verifies outputs off before
reporting connection. It does not try alternative firmware files. Loopback and
self-test own the board until checked cleanup. Manual DAC/digital writes first
stop existing processes and clear outputs, then retain the lifetime lease across
clicks. Recover Outputs Off or closing the window verifies cleanup and releases
ownership. Board bindings cannot change while manual outputs or workers are
active. A partial boot retains the controller for recovery even if connection
fails. Cancellation never emits a success signal; close waits for actual thread
completion without terminating a worker or discarding a running thread.

Both helpers use Controls/Plots tabs on compact screens, preserving their widgets
and recovery actions. DAC channels are 1..2 and ADC channels 1..16 for the supported
ADwin-light-16 hardware. Native DAC writes reject unsupported channels, nonfinite
voltages and values outside +/-10 V before DLL output calls.

Remaining full-system work includes guarded ownership for motor
helpers and embedded direct diagnostics, non-treatment motion/acquisition restart
recovery, calibration acceptance tooling, probe workflows, rebuilding the portable
pilot, and physical/scientific acceptance. No actual instruments were actuated by
the development checks.

Validation includes helper/main ownership, manual output lifetime and original
board cleanup, explicit boot versus read-only probe, the complete clipping range,
worker cancellation/signal ordering, native capture and compact-control visibility
regressions. Six-panel/six-helper source UI smoke passes. Native interfaces are
injected fakes in development checks; compact screenshots were inspected at
736x720. See ROADMAP.md for the latest complete suite checkpoint.
