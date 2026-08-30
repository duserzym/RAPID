# RAPID transition readiness — 2026-08-29

## Executive verdict

Do not replace the VB6 application with RapidPy for unattended or production measurements yet.

The newly reported 2G blank-holder failure is now guarded in both implementations: a discontinuous pair of bracketing zeros is rejected before interpolation, the current holder correction is preserved, and a backend with a safe recovery operation can re-latch/re-zero and repeat. RapidPy unit tests cover the archived `0.09000563` X-axis step, a valid bracketed block, the independent monotonic-staircase guard, and worker recovery.

That guard is not yet active in RapidPy's actual live acquisition path. The live adapter does not acquire the same coherent two-zero/four-orientation block as VB6. Completing and validating that path is the primary transition blocker.

## VB6 launch status on this computer

`VB6/Test-LaunchReadiness.ps1` found the 32-bit VB6 runtime and eight registered dependencies, but launch is currently blocked by:

- no `PALEOMAG2013.exe`/`PALEOMAG.exe` and no VB6 IDE/compiler;
- missing `vbSendMail_v3.0.dll`;
- missing `MSDERUN.DLL`;
- missing `mshflxgd.ocx`;
- missing `TABCTL32.OCX`;
- missing `comct332.ocx`;
- missing `RICHTX32.OCX`.

Run from the repository root:

```powershell
powershell -ExecutionPolicy Bypass -File .\VB6\Test-LaunchReadiness.ps1
```

When a known-good executable or VB6 IDE and the licensed legacy components have been installed and registered, re-run the preflight. It checks both file presence and 32-bit type-library registration. Use `C:\Windows\SysWOW64\regsvr32.exe` for 32-bit controls. A passing preflight is necessary but not sufficient: a no-communication-mode launch and form-loading smoke test are still required because no VB6 compiler or executable is available on this machine today.

## Measurement bug and implemented behavior

The failure mode is a discrete one-flux-count change between the two zeros bracketing the four holder orientations. The archived increments are approximately X `0.090`, Y `0.106`, and Z `0.066` in calibrated 2G raw units. Linear interpolation turns that step into a monotonic ramp and, after the four orientation transforms, produces false holder vectors separated by roughly 90 degrees.

### VB6

`VB6/modMeasure.bas` now:

- rejects any bracketing-zero axis delta greater than `0.02` before the legacy relaxed retry logic can accept it;
- independently rejects a dominant monotonic holder-frame staircase with span at least `0.04`;
- issues `CLP A` and `ResetCount A`, waits, and repeats the complete measurement block;
- halts after the configured retry limit instead of recording the bad block;
- retains the last accepted global holder until a replacement holder block succeeds.

This source change has been reviewed statically, but it has not been compiled because this computer lacks VB6 and a built executable.

### RapidPy

`rapid_main.magnetometer` now defines a coherent `BracketedMeasurementBlock`, validates the zero pair before interpolation, applies the VB6 `i/5` baseline weights and orientation mapping, subtracts the holder, and rejects the independent monotonic-staircase signature. It raises `FluxCountDiscontinuityError`, so no reduced result is available to save or propagate.

`MeasurementWorker` retries only when the backend implements `recover_flux_count_discontinuity`; otherwise it fails the run. This makes recovery an explicit hardware operation instead of silently accepting or simulating a result.

## What RapidPy still needs before replacing VB6

### P0 — measurement integrity and safe hardware operation

1. **Implement the real bracketed acquisition.** `SquidBackendAdapter.test_connection` currently records one startup baseline and `read_squid` returns one current-position moment. The worker merely repeats that read and averages. It does not perform zero-before, four coherent 0/90/180/270-degree holder reads, zero-after, or the VB6 up/down and turn sequence. The live backend must return `BracketedMeasurementBlock` and log each raw count, DVM value, latch command, timestamp, motion position, and range as one auditable block.

2. **Implement safe recovery.** The live backend lacks `recover_flux_count_discontinuity`. Recovery must lift to the verified zero position, re-latch/clear/reset the 2G count path, wait for stabilization, discard the entire invalid block, and repeat from zero-before. Exhaustion must halt motion and preserve the last accepted holder.

3. **Make holder measurement real and persistent.** The queue compiler emits `Holder`, but `QueueHardwareBackend.holder` currently performs changer pickup/motion only. It does not measure or install a holder correction, and the UI shows holder/induced statistics as `N/A`. Holder identity, direction, four position vectors, zero-pair evidence, timestamp, and validity must be persisted and associated with every corrected sample.

4. **Fail closed in hardware mode.** Several backend factories and `QueueHardwareBackend` catch broad exceptions and substitute no-communication backends. In hardware mode, missing drivers, ports, calibration, or adapters must block preflight and queue start. Simulation must be an explicit, highly visible mode and its outputs must never share production destinations without a simulation marker.

5. **Validate motion and interlocks on the physical system.** Confirm zero/measurement heights, all turn angles and direction conventions, home/limit behavior, changer pickup/drop-off, vacuum interlocks, emergency halt, timeout behavior, and recovery to a known safe state. RapidPy and VB6 must never own the same COM port at the same time.

6. **Prove output and metadata parity.** RapidPy currently starts measurements with several `SpecimenMeta` fields blank. Preserve the full sample/site/location/orientation/volume/treatment/comment metadata and compare calculation and output artifacts against VB6 fixtures. Verify atomic writes, interrupted-run handling, duplicate prevention, and that a rejected block leaves no accepted result.

### P1 — live capability parity

- AF demagnetizer execution, telemetry, decay, and interlock acceptance on the real rig;
- vacuum live readback plus high-pressure and pump fault-to-halt tests;
- IRM/ARM, ADWIN, DAC, and MCC calibration/loopback evidence;
- furnace/oven control and thermal safety acceptance;
- susceptibility, VRM, and rock-magnetic acquisition with reproducible run bundles;
- interrupted-queue recovery and physical safe-state acceptance;
- DC motor encoder, torque/current, stall, limit, and direction verification;
- serial retry/backoff, stale-reply rejection, and adapter-by-adapter raw transport logs.

### P2 — operator and analysis parity

- complete advanced plot/analysis parity and validated export formats;
- all dialogs, status displays, hardware monitor states, and operator messages verified at production resolution;
- installation, configuration migration, calibration backup/restore, and rollback procedure exercised on the RAPID computer.

## Safe connection and acceptance sequence

1. Install or build VB6 and satisfy the launch preflight without connecting hardware.
2. Launch VB6 and RapidPy separately in no-communication mode; load all primary forms and confirm configuration/output paths.
3. Inventory COM ports and map each physical device. Keep outputs and port ownership separate.
4. Connect one subsystem at a time. Start with read-only identity/status queries, then dry motion with no specimens.
5. Acquire and retain raw stable zero/count/DVM logs before running a sample.
6. Inject/replay the archived discontinuous block and prove both applications reject it without updating holder or sample output.
7. Measure an empty holder repeatedly, then a reference specimen, and compare raw blocks, transformed vectors, statistics, and saved files side by side.
8. Only expand to treatment queues after halt, limit, vacuum, timeout, and restart tests pass with saved evidence.

## Current software verification

- Full RapidMain test suite: 261 passed after this fix.
- RapidPy targeted magnetometer/worker tests: 23 passed.
- Existing RapidMain domain-service targeted tests before this fix: 139 passed.
- VB6 source: static review and launch preflight only; no compile/runtime verification is possible on this computer yet.
