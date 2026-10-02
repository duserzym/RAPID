# RAPID transition readiness — 2026-08-29

## Executive verdict

Do not replace the VB6 application with RapidPy for unattended or production
measurements yet. RapidPy is **transition/testing software**, not a full VB6
replacement.

The P0 measurement-integrity path is now **code-complete and covered by replay
and fault tests**. It is **not hardware-validated**: no part of it has been run
against the physical RAPID system, and no signed physical evidence exists for
the measurement sequence, rejection/recovery, motion/interlocks, safe halt,
holder integrity, output parity, or restart behavior.

Two things gate the replacement claim:

1. physical acceptance on the RAPID system (see
   `docs/rapid_hardware_acceptance_procedure_2026-08-29.md`);
2. a no-communication runtime smoke test of the legacy application, which now
   compiles on this computer (gates 1-3 closed 2026-09-03) but has not yet been
   launched against a `Paleomag.ini`.

## VB6 gate status on this computer (updated 2026-09-03)

Superseded twice: first by the English VB6 repair, then on 2026-09-03 when the
Visual Studio 6.0 Enterprise setup was finally run and the project compiled.
Four of the six gates are now closed.

### Gate 1 — English IDE startup: **verified**

- `C:\Program Files (x86)\Microsoft Visual Studio\VB98\VB6.EXE`, file version
  `6.00.8176` (English RTM).
- Mixed Simplified Chinese templates, T-SQL components, wizards, Common Tools
  binaries, Designer satellites, and five system components were replaced with
  hash-matching English-media versions.
- Displaced files are preserved at `C:\VB6-English-Repair-Backup-20260829-2`.
  **Do not delete that backup.**

### Gate 2 — real `Paleomag v3.vbp` load: **verified**

The project compiles, which requires every form, module, class, and control to
bind. Two things had to be resolved first.

**`Error accessing the system registry` — root cause and fix.** This was never a
project fault. While a project loads, the IDE registers the project's own type
library and the licence keys for its licensed controls into the machine COM
hive (`HKLM\SOFTWARE\Classes`, `...\Classes\Licenses`,
`...\WOW6432Node\Classes`). A standard user token cannot write there — UAC hands
even a member of Administrators a filtered token — and VB6 predates UAC, so it
reports the registry error instead of requesting elevation.

The fix is to run the IDE elevated. `VB6/Start-VB6.ps1` diagnoses the exact keys
and self-elevates; `-Shortcut` writes a desktop shortcut with the elevation flag
set. Granting the user account write access to `HKLM\SOFTWARE\Classes` was
deliberately **not** done: it would let any process running as that user
redirect COM registrations for every account on the machine, permanently.

**`vbSendMail_v3.0.dll` — resolved.** Obtained from the operator's legacy archive
at `F:\Paleomag2013\vbSendMail\vbSendMail.dll` (FreeVBCode.com, file version
3.06.0005). Verified 32-bit and confirmed to carry type-library GUID
`{332B82D3-3ED6-11D4-B1B5-00105AA5CCFF}` before installation, then copied to
`C:\Windows\SysWOW64\vbSendMail_v3.0.dll` and registered. It registers type
library **5.7**, exactly the version the project references.

**`MSCOMCTL.OCX` — resolved 2026-09-03.** The project asks for type library
**2.2**. Neither 6.01.9782 nor 6.01.9786 provides it: both embed 2.0, and even
the MS12-027 build (6.01.9834, KB2708437) only reaches **2.1**. Only
**KB3096896** (MS16-004) ships `mscomctl.OCX` **6.01.9846**, which embeds 2.2.

That package was downloaded from the Microsoft Download Center, verified by
Authenticode (valid, timestamped November 2020, chaining to Microsoft Root
Certificate Authority 2011), extracted with an administrative install, and only
its `mscomctl.OCX` was installed. The `comctl32.ocx` in the same package embeds
`ComctlLib` **1.5** while the project asks for **1.3**, so applying it would
have traded one mismatch for another.

Type library 2.2 is now registered against
`C:\Windows\SysWOW64\MSCOMCTL.OCX`, the previous control is kept as
`MSCOMCTL.OCX.bak-20260903-153324`, and the project now compiles with
`-NoFixups`: zero substitutions, zero unresolved references.

The build-time substitution machinery in `VB6/Build-VB6Project.ps1` remains for
machines that still carry an older control.

### Gate 3 — compiler edition and `/make`: **verified**

Previously `/make` answered `No make available in the Working Model Edition`.
The cause was that VB6 had never been installed: the `VB98` folder had been
copied, so the setup key carried no `ProductDir` and VB6 fell back to its most
restricted mode. The binaries all ran, including `LINK.EXE` and `C2.EXE`.

Running the English Visual Studio 6.0 Enterprise setup (volume
`VISUAL_BASIC_6`, `VS98ENT.STF`) registered
`ProductDir = C:\Program Files (x86)\Microsoft Visual Studio\VB98`, and the
compiler was enabled.

First successful build, 2026-09-03:

```
Build of 'PALEOMAG2013.exe' succeeded.
```

`build\vb6\PALEOMAG2013.exe` — 1,736,704 bytes, 32-bit i386, file version
`3.01.0009`, product "Paleomagnetic Magnetometer Control System 2013".
Full evidence in `build\vb6\vb6-build-receipt.json`. Since the MSCOMCTL fix
recorded under gate 2, the project builds with `-NoFixups` — no machine-local
substitution at all.

Note for anyone writing preflight checks: the per-user key
`HKCU\SOFTWARE\Microsoft\VisualStudio\6.0` and a "Visual Studio 6.0" uninstall
entry are **not** reliable indicators. This installation compiles with neither.
`ProductDir` plus the presence of `LINK.EXE`/`C2.EXE` is what actually predicts
`/make`.

### Gate 4 — Service Pack 6: **not applied, and not required for the build**

The IDE remains `6.00.8176` (RTM); the runtime `MSVBVM60.DLL` is `6.00.9848`
(SP6-era). The build above succeeded on the RTM IDE, so SP6 is no longer a
blocker for producing an EXE. Applying the legitimate English base SP6 and the
current signed Microsoft rollup is still worth doing for the security fixes.
Do not extract and overwrite OCXs to evade installer prerequisites.

### Gate 5 — no-communication runtime smoke test: **not yet run**

The EXE exists but has not been launched. It needs a `Paleomag.ini` before it
will start: `Sub Main` in `VB6/modProg.bas` reads the INI path from
`GetSetting(App.EXEName, "Settings", "INIFile", ...)` and, on first run, opens a
file dialog to choose one, then remembers it in the registry. Point it at the
lab's real `Paleomag.ini`, or at a copy of `VB6/Defaults.ini` for a
no-communication smoke test.

### Gate 6 — physical RAPID-system test: **not attempted**

No hardware was connected or actuated. The procedure and evidence schema are in
`docs/rapid_hardware_acceptance_procedure_2026-08-29.md`.

### Tooling

`VB6/LAUNCH-AND-BUILD.md` documents the five scripts:
`Test-LaunchReadiness.ps1`, `Find-VB6Projects.ps1`, `Start-VB6.ps1`,
`Install-VB6Dependency.ps1`, and `Build-VB6Project.ps1`.

Re-run the preflight from the repository root:

```powershell
powershell -ExecutionPolicy Bypass -File .\VB6\Test-LaunchReadiness.ps1
```

## Measurement bug and implemented behavior

The failure mode is a discrete one-flux-count change between the two zeros
bracketing the four holder orientations. The archived increments are
approximately X `0.090`, Y `0.106`, and Z `0.066` in calibrated 2G raw units.
Linear interpolation turns that step into a monotonic ramp and, after the four
orientation transforms, produces false holder vectors separated by roughly
90 degrees.

### VB6

`VB6/modMeasure.bas`:

- rejects any bracketing-zero axis delta greater than `0.02` before the legacy
  relaxed retry logic can accept it;
- independently rejects a dominant monotonic holder-frame staircase with span
  at least `0.04`;
- issues `CLP A` and `ResetCount A`, waits, and repeats the complete block;
- halts after the configured retry limit instead of recording the bad block;
- retains the last accepted global holder until a replacement block succeeds.

The updated VB6 project compiled successfully with `-NoFixups` on 2026-09-03.
The rejection behavior still requires the pending no-communication runtime
smoke test and physical-system replay before it is operationally accepted.

### RapidPy

`rapid_main.magnetometer` defines the coherent `BracketedMeasurementBlock`,
validates the zero pair before interpolation, applies the VB6 `i/5` baseline
weights and orientation mapping, subtracts the holder, and rejects the
independent monotonic-staircase signature. It raises
`FluxCountDiscontinuityError`, so no reduced result is available to save or
propagate.

`MeasurementWorker` retries only when the backend implements
`recover_flux_count_discontinuity`; otherwise the run fails. Recovery is an
explicit hardware operation, never a silent accept.

## P0 status — code-complete, awaiting hardware validation

| P0 item | Software status | Hardware status |
|---|---|---|
| 1. Real bracketed acquisition | **Code-complete.** `rapid_main.acquisition.BracketedAcquisitionService` runs the exact `Measure_ReadSample` order behind injected transport/lift/turn/clock protocols. | Not validated |
| 2. Safe recovery | **Code-complete.** `recover_flux_count_discontinuity` returns to the verified zero position, issues `CLP`/`RC`, waits, and re-latches. Retry exhaustion halts through the shared safe-state path. | Not validated |
| 3. Real, persistent holder | **Code-complete.** `HolderCorrection` + `HolderStateStore` + `HolderMeasurementService`; atomic install, previous correction retained on any failure, UI shows identity/magnitude/age/validity, blocks samples when absent or stale. | Not validated |
| 4. Fail closed in hardware mode | **Code-complete.** Factories raise `HardwareUnavailableError`; `UnavailableBackend` replaces silent simulators in UI surfaces; simulated output is labelled and redirected. | Not validated |
| 5. Motion and interlocks | **Not software-closable.** Every motion is verified in software and a failed motion aborts the block. | **Requires the physical system** |
| 6. Output and metadata parity | **Code-complete.** Transactional publish, duplicate-free resume, full provenance, resolved specimen metadata, deterministic VB6 parity fixtures. | VB6 builds successfully; side-by-side physical comparison has not been run |

### What "code-complete" means here

- The whole acquisition sequence, its faults, and its recovery are exercised by
  deterministic tests with no hardware.
- The archived X step and the equivalent Y and Z steps are rejected, with proof
  that holder state and production output are unchanged afterwards.
- No path returns a fabricated value: every failure raises.

### What is still explicitly missing in software

- The 2G range letter is chosen from the UI label; the mapping has not been
  confirmed against the instrument.
- `motion.zero_pos` / `motion.meas_pos` default to unconfigured. RapidPy will
  not invent lift positions, so hardware preflight blocks until an operator
  imports the legacy `[SteppingMotor]` INI values or enters measured ones.
- Holder direction policy defaults to `shared` to match VB6, which subtracts
  one holder for both sample directions. `strict` is available and tested but
  is a deliberate deviation, not parity.
- VB6 issues each motion twice (non-blocking start, then blocking repeat).
  RapidPy issues one blocking, verified motion. Same physical order, no
  un-awaited command, and a motion that misses its target fails loudly.

## What RapidPy still needs before replacing VB6

### P0 — remaining

1. **Physical validation of the whole measurement block.** Zero/measurement
   heights, all turn angles, direction conventions, home/limit behavior,
   changer pickup/drop-off, vacuum interlocks, emergency halt, timeouts, and
   recovery to a known safe state. RapidPy and VB6 must never own the same COM
   port at the same time.
2. **Side-by-side output parity against VB6.** The legacy application now
   builds on this computer (`build\vb6\PALEOMAG2013.exe`, 2026-09-03), so this
   comparison is unblocked once both applications can run against the same
   reference specimen.
3. **Confirmed instrument settings.** Range letters, `ReadDelay`, ARC delay,
   and the lift positions above.

### P1 — live capability parity

- AF demagnetizer execution, telemetry, decay, and interlock acceptance;
- vacuum live readback plus high-pressure and pump fault-to-halt tests;
- IRM/ARM, ADWIN, DAC, and MCC calibration/loopback evidence;
- furnace/oven control and thermal safety acceptance;
- susceptibility, VRM, and rock-magnetic acquisition with reproducible bundles;
- interrupted-queue recovery and physical safe-state acceptance;
- DC motor encoder, torque/current, stall, limit, and direction verification;
- serial retry/backoff, stale-reply rejection, and per-adapter raw transport
  logs.

### P2 — operator and analysis parity

- complete advanced plot/analysis parity and validated export formats;
- all dialogs, status displays, hardware monitor states, and operator messages
  verified at production resolution;
- dependency-clean installation and configuration migration exercised under
  the operator account. The wheel itself builds, installs to an isolated target,
  finds packaged assets/configuration, constructs the main window without the
  source tree, and exposes seven console entry points on this computer;
- calibration approvals exercised against real reference standards and signed
  by the operating lab. The append-only version/expiry/invalidity/rollback
  registry and measurement-record linkage are software-complete; that does not
  constitute physical calibration acceptance.

## Safe connection and acceptance sequence

The detailed, executable version — with the evidence schema every record must
carry — is in `docs/rapid_hardware_acceptance_procedure_2026-08-29.md`. In
summary:

1. Reconfirm the verified VB6 IDE/project/build gates, then complete the
   no-communication runtime smoke test against a copy of `VB6/Defaults.ini`.
2. Launch VB6 and RapidPy separately in no-communication mode; load all primary
   forms and confirm configuration/output paths.
3. Inventory COM ports and map each physical device. Keep outputs and port
   ownership separate.
4. Connect one subsystem at a time. Read-only identity/status queries first,
   then dry motion with no specimens.
5. Acquire and retain raw stable zero/count/DVM logs before running a sample.
6. Replay the archived discontinuous block and prove both applications reject
   it without updating holder or sample output.
7. Measure an empty holder repeatedly, then a reference specimen, and compare
   raw blocks, transformed vectors, statistics, and saved files side by side.
8. Only expand to treatment queues after halt, limit, vacuum, timeout, and
   restart tests pass with saved evidence.

## Current software verification

- Full RapidMain test suite: **454 passed** (`python -m unittest discover -s
  tests -p 'test_*.py'` from `RapidPy/rapid_main`); the pre-existing baseline
  was 261.
- New coverage added in this pass (121 tests): bracketed acquisition sequence
  and faults (22), live transport and backend facade (21), holder state and
  holder command (19), queue-level fail-closed and holder behavior (13),
  transactional output and simulation isolation (15), VB6 parity fixtures (18),
  replay fixtures (13), and holder UI (4, inside `test_measurement_panel`).
- October 2 main-shell coverage adds truthful live/simulated/unavailable/fault
  Dashboard snapshots, responsive glass-card reflow without horizontal
  overflow, synchronized workflow/sample/step state, wired Flow/session
  actions, halt-after-confirm shutdown ordering, atomic sequence documents,
  malformed-file reporting, and executable/saveable imported sequences.
- The current checkpoint also covers truthful empty plot/sample/queue states,
  real `.sam`/`.csv` sample-index selection into the queue with retained
  metadata registrations, hardware-mode refusal of simulated AF examples, and
  versioned atomic settings backup/validated restore with active-run blocking
  and explicit restart-required state.
- The first hardware-dialog glass/accessibility slice covers Vacuum, IRM/ARM,
  and SQUID with centralized semantic states, non-color-only status text,
  accessible control names, compact 360x520 geometry, a scrollable SQUID
  settings surface, config-backed SQUID values, and explicit fail-closed
  unavailable states. Live hardware behavior is still an acceptance gate.
- `python -m compileall` is clean across `rapid_main`, its tests,
  `updown_control`, and `rapidpy_common`. The repository configures no linter
  or type checker (no ruff/flake8/mypy config and no lint CI job), so none was
  introduced here.
- VB6 static check: `modMeasure.bas`, `MeasurementBlock.cls`, and
  `modMotor.bas` have balanced `If`/`With`/`For`/`Select`/`Do`/procedure blocks
  with continuations joined and single-line `If` forms excluded. No VB6 source
  was changed in this pass.
- VB6 source: the full project compiled successfully with `-NoFixups` on
  2026-09-03. Runtime verification of the measurement change is still pending
  the no-communication smoke test and physical-system replay.
