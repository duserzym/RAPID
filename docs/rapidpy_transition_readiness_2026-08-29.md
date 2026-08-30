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
2. the VB6 environment gates below, which still prevent building or running the
   legacy application for side-by-side comparison.

## VB6 gate status on this computer (verified 2026-08-29)

The earlier version of this section was written before the English VB6 repair
and is superseded. Each gate below was re-verified directly, read-only, without
installing software, editing the registry, or rebooting.

### Gate 1 — English IDE startup: **verified**

- `C:\Program Files (x86)\Microsoft Visual Studio\VB98\VB6.EXE` is present,
  file version `6.00.8176` (English RTM).
- The IDE opens to the English New Project dialog and loads a small test
  project without a startup error.
- Mixed Simplified Chinese templates, T-SQL components, wizards, Common Tools
  binaries, Designer satellites, and five system components were replaced with
  hash-matching English-media versions; the active VB6/Common/Designer trees
  report no Chinese language metadata.
- Displaced files are preserved at `C:\VB6-English-Repair-Backup-20260829-2`.
  **Do not delete that backup.**
- The Paleomag project must currently be launched **as administrator** to avoid
  `Error accessing the system registry` during design-time component loading.

### Gate 2 — real `Paleomag v3.vbp` load: **blocked**

`VB6/Test-LaunchReadiness.ps1` now reports only one missing dependency:

```
[MISSING FILE] vbSendMail_v3.0.dll
```

Every other referenced control resolves and is registered (`MSSTDFMT.DLL`,
`msscript.ocx`, `MSDERUN.DLL`, `comdlg32.ocx`, `comctl32.ocx`, `mscomm32.ocx`,
`MSFLXGRD.OCX`, `mshflxgd.ocx`, `MSCHRT20.OCX`, `TABCTL32.OCX`, `comct332.ocx`,
`RICHTX32.OCX`, `MSCOMCTL.OCX`).

Two separate blockers remain for an actual project load:

- **`MSCOMCTL.OCX` type-library version.** The project requires
  `Object={831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0; MSCOMCTL.OCX`. The
  installed `C:\Windows\SysWOW64\MSCOMCTL.OCX` is file version `6.01.9782` and
  registers type library **2.0** only. The registry does contain a `2.2` key
  under that TypeLib GUID, but it is an **empty stub**: it has no default value
  and no `2.2\0\win32` path, so nothing resolves. Only `2.0` has
  `2.0\0\win32 = C:\Windows\SysWow64\MSCOMCTL.OCX`
  (`Microsoft Windows Common Controls 6.0 (SP6)`).
  A **disposable copy** of the project loaded completely when only that
  reference was changed from `2.2` to `2.0`. The repository project was
  deliberately left unchanged. Do not downgrade the committed reference and do
  not alias the type library without explicit user approval and cross-machine
  compatibility evidence.
- **`vbSendMail_v3.0.dll` is missing.** The project references
  `{332B82D3-3ED6-11D4-B1B5-00105AA5CCFF}#5.7#0 ... vbSendMail_v3.0.dll`, and
  `VB6/README.txt` documents installing and registering it. It is not in this
  repository and never was: `git log --all --diff-filter=A -- "*vbSendMail*"`
  returns nothing, and only the consuming form `VB6/frmSendMail.frm` is
  present. Obtain it from the authorized legacy RAPID archive or the original
  licensed/source package, verify it, and register the 32-bit COM server. Do
  not use generic DLL-download sites.

### Gate 3 — compiler edition and `/make`: **blocked**

- `VB6.EXE /make` reports `No make available in the Working Model Edition`.
- Read-only registry evidence is consistent with an incomplete or inconsistent
  edition/licence registration:
  `HKLM\SOFTWARE\WOW6432Node\Microsoft\VisualStudio\6.0\Setup\Microsoft Visual Basic`
  exists but carries **no values** (no `ProductDir`, no edition record), and
  `HKCU\SOFTWARE\Microsoft\VisualStudio\6.0` does not exist.
- The English Enterprise edition must be registered from legitimate licensed
  media. Do not bypass licensing, invent a product key, or modify
  licence-related registry data.

### Gate 4 — Service Pack 6: **blocked**

- IDE `VB6.EXE` is `6.00.8176` — **RTM, not SP6** (an SP6 IDE reports
  `6.00.9782`).
- The runtime `C:\Windows\SysWOW64\MSVBVM60.DLL` is `6.00.9848`, i.e. an SP6-era
  redistributable runtime. Runtime and IDE are therefore mismatched.
- Microsoft's signed `VB60SP6-KB2708437-x86-ENU.msi` refused with exit `1603`
  because the base SP6 is absent. Obtain the legitimate English base SP6
  package first, then apply the current signed Microsoft common-controls /
  security rollup with restart disabled. Do not extract and overwrite OCXs to
  evade installer prerequisites.

### Gate 5 — no-communication runtime smoke test: **blocked**

Blocked behind gates 2–4. No `PALEOMAG2013.exe` / `PALEOMAG.exe` exists on this
machine and the project cannot be compiled, so no form-loading or
no-communication runtime test has been performed. **IDE startup is not compile
verification and is not runtime verification.**

### Gate 6 — physical RAPID-system test: **not attempted**

No hardware was connected or actuated. The executable procedure and evidence
schema are in `docs/rapid_hardware_acceptance_procedure_2026-08-29.md`.

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

Reviewed statically only. Not compiled — see gates 3 and 4.

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
| 6. Output and metadata parity | **Code-complete.** Transactional publish, duplicate-free resume, full provenance, resolved specimen metadata, deterministic VB6 parity fixtures. | Side-by-side comparison with VB6 output still blocked by gates 2–4 |

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
2. **Side-by-side output parity against VB6.** Blocked until the legacy
   application can be built or a known-good executable is available.
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
- installation, configuration migration, calibration backup/restore, and
  rollback exercised on the RAPID computer.

## Safe connection and acceptance sequence

The detailed, executable version — with the evidence schema every record must
carry — is in `docs/rapid_hardware_acceptance_procedure_2026-08-29.md`. In
summary:

1. Close VB6 gates 2–4 without connecting hardware.
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

- Full RapidMain test suite: **376 passed** (`python -m unittest discover -s
  tests -t .` from `RapidPy/rapid_main`).
- New coverage added in this pass: bracketed acquisition sequence and faults
  (22), live transport and backend facade (21), holder state and holder command
  (19), queue-level fail-closed and holder behavior (12), transactional output
  and simulation isolation (15), VB6 parity fixtures (18), holder UI (4).
- VB6 source: static review and launch preflight only. No compile or runtime
  verification is possible on this computer until gates 2–4 are closed.
