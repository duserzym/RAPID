# Prompt for Claude Code Opus 5 Max

Copy everything below into Claude Code Opus 5 Max while it is opened at the root of this repository.

---

You are working in the RAPID instrument-control repository on the current branch. Your goal is to close the remaining safety and functional gaps required for RapidPy to replace the legacy VB6 Paleomag application, beginning with the P0 live-measurement path. Treat this as laboratory instrument software: do not fabricate hardware evidence, do not silently fall back to simulation, and do not allow a rejected measurement block to update holder state or production output.

Read these files first:

- `docs/rapidpy_transition_readiness_2026-08-29.md`
- `RapidPy/rapid_main/PRODUCTION_READINESS_ASSESSMENT.md`
- `docs/vb6_parity_inventory.md`
- the pasted/archive diagnosis if it exists in the working context
- `VB6/modMeasure.bas`
- `VB6/MeasurementBlock.cls`
- `RapidPy/rapid_main/rapid_main/magnetometer.py`
- `RapidPy/rapid_main/rapid_main/measurement_worker.py`
- `RapidPy/rapid_main/rapid_main/diagnostic_services.py`
- `RapidPy/rapid_main/rapid_main/hardware_contracts.py`
- `RapidPy/rapid_main/rapid_main/queue_compiler.py`
- `RapidPy/rapid_main/rapid_main/panels/measurement.py`
- `RapidPy/updown_control/updown_control/app.py`

Preserve all current user changes and existing commits. Inspect `git status`, `git diff`, and recent history before editing. Never use destructive Git commands. Split your work into focused commits and list every commit in your handoff.

## Verified VB6 environment state on this computer

The VB6 launch section in `docs/rapidpy_transition_readiness_2026-08-29.md` is stale. Update that document as part of this handoff; do not repeat its claim that the IDE and Microsoft controls are absent.

The following was verified directly on this computer without rebooting:

- English VB6 RTM `6.00.8176` is installed at `C:\Program Files (x86)\Microsoft Visual Studio\VB98\VB6.EXE`.
- The IDE opens normally to the English New Project dialog and can load a small test project without a startup error.
- DAO 3.5, `VB6.OLB`, `VB6EXT.OLB`, `MSDERUN.DLL`, `MSO97RT.DLL`, `MRT7ENU.DLL`, and the required Microsoft OCX files were restored or registered.
- Mixed Simplified Chinese VB6 templates, T-SQL components, wizards, Common Tools binaries, Designer satellites, and five system components were replaced with hash-matching English-media versions. The active VB6/Common/Designer trees had zero files reporting Chinese language metadata after repair.
- Displaced files are preserved at `C:\VB6-English-Repair-Backup-20260829-2`. Do not delete that backup.
- The Visual Studio setup front end that requested a reboot was terminated, and the machine was not rebooted.
- The Paleomag project must currently be launched as administrator to avoid `Error accessing the system registry` during design-time component loading.

Do not claim that the VB6 project is build-ready yet. These blockers remain:

1. `VB6.EXE /make` reports `No make available in the Working Model Edition`. The English Enterprise edition/license registration is incomplete or inconsistent. Do not bypass licensing, invent a product key, or modify license-related registry data. Record this as an environment blocker requiring legitimate licensed media/registration.
2. `VB6/Paleomag v3.vbp` requires `MSCOMCTL.OCX` type library `2.2`, while the installed English `C:\Windows\SysWOW64\MSCOMCTL.OCX` is file version `6.01.9782` and registers type library `2.0`. Registration succeeds but the real project reports that `MSCOMCTL.OCX` could not be loaded. A disposable copy of the project loaded completely when only that reference was changed from `2.2` to `2.0`; the repository project was deliberately left unchanged. Do not downgrade the committed reference or alias the type library without explicit user approval and cross-machine compatibility evidence.
3. The base English Visual Studio/VB6 Service Pack 6 is not installed. Microsoft's signed `VB60SP6-KB2708437-x86-ENU.msi` refused with exit `1603` because SP6 was absent. Obtain the legitimate English base SP6 package first, then apply the current signed Microsoft common-controls/security rollup with restart disabled. Do not extract and overwrite OCXs to evade installer prerequisites.
4. `vbSendMail_v3.0.dll` is still missing. It is referenced by the project and documented in `VB6/README.txt`, but it is not in this repository or its Git history. Obtain it from the authorized legacy RAPID archive or original licensed/source package, verify it, and register its 32-bit COM server. Do not use generic DLL-download sites.

Before changing VB6 source, re-run `VB6/Test-LaunchReadiness.ps1`, inspect the real project references, and distinguish IDE startup, project loading, and compilation as three separate gates. Do not install system software, edit registry permissions, actuate hardware, or reboot from Claude Code unless the user explicitly authorizes that action in the active session.

## Known safety defect and code already present

The archived defect is a discrete 2G flux-count step between the two zero readings that bracket four holder orientations. Observed one-count increments are approximately X `0.090`, Y `0.106`, and Z `0.066` in calibrated raw units. Linear interpolation across the step creates a monotonic ramp; rotating the four contaminated readings into the holder frame creates a false pattern separated by roughly 90 degrees.

The repository already contains:

- VB6 categorical rejection/re-latch/re-zero/retry logic in `VB6/modMeasure.bas`;
- RapidPy `BracketedMeasurementBlock`, `validate_zero_pair`, `reduce_bracketed_measurement`, and `FluxCountDiscontinuityError`;
- worker logic that invokes `recover_flux_count_discontinuity` only when a backend explicitly implements it;
- unit tests for the archived X step, valid reduction, monotonic staircase rejection, and recovery-hook ordering.

Do not weaken those rules or reintroduce “accept after N retries.” A discontinuous block must always be discarded. Review the implementation and add tests for Y and Z steps, down-orientation transforms, holder subtraction, retry exhaustion, no recovery hook, and proof that output/holder state is unchanged after rejection.

## Phase 1: implement the production bracketed SQUID path

Replace the current single-startup-baseline/live-single-read design with an explicit, testable acquisition state machine. The production measurement backend must acquire:

1. verified up/down motion to zero position;
2. zero-before as a coherent three-axis latch/count/DVM observation;
3. motion to measurement position;
4. four specimen/holder observations at the exact VB6 orientation convention;
5. verified turns between observations and return to the known start orientation;
6. motion back to zero;
7. zero-after as a coherent observation;
8. a `BracketedMeasurementBlock` containing raw vectors, holder vectors, direction, calibration/range, timestamps, positions, and audit identifiers;
9. validation and reduction before any accepted-state mutation or output write.

Use interfaces and dependency injection so every state transition is deterministic in tests. Do not bury changer, vertical motion, turning, SQUID, or logging operations in untestable UI code. Prefer an acquisition service with typed protocol boundaries and a clearly ordered command/event record.

Verify the exact VB6 semantics rather than guessing:

- baseline interpolation weights are `i/5` for positions 1 through 4;
- confirm all 0/90/180/270/360 turn calls and the up/down orientation mapping from source;
- confirm whether raw SQUID values combine DVM plus count increments and preserve both components in evidence;
- confirm range and per-axis calibration application order;
- preserve the last accepted holder until a new holder block passes every check.

The current `RawSquidClient.read_xyz_raw` returns combined values. Extend the transport/evidence model to retain the atomic latch sequence, separate counter and DVM readings, raw replies, timestamps, and range. Reject partial, stale, malformed, mismatched, timed-out, or non-finite observations. Ensure all three axes belong to one documented latch/read cycle.

Implement `recover_flux_count_discontinuity` for the live backend. It must move to a safe zero state, clear/re-latch/reset the relevant 2G path using the verified hardware commands, wait for stabilization, record the recovery, and restart the entire block. On retry exhaustion or recovery failure, halt through the shared safe-state path and produce a clear operator error. Never synthesize a successful result.

## Phase 2: make holder correction a real measured state

The queue currently treats `Holder` mostly as changer motion. Implement a complete holder-measurement command:

- acquire and validate a full bracketed block;
- persist holder ID/location, direction, four raw and adjusted vectors, zeros, range/calibration, timestamps, software version, and validation evidence;
- install the new correction atomically only after success;
- keep the previous correction on rejection, abort, timeout, or write failure;
- expose holder magnitude, induced/asymmetry metrics, age, identity, and validity in the operator UI;
- associate the exact holder record/version with each sample result;
- prevent sample measurement if required holder state is absent, stale, direction-incompatible, or otherwise invalid.

Add replay fixtures representing stable holders and the archived failure. Test queue-level behavior, not just math helpers.

## Phase 3: make hardware mode fail closed

Audit every backend factory and `QueueHardwareBackend`. In explicit simulation/no-communication mode, simulated backends are allowed and must be visibly labeled. In hardware mode, any missing package, driver, port, calibration, adapter, or connection must fail preflight and prevent queue start. Remove broad exception-to-simulator fallback in hardware mode. Preserve diagnostic details and a safe operator message.

Guarantee that simulated results:

- carry an unmistakable simulation marker in memory, logs, UI, and files;
- cannot be written to production output paths unless explicitly enabled for a test;
- cannot be mistaken for hardware readiness evidence.

Add tests for missing imports, unavailable ports, constructor errors, failed identity checks, and mixed live/sim configurations.

## Phase 4: metadata, persistence, and parity evidence

Trace the complete metadata path from sample selection and queue compilation to `SpecimenMeta` and every output. Remove blank placeholder fields where VB6 has real values. Preserve sample, site, location, volume, orientation, holder association, treatment, comments, run ID, calibration version, and raw evidence references.

Make accepted output transactional: write to a temporary file or transaction, fsync/close as appropriate, then atomically publish. A rejected/aborted block must not append an accepted measurement. Test crash/exception points and restart behavior. Prevent duplicates when resuming an interrupted queue.

Create deterministic VB6 parity fixtures for:

- baseline interpolation and holder subtraction;
- all four rotations in both sample directions;
- average vector, moment scaling, direction/declination/inclination, CSD/noise metrics, and output formatting;
- holder replacement and retention;
- rejected zero/count discontinuities on X, Y, and Z;
- interruption and resume.

Document every intentional difference from VB6 and why it is safer or required.

## Phase 5: remaining live acceptance work

After the P0 path is complete in software, work through the P1/P2 inventory in `docs/rapidpy_transition_readiness_2026-08-29.md`: AF, vacuum, IRM/ARM, ADWIN/DAC/MCC, thermal, susceptibility, VRM, rockmag, motors, serial robustness, plots/exports, installer/config migration, backup/restore, and rollback.

For work that requires the physical RAPID system, provide an executable acceptance procedure and evidence schema instead of claiming success. Every hardware test record should include date/time, operator, machine/software commit, configuration hash, device/port identity, commands and raw replies, expected result, observed result, pass/fail, and artifact paths.

## Required verification

Run the smallest relevant tests while iterating, then the full RapidPy suite. Add type/lint checks already supported by the repository. Perform static checks on VB6 changes. Attempt a VB6 compile and no-communication-mode smoke test only after the licensed-edition, SP6, `MSCOMCTL` `2.2`, and `vbSendMail_v3.0.dll` gates above are satisfied; until then, report those gates precisely instead of treating IDE startup as compile verification. Do not connect or actuate hardware unless the operator explicitly authorizes it and confirms the physical area is safe.

Before finishing:

1. show `git status` and ensure no generated test/config artifacts are included;
2. run `git diff --check`;
3. update the readiness assessment with code-complete versus hardware-validated status;
4. produce focused commits, with tests paired with their implementation;
5. report exactly what is complete, what remains, tests and counts, commits, hardware evidence obtained, and the first safe test to run when the system is connected.

Also report the VB6 gates separately in the final handoff:

- English IDE startup: verified or regressed;
- real `Paleomag v3.vbp` load: verified or blocked, with the exact reference/dialog;
- compiler edition and `/make`: verified or blocked;
- no-communication runtime smoke test: verified or blocked;
- physical RAPID-system test: not attempted unless explicitly authorized, otherwise include the exact safe procedure.

Acceptance for the replacement claim requires all P0 items to be code-complete and pass replay/fault tests, plus signed physical-system evidence for measurement sequence, rejection/recovery, motion/interlocks, safe halt, holder integrity, output parity, and restart behavior. Until then, label RapidPy as transition/testing software, not a full VB6 replacement.

---
