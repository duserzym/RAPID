# RAPID physical acceptance procedure and evidence schema — 2026-08-29

This document is the executable procedure for work that cannot be closed in
software. Nothing in it has been performed. It exists so that when the physical
RAPID system is available, the evidence collected is sufficient to support (or
refuse) the claim that RapidPy replaces the VB6 Paleomag application.

**Authorization.** Do not connect, power, or actuate any part of the RAPID
system from an automated agent. Every step below is performed by an operator
who has confirmed the physical area is safe and clear.

---

## 1. Evidence schema

Every record produced by any test below must carry all of these fields. A test
without a complete record does not count as evidence.

| Field | Meaning |
|---|---|
| `record_id` | Unique id for this test execution |
| `datetime_iso` | UTC start and end, ISO 8601 with offset |
| `operator` | Person performing and signing the test |
| `machine` | Hostname of the controlling computer |
| `software_commit` | `git rev-parse HEAD` of this repository at run time |
| `software_version` | `rapid_main.software_version()` (set `RAPID_BUILD_COMMIT` at deploy time) |
| `config_hash` | `rapid_main.hardware_contracts.config_fingerprint(config)` |
| `device_identity` | Port, baud, adapter, firmware/serial reported by the device |
| `commands` | Ordered commands issued, verbatim |
| `raw_replies` | Verbatim device replies, including timeouts and empty reads |
| `expected_result` | Written down **before** the test is run |
| `observed_result` | What actually happened |
| `verdict` | `pass` / `fail` / `blocked`, with a reason for anything but `pass` |
| `artifact_paths` | Absolute paths of every file the run produced |
| `signature` | Operator sign-off |

RapidPy already emits most of this automatically per run:

- `workflow_summary.json` — phases, abort state, simulated flag, run id, flux
  recovery count, holder record id;
- `artifact_index.json` — every expected artifact with presence and size, plus
  the published paths;
- `provenance.json` — specimen metadata, calibration, range, holder record and
  timestamp, software version, config hash, published labels;
- `communication.tsv` — transport transcript;
- the block audit inside each `BracketedMeasurementBlock` — ordered command log
  with per-command start/end timestamps and replies.

Attach these to the record rather than transcribing them.

---

## 2. Preconditions (no hardware connected)

| # | Step | Expected |
|---|---|---|
| P1 | Close VB6 gates 2–4 (see the readiness document) | `VB6.EXE /make` produces `PALEOMAG2013.exe`; the real `.vbp` loads with no missing-control dialog |
| P2 | `powershell -ExecutionPolicy Bypass -File .\VB6\Test-LaunchReadiness.ps1` | No `[MISSING FILE]` lines |
| P3 | Launch VB6 in no-communication mode; open every primary form | No load error; output/config paths as configured |
| P4 | Launch RapidPy with `general.nocomm = true`; open every panel and dialog | No load error; every simulated backend is labelled |
| P5 | Confirm a simulated RapidPy run writes into `…/SIMULATED/` with `SIMULATION.txt` | Nothing lands in the production output path |
| P6 | Import the legacy INI, then confirm `motion.zero_pos` / `motion.meas_pos` | Both non-zero and distinct; otherwise hardware preflight must block |
| P7 | Set `general.nocomm = false` with no hardware attached and start a queue | Preflight **blocks** and names the missing device; no simulator substitution |

P7 is the fail-closed regression test. If a queue starts here, stop and fix the
software before touching hardware.

---

## 3. Port inventory and ownership

| # | Step | Expected |
|---|---|---|
| C1 | Enumerate COM ports; record device, VID/PID, and driver for each | Every port maps to exactly one physical device |
| C2 | Record which application owns which port | VB6 and RapidPy never hold the same port simultaneously |
| C3 | Configure RapidPy output directory separately from the VB6 output directory | No shared file targets |

---

## 4. Read-only identity and status (one subsystem at a time)

For each of SQUID, changer/lift/turning motors, vacuum, AF, IRM/ARM, DAC/MCC,
thermal, susceptibility:

| # | Step | Expected |
|---|---|---|
| R1 | Open the port and issue only identity/status queries | Device replies; RapidPy reports connected and **not** simulated |
| R2 | Disconnect the cable and repeat | RapidPy reports the real error; no synthetic reading appears |
| R3 | Record raw replies for both cases | Transcript captured in `communication.tsv` |

---

## 5. Dry motion, no specimens

| # | Step | Expected |
|---|---|---|
| M1 | Home the lift to top | Limit switch reached; internal status bit 4 set |
| M2 | Move the lift to `motion.zero_position()` and read back | Final position within the accepted tolerance; `MotionOutcome.ok` true |
| M3 | Move the lift to `motion.measurement_position()` and read back | As above |
| M4 | Command the lift to a target it cannot reach | Motion reports `ok = false`; RapidPy aborts the block and does not read the SQUID |
| M5 | Rotate the turning axis to 0/90/180/270/360 and read the encoder each time | Each within 5 degrees; `MotorTurn_360` equivalent re-labels 360 as 0 |
| M6 | Obstruct or stall the turn (safely) | Motion reports `ok = false`; block aborts |
| M7 | Changer move to several holes, plus pickup/drop-off with a dummy | Positions within 0.02 hole; pickup reports success |
| M8 | Trigger emergency halt mid-motion | All four axes halt; RapidPy returns to the safe-state path |
| M9 | Vacuum interlock: force a high-pressure condition | Queue blocks or halts through the safe-state path |

---

## 6. SQUID zero stability (no specimen, no holder change)

| # | Step | Expected |
|---|---|---|
| S1 | `CLP A`, `RC A`, wait the ARC delay, then latch and read X/Y/Z 100 times | Record counter and DVM separately for every read |
| S2 | Compute the per-axis spread of the zero reads | Well below the `0.02` continuous-drift limit |
| S3 | Record the observed one-count magnitude per axis | Compare against the archived `X 0.090`, `Y 0.106`, `Z 0.066`; update `DEFAULT_AXIS_FLUX_INCREMENTS` only with this measured evidence |
| S4 | Confirm the range letter actually set on the instrument for each UI label | The `_range_label` mapping in `squid_transport.py` is correct or is corrected |

---

## 7. Discontinuity rejection on the real instrument

| # | Step | Expected |
|---|---|---|
| D1 | Replay the archived discontinuous block through RapidPy (fixture injection, no hardware) | Block rejected; holder unchanged; no accepted output written |
| D2 | On hardware, induce or wait for a genuine flux-count step during a blank-holder block | RapidPy rejects the block, runs the recovery sequence, and repeats from zero-before |
| D3 | Force recovery to fail (e.g. counter reset refused) | Run halts through the safe-state path with a clear operator error; nothing saved |
| D4 | Exhaust the retry limit | Run halts; previous holder correction still active; output directory contains no accepted measurement |
| D5 | Repeat D1 against VB6 once it can be built | Same rejection; no holder or sample output update |

---

## 8. Blank holder and reference specimen

| # | Step | Expected |
|---|---|---|
| H1 | Measure the empty holder 10 times | Every block validated; installed correction changes only after a successful block |
| H2 | Compare consecutive holder corrections | Magnitude, induced/asymmetry, and CSD stable within the operator-agreed tolerance |
| H3 | Kill RapidPy between holder blocks | Previous correction and its file intact on restart |
| H4 | Measure a reference specimen up and down | Compare raw block, baseline-adjusted vectors, holder-frame vectors, average, moment, declination/inclination, and CSD against VB6 for the same specimen |
| H5 | Compare saved files byte-for-byte where formats are identical | Differences explained and recorded, or fixed |

---

## 9. Interruption, resume, and restart

| # | Step | Expected |
|---|---|---|
| I1 | Kill the process mid-run | No accepted measurement published; no staging directory left behind |
| I2 | Restart and resume the same specimen | Already-published steps skipped, remaining steps appended once |
| I3 | Power-cycle the controlling computer mid-run | As I1, plus holder correction file still valid |
| I4 | Resume a queue with interrupted rows | Operator must explicitly choose Resume / Re-run / Skip / Abort |

---

## 10. VRM physical session acceptance

Run these steps only after SQUID identity/status and port-ownership checks pass.
Use a dedicated output directory and an operator-approved reference specimen or
stable test source.

| # | Step | Expected |
|---|---|---|
| V1 | From idle `rapid_main`, launch VRM and inspect port ownership | Main SQUID clients disconnect; main measurement and SQUID actions remain blocked until the VRM process exits |
| V2 | Record a short new-file session, then stop normally | One unique `*.vrm.json` records `operator_stopped`, nonzero rows, live port, baseline/calibration, CSV size, and matching SHA-256 |
| V3 | Append a second session to the same CSV | The first manifest remains unchanged; a second uniquely named manifest records only the second session's row count and `write_mode=append` while hashing the resulting CSV |
| V4 | Close the VRM window during acquisition | Final manifest records `window_closed`; CSV is closed and hashable; main ownership releases after the child exits |
| V5 | Disconnect or fault the SQUID during acquisition | Final manifest records `error` and the real message; no success/completion claim appears |
| V6 | Run for the laboratory-approved long-duration interval | Sampling timing, baseline drift, missing rows, process memory, CSV integrity, and displayed decay remain within signed tolerances |
| V7 | Compare CSV values and timing with the VB6 VRM workflow on the same source | Differences are quantified, explained, and approved or fixed |
| V8 | Verify run association | A manifest with a real production `run_id` says `associated`; a launch without one remains explicitly `unassociated` and is not indexed as production evidence |

---

## 11. Rockmag physical run-bundle acceptance

Use a compiled routine whose treatment families have individually passed their
device/interlock acceptance. Do not start with the mixed Works preset while any
family still has a software preflight blocker.

| # | Step | Expected |
|---|---|---|
| K1 | Run the compiled Hawaiian AF preset in No-Communication mode | Output stays under `SIMULATED/`; `rockmag_run.json` names the compiled routine and exact completed labels; the artifact index digest matches the file |
| K2 | Attempt Rockmag the Works in hardware mode before the backfield route exists (and before SUSC holder/bridge inputs are accepted) | Plan blocks before ordinary hardware preflight or treatment dispatch and writes an `ABORTED` rockmag artifact naming the blocker |
| K3 | After AF acceptance, run Hawaiian AF on a reference specimen | Every requested label is treated and measured once in order; raw treatment/SQUID traffic, safe return, holder version, and rockmag artifact share one run ID |
| K4 | Halt between two rockmag steps | No partial accepted bundle is published; abort artifact lists only the completed prefix and the halt/final phase |
| K5 | Resume the interrupted specimen | Previously published labels are not duplicated; skipped duplicate labels and the remaining completed suffix are explicit |
| K6 | Exercise each additional approved family separately before adding it to a mixed routine | Field/readback, interlock, failure, and safe-return evidence passes that family's acceptance procedure |
| K7 | Run the approved mixed routine | `provenance.json`, `workflow_summary.json`, `rockmag_run.json`, scientific outputs, and `artifact_index.json` agree on routine, run, sample, operator, labels, outcome, and hashes |

---

## 11a. Susceptibility bridge and coil acceptance

Software decision and open questions:
`docs/susceptibility_integration_decision.json`. Every acquisition writes an
immutable `rapidpy.susceptibility.acquisition.v1` record (holder records under
`<data_dir>/holder_susceptibility/`, sample records under the run's
`susceptibility_acquisitions/`, indexed in `artifact_index.json` with size and
SHA-256). Attach those files to the evidence record.

| # | Step | Expected |
|---|---|---|
| S1 | With the lift unpowered, open the Susceptibility Bridge dialog, press Connect, then Zero and Measure | Exact `Z`/`M` + CRLF in the transcript, CR-terminated numeric replies, scaled value shown; no axis moves |
| S2 | Disconnect the bridge cable and repeat S1 | Connection or reply error is reported verbatim; no value, no zero |
| S3 | Confirm `SCoilPos`, `SampleTop`, `SampleBottom` from the legacy INI and compute `Int(SCoilPos + (SampleTop - SampleBottom) / 2)` | Written expected target matches the record's `target_position`; target is on the coil side |
| S4 | With an empty rod, jog the lift to the computed target at speed index 0 and inspect clearance | Rod is centred in the coil with no contact; record measured position and settle error |
| S5 | Run a queue Holder command with the bridge enabled | Record order is home, zero, move, measure, home; holder `susceptibility_raw` equals the record's `bridge_scaled_value`; `susceptibility_evidence_id` resolves to the holder evidence file |
| S6 | Measure the Bartington reference standard as a sample (`SUSC`) | Value `(scaled - holder) * SusceptibilityMomentFactorCGS` agrees with the standard within the lab's written tolerance; VB6 side-by-side value recorded |
| S7 | Repeat S6 five times without re-measuring the holder | Spread and drift recorded; answers the holder re-measurement question in the decision record |
| S8 | Halt during the move to the coil | Run aborts, lift returns home, record outcome `failed` with `cancelled`; no `.rmg`, specimen, or `susceptibility.json` output |
| S9 | Unplug the bridge between zero and measure | Record shows the reply error, `safe_state_confirmed = true`, and no value; the prior holder (if a holder run) is unchanged |
| S10 | Block the home limit switch (controlled fault) after a read | `SAFE-STATE NOT CONFIRMED` error names both the acquisition and safe-return failures; operator inspection is required before continuing |
| S11 | Run the Bartington calibration scan (`frmCalRod.RunSusceSeq` equivalent) only as a separate calibration workflow | Never triggered by ordinary `SUSC` steps |

---

## 12. Sign-off matrix for the replacement claim

All rows must be `pass` with attached evidence before RapidPy may be described
as a replacement rather than transition/testing software.

| Area | Evidence required |
|---|---|
| Measurement sequence | Section 6 + H4 |
| Rejection and recovery | Section 7 |
| Motion and interlocks | Section 5 |
| Safe halt | M8, D3, D4 |
| Holder integrity | Section 8 |
| Output parity | H4, H5 |
| Restart behavior | Section 9 |
| VRM acquisition | Section 10 |
| Rockmag execution | Section 11 |
| Susceptibility acquisition | Section 11a |
| VB6 side-by-side | Requires readiness gates 2–4 closed |

---

## 13. First safe test to run when the system is connected

Run **R1 for the SQUID only**, with the changer, lift, and turning axes
unpowered:

1. Confirm the physical area is clear and the lift is at a known safe height.
2. Set `general.nocomm = false`, configure only the SQUID port, and leave the
   changer port unconfigured.
3. Start RapidPy and open the SQUID diagnostics dialog. Do **not** start a
   queue: preflight will correctly block on the unconfigured changer port.
4. Issue identity/status only, then one `CLP A` / `RC A` / latch / read cycle.
5. Save the transcript and the raw counter and DVM values for X, Y, and Z.

Expected: three finite readings with separate counter and DVM components, a
transcript containing exactly `ACLP`, `ARC`, `ALC`, `ALD`, `XSC`, `XSD`,
`YSC`, `YSD`, `ZSC`, `ZSD`, and no motion of any axis.

If any reading is empty, non-numeric, or arrives after the latch timeout, stop
and record it — that is transport evidence, and RapidPy is expected to raise
rather than return a value.
