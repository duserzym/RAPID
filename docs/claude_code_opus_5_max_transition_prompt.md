# Prompt for Claude Code Opus 5 Max

Copy everything below into Claude Code Opus 5 Max while it is opened at the root of this repository.

---

You are working in the RAPID instrument-control repository on the current
branch. Take ownership of completing and polishing the main
`RapidPy/rapid_main` application so it can become the practical replacement for
the legacy VB6 Paleomag application. The immediate goal has two equal parts:

1. redesign the main application with a coherent, production-quality
   glassmorphism interface; and
2. finish the remaining software functionality required for VB6 parity while
   keeping hardware-only validation gates explicit.

Treat this as laboratory instrument software. Do not fabricate hardware
evidence, silently fall back to simulation, actuate hardware without explicit
operator authorization, or allow a rejected measurement block to update holder
state or production output.

## Current-state rule

The detailed P0 phases later in this prompt describe safety requirements that
must remain true, but much of that implementation is now already present. The
current readiness assessment reports bracketed acquisition, coherent SQUID
evidence, X/Y/Z discontinuity rejection, recovery and retry exhaustion,
measured holder correction, holder-gated sample measurement, fail-closed
hardware mode, simulation isolation, transactional output, duplicate-free
resume, metadata resolution, and replay fixtures as code-complete. Verify those
claims against code and tests before changing them; do not blindly reimplement
working services.

Treat the repository and current tests as the source of truth. Correct stale
documentation and inherited test counts. Classify every finding as implemented,
intentionally retired, hardware-gated, or a real software gap. “Complete” does
not mean inventing physical-system acceptance evidence.

## Repository handoff checkpoint (2026-10-02)

Continue from the current branch; do not restart the redesign or recreate work
that is already committed. The following focused commits are the present
implementation baseline:

- `04cd361` — refocus this Claude handoff on main-app completion;
- `5969d17` — build the responsive glass main workspace;
- `5dae514` — complete main-panel actions and fallback plotting;
- `ccb900f` — use the acquisition clock for holder records;
- `3afc7f3` — make sequence editing stateful and atomic;
- `676e62c` — wire the truthful responsive dashboard shell;
- `cc61ffe` — align transition evidence with verified application state;
- `96fd2a7` — make measurement plots truthful by default;
- `6a17d50` — load real sample indexes into an empty queue;
- `df8c3ce` — isolate AF examples from live hardware; and
- `abebb65` — add atomic settings backup and restore;
- `a861d46` — refresh readiness and Claude handoff evidence; and
- `dcb402d` — add searchable VB6 transition help;
- `d784bdb` — update this handoff after transition help; and
- `0fd469b` — add the auditable calibration lifecycle and measurement linkage;
- `4a271e7` — record calibration lifecycle readiness; and
- `77f53c1` — package RapidPy for clean installation and startup diagnostics;
  and
- `13bad6b` — record the isolated package verification evidence; and
- `dd2e057` — finish the first glass hardware-dialog and accessibility slice;
  and
- `3a6ac66` — record the first glass-dialog readiness evidence; and
- `153fb63` — finish the glass guidance and transition-dialog slice;
- `786c36a` — record the guidance-dialog readiness evidence; and
- `d170a93` — finish the glass data/review-dialog slice; and
- `1536b5b` — finish the remaining glass dialogs and runtime fixes; and
- `87a0bb6` — record the completed reachable-dialog audit; and
- `dfa70ae` — add transactional plot exports and complete simulation-artifact isolation; and
- `0215d3b` — record export and simulation-isolation readiness; and
- `c0047ac` — reconcile quicklook publication with the run artifact index; and
- `be275be` — record quicklook artifact reconciliation; and
- `af2eee0` — integrate principal-axis and moment-decay review into Plots; and
- `27bd276` — record integrated analysis readiness; and
- `42edead` — carry exact production SQUID transport evidence into run bundles;
  and
- `6b40b81` — reject stale buffered and unterminated partial SQUID replies.
- `1ba9c6c` — recover SQUID transport faults by safely reacquiring whole blocks.
- `254fa71` — make live vacuum control fail closed and auditable.

The complete RapidPy suite now reports **499 passing tests** using
`python -m unittest discover -s tests -p 'test_*.py'` from
`RapidPy/rapid_main`. Preserve or increase that count, but treat the repository
and current test discovery as authoritative if later commits add tests.

The main shell, Dashboard, Sample Queue, Sequence editor, Live Measure panel,
Settings, Calibration Center, and reachable dialogs already share the
responsive glass design system. Dashboard cards are backed by real snapshots,
sequence import/save is atomic and dirty-state aware, the primary menu actions
are wired, shutdown ordering is controlled, and fallback plotting is present.
Verify these claims, fix regressions, and extend the system only where evidence
shows a remaining gap.

Work in this priority order:

1. extend raw communication evidence and deterministic fault handling to the
   remaining retained AF, IRM/ARM, and DC/ADWIN live adapters without touching
   hardware. SQUID stale/partial rejection and safe whole-block retry/backoff,
   plus vacuum connection/acknowledgment fail-closed behavior, are already
   complete; do not duplicate those implementations;
2. finish the remaining evidence-backed software parity rows and testable
   protocol/replay boundaries for thermal, AF, IRM/ARM, DC/ADWIN,
   susceptibility, VRM, rockmag, and transport robustness; keep vacuum physical
   pump/valve and independent pressure-instrument acceptance, and every other
   RAPID-system-only gate, explicitly pending;
3. complete deployment acceptance in a dependency-clean environment and on the
   operator account. The installable wheel, packaged icons, configuration and
   dependency diagnostics, installed helper entry points, source-tree-free
   main-window construction, and helper resolution are verified on this
   computer; a fresh dependency install or signed installer is not yet tested;
4. reconcile the parity/readiness documents only after the corresponding code
   and tests exist.

### Completed assignment: raw SQUID transport evidence in the run bundle

Commit `42edead` closes the first adapter-level gap listed for VB6
`modListenAndLog` parity. The production bracketed SQUID path now feeds exact
command/reply/error traffic into the run's `communication.tsv`. Do not recreate
this path; preserve and extend its contract.

The completed service path:

- adds injectable communication logging to `RawSquidTransport` without changing
  command ordering, parsing, timeouts, recovery behavior, or exception types;
- records exact TX commands for clear/reset, range selection, latch, counter
  reads, and DVM/data reads, plus exact RX counter and data replies. Include
  useful axis/latch context in event details, never fabricate an RX event for a
  command that has no reply, and log an ERROR event before re-raising an
  adapter failure unchanged;
- exposes an immutable snapshot of those events through `RawSquidTransport`,
  `BracketedSquidBackend`, and `QueueHardwareBackend` using a narrow public
  contract rather than reaching through private attributes;
- merges the backend events into the same per-run `communication.tsv` written by
  `MeasurementWorker`. Transcript publication is idempotent so cleanup or
  repeated finalization cannot duplicate events;
- preserves event timestamps, channel/port identity, direction, normalized
  payload, and context through the merge. Do not log credentials or unrelated
  configuration values;
- keeps explicit no-communication/simulation behavior truthful and isolated. A
  simulated transcript must not be presented as a live hardware trace; and
- was verified without opening ports, connecting to the RAPID system, moving
  axes, or issuing instrument commands.

Focused tests prove:

1. the exact TX/RX order and payloads for a successful coherent X/Y/Z
   latch/read cycle;
2. clear/reset, range, and latch commands appear once and in their real order;
3. adapter exceptions produce an ERROR event and propagate unchanged;
4. backend snapshots cannot be mutated by callers;
5. worker finalization places the adapter events in `communication.tsv` once,
   alongside worker lifecycle events; and
6. simulated/no-communication evidence remains explicitly labelled and cannot
   satisfy a live-transport assertion.

Coverage lives in `tests/test_squid_transport.py`,
`tests/test_communication_log.py`, and `tests/test_measurement_worker.py`. The
Commit `1ba9c6c` adds deterministic whole-block transport recovery: return to
verified zero, clear/reset, bounded exponential backoff, and complete
reacquisition from zero-before. It does not retry motion failures, returns no
block on exhaustion or recovery failure, removes the unsafe uncancellable
eight-second outer read timeout, warns the operator, and retains current-run
recovery records in provenance and `workflow_summary.json`. Do not duplicate
this state machine. Next, add equivalent raw evidence and deterministic fault
handling for other retained live adapters. Keep physical SQUID transcript and
fault-injection acceptance explicitly pending.

### Completed assignment: fail-closed and auditable vacuum control

Commit `254fa71` closes the vacuum software-truthfulness gap without claiming
physical-system acceptance. The legacy VB6 vacuum controller exposes pump and
valve commands but no pressure telemetry. RapidPy must preserve that boundary:
do not synthesize or model a live pressure value and do not treat a successful
serial connection as pressure evidence.

The completed vacuum path:

- fails hardware-mode construction when the dependency, port, or connection is
  unavailable instead of silently falling back to simulation;
- records exact controller TX, CR-terminated RX acknowledgment, and ERROR
  events, correlated to the command, through an immutable event snapshot;
- rejects empty or partial acknowledgments and does not report the requested
  pump state as confirmed after such a failure;
- reports pressure as unavailable in live command-only mode rather than using a
  modeled value;
- blocks queue readiness when a positive vacuum-pressure threshold is required,
  because this controller cannot prove that threshold. An explicit threshold of
  zero is the narrowly defined opt-out and is accepted only when pump-on state
  has been acknowledged; and
- makes current-run production vacuum events eligible for idempotent merge into
  the run's `communication.tsv`, alongside SQUID events. Per-source cursors are
  established when the worker is constructed so earlier traffic is excluded,
  and simulated sources cannot satisfy live evidence.

Coverage lives in `tests/test_vacuum_transport.py`,
`tests/test_diagnostic_services.py`, and `tests/test_measurement_worker.py`.
These tests use fakes and do not open a serial port or actuate equipment. Do not
redo this software path. Physical pump/valve command-and-acknowledgment
round-trip, an independent pressure adapter/interlock if pressure gating is
retained, and controlled physical fault injection remain explicitly pending.

The next implementation assignment is the remaining AF/IRM/DC evidence and
fault boundary. Inspect the VB6 commands and existing RapidPy adapters first,
then add exact, immutable, current-run TX/RX/ERROR evidence and fail-closed
connection/acknowledgment behavior where it is genuinely missing. Use injected
fake transports and deterministic replay tests; do not open ports or actuate
hardware. Do not copy SQUID's whole-block recovery state machine into treatment
devices unless the legacy protocol and safety contract specifically require
that behavior.

The truthful empty Plots/Sample Selection/Queue states, real `.sam`/`.csv`
sample-index-to-queue workflow, No-Communication-only AF examples, atomic
settings backup/validated restore/restart messaging, and searchable
“Where did this VB6 control go?” transition help are already implemented and
tested. Plots now distinguishes real measurement data from visibly labeled
example data, Sample Selection disables selection until a visible row is
chosen, and Webcam uses the correct `QtGui.QScreen` geometry contract while
remaining explicit when its optional dependency is unavailable. The
Calibration Center now also has immutable artifact snapshots,
versioned approvals, expiry and invalidation state, event-based rollback, hash
integrity checks, and active record IDs in measurement provenance. Do not
recreate these paths; hardware/reference-standard acceptance remains pending.
The `berkeley-rapidpy` wheel and seven console entry points are also implemented
and tested from an isolated install target. Do not restore repository-relative
launch assumptions.
Plot review also has operator-facing atomic JSON/CSV export with schema checks
and explicit real/simulated provenance. All simulated/replay core and auxiliary
artifacts publish under `SIMULATED`, and a live-declared backend returning a
simulated block is rejected before accepted state or output mutation. Preserve
this boundary; do not move sidecars back beside production output.
The Plots Analysis tab and exports now reuse `rapid_main.analysis` for
North/East/Down PCA direction, explained variance, RMS perpendicular residual,
moment statistics, decay ratios, monotonicity, and log-decay slope. Do not build
a parallel analysis implementation; remaining plot work must be tied to a
specific missing legacy surface or real-bundle acceptance evidence.

### Completed assignment: reachable dialogs

The completed implementation pass systematically audited and repaired every
dialog reachable from `rapid_main`; do not start another main-window redesign.
The first slice is complete in `dd2e057`: `dialogs/vacuum.py`,
`dialogs/irm_arm.py`, and `dialogs/squid_comm.py` now use shared glass surfaces,
semantic non-color-only states, accessible control metadata, compact form
wrapping/scrolling, and explicit unavailable behavior. The second slice is
complete in `153fb63`: Login, About, Quick Start, and Transition Help now use
the shared system with compact action visibility and accessibility coverage.
The third slice is complete in `d170a93`: Plots, Sample Selection, and Webcam
now use the shared dialog contract, expose accessible controls, retain compact
actions, communicate real versus simulated data in words and semantic state,
and fix Webcam's screen-type handling. The final slice is complete in
`1536b5b`: Debug Console, Step Monitor, and DC Motors now use the shared glass
and accessibility contracts; Debug Console uses the supported PySide6 rich-text
path and escapes backend text; Step Monitor uses `QtGui.QScreen`, normalized
progress, and written execution state; DC Motors labels simulated/unavailable
state in words, fails closed when unavailable, and refreshes controls after
disconnect. Do not redo this audit. Revisit a dialog only for a demonstrated
regression or a real backend/service gap.

The completed audit found and repaired dialog-local `setStyleSheet(...)`
conflicts, sparse accessible names/descriptions, color-dependent status labels,
and rigid geometry assumptions in the reachable set. Preserve the centralized
contracts; do not mechanically delete styles in unrelated standalone apps or
shrink safety-critical controls.

For any later dialog regression fix:

- reuse the centralized dialog title/subtitle, readout-card, and
  semantic-status contracts in `rapid_main/glass_theme.py`; extend them only
  when a remaining dialog demonstrates a reusable gap;
- use the existing semantic-state helper so visible text, the stable Qt
  property, accessibility metadata, and widget repolishing stay consistent;
- ensure states such as ready, warning, error, unavailable, and simulated are
  understandable without color alone;
- replace conflicting page-local styling in the audited dialogs with shared
  object names/properties while preserving device behavior and ownership;
- give actionable controls and live readouts useful accessible names and, where
  needed, descriptions;
- keep dialogs within the available screen geometry and add scrolling or
  internal table sizing when content cannot fit, rather than clipping controls;
- preserve clear destructive/safety affordances and deterministic reduced-motion
  behavior; and
- add targeted contract tests plus offscreen renders at compact and normal
  sizes. Inspect the rendered images; a successful `grab()` alone is not visual
  acceptance.

Do not mix this visual/accessibility pass with protocol invention or unapproved
hardware execution. If a dialog exposes incomplete live behavior, record the
software gap separately and address it through the shared service/backend layer,
not by faking a successful state in the widget.

Do not represent placeholder thermal estimates, simulated plots, demo samples,
or No-Communication results as hardware evidence. An explicit simulation tool
may remain when it is visibly marked, isolated from production outputs, and
disabled from live actuation.

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
- `RapidPy/rapid_main/rapid_main/glass_theme.py`
- `RapidPy/rapid_main/rapid_main/hardware_contracts.py`
- `RapidPy/rapid_main/rapid_main/queue_compiler.py`
- `RapidPy/rapid_main/rapid_main/panels/measurement.py`
- `RapidPy/updown_control/updown_control/app.py`
- `RapidPy/rapidpy_common/ui.py`
- every module under `RapidPy/rapid_main/rapid_main/panels/` and every dialog
  reachable from the main application
- the main-app UI, layout, bootstrap, workflow, orchestration,
  output-integrity, acquisition, holder, and parity tests

Preserve all current user changes and existing commits. Inspect `git status`, `git diff`, and recent history before editing. Never use destructive Git commands. Split your work into focused commits and list every commit in your handoff.

## Primary workstream: glassmorphism and operator usability

The UI now has a centralized, testable glass design system and a responsive
main workspace. Audit and finish it rather than replacing it or layering more
page-local CSS onto it. Preserve the existing token/helper direction unless a
verified defect requires a compatible change.

Use the existing Berkeley/RAPID identity as the anchor: deep maroon, restrained
gold, warm neutral foregrounds, and high-contrast semantic colors. The result
must feel like a serious scientific control surface, not a generic consumer
dashboard.

Implement:

- a subtle layered background with restrained gradients or aurora treatment;
- translucent glass surfaces for the header, navigation, cards, dialogs,
  tables, plots, and grouped controls;
- consistent surface opacity, borders, radii, elevation/shadows, spacing,
  typography, focus rings, and hover/pressed/checked/disabled states;
- reusable design tokens and helpers in the shared or main-app UI layer;
- a reliable opaque/translucent fallback where platform effects are unavailable;
- clear hierarchy for live values, units, warnings, queue state,
  hardware/simulation state, and destructive actions;
- responsive layouts with scroll behavior where content cannot fit;
- keyboard navigation, visible focus, reasonable hit targets, accessible
  contrast, and no color-only status communication;
- reduced-motion or non-animated presentation for safety-critical state changes.

Qt stylesheets do not provide true cross-platform backdrop blur. Do not claim
that they do. If native blur is added, isolate it behind a best-effort Windows
helper and keep a deterministic fallback. Prefer reliable layered painting,
alpha surfaces, and restrained shadows over fragile platform tricks.

Review the current compact-window constants critically. Keep every window
on-screen, including on small and mixed-DPI displays, but do not force the main
scientific workspace into a tiny phone-like width merely because an old test
encodes the current value. Update stale layout tests when the intended desktop
behavior changes. Navigation labels, charts, tables, and primary actions must
not clip.

Apply the design consistently to the application shell, Dashboard, Sample
Queue, Sequence, Live Measure, Settings, Calibration Center, and every dialog
reachable from the main application. Most of this pass is already committed;
look first for omissions, regressions, and dialogs that still use local styling.
Do not redesign unrelated standalone apps
unless a shared-theme change requires a compatibility fix; keep shared-theme
changes backward-compatible or create an explicit main-app variant.

For status communication, use both words/icons and the shared semantic visual
state. Never encode connected, safe, simulated, unavailable, warning, or fault
state only through a foreground/background color. Accessibility names should
identify the instrument and value/action rather than merely repeat a generic
widget class. Preserve logical tab order and make focus visible against every
glass surface.

Render representative offscreen screenshots at more than one window size and
inspect them. If native Windows capture is available, also inspect the real
rendered application. Do not approve the design only because the stylesheet
parses. Check contrast, clipping, scrolling, alignment, plot legibility,
keyboard focus, disabled states, and unmistakable safety/simulation status.

## Primary workstream: close real main-app functionality gaps

Build a traceability table from every visible action/control to the service it
invokes, the state it changes, its persistence/output, and its test. Search for
placeholders, disconnected or permanently disabled controls, `pass`, `TODO`,
`NotImplementedError`, broad exception-to-simulator fallbacks, and actions that
only update a label without invoking a real service. Then close the genuine
software gaps.

At minimum, verify and finish:

1. Dashboard diagnostics, truthful live/no-communication/simulation state,
   refresh evidence, errors, and recovery/configuration routes.
2. Queue add/edit/remove/reorder/import, preflight, holder commands, treatment
   and measurement orchestration, pause/resume/halt, interrupted-run choices,
   safe return, persistence, progress, and actionable failures.
3. Sequence validation, treatment-specific parameters, import/edit/save-as,
   dirty-state prompts, compilation, and mapping into queue commands.
4. Live readings and quality evidence, complete step history, retained
   scientific plots, exports/quicklook artifacts, and truthful empty/error
   states without dummy data outside explicit simulation.
5. Settings validation and persistence, VB6 INI/calibration import,
   unmapped-key reporting, backup/restore, and restart-required messaging.
6. Auditable calibration versions, operator/context metadata,
   validation/expiry/invalid state, import/export, non-destructive rollback,
   and measurement linkage. The append-only registry, UI lifecycle controls,
   artifact integrity checks, and measurement linkage are complete; verify
   them and add only evidence-backed gaps or required import/export formats.
7. Hardware-dialog control wiring, ownership/preflight, timeouts, telemetry,
   safe failures, communication evidence, and visible unavailable/simulated
   state.
8. Searchable “Where did this VB6 control go?” help, first-run guidance, and
   transition sheets for active workflows. The searchable task map and direct
   routing are complete; verify them and add only evidence-backed missing
   workflow guidance.
9. Login/authorization where retained, controlled shutdown, unsaved-work
   prompts, halt-before-exit, layout restore, logging, and useful diagnostics.
10. Clean-environment startup and packaging, resources/icons, configuration
    discovery, useful missing-dependency errors, and removal of source-tree-only
    assumptions. The wheel and isolated-target smoke test are complete; retain
    them and close only the remaining dependency-clean/operator deployment gate.

Do not make fake implementations merely to make controls appear wired. When a
feature requires unavailable physical hardware, complete its protocol boundary,
preflight, state machine, replay/simulator fixture, failure behavior, evidence
schema, and tests, then label live acceptance as pending.

Use `docs/vb6_parity_inventory.md` as the controlling inventory and reconcile
it with active items in `VB6/Paleomag v3.vbp`. Every item must end in one of
four evidence-backed states: implemented and software-verified; implemented but
pending named hardware acceptance; intentionally retired with an approved
replacement; or blocked with the exact dependency and next safe action.

Do not treat stale summary prose in `ROADMAP.md` as stronger evidence than the
current implementation, tests, readiness assessment, and parity inventory.
Update stale roadmap claims only after verifying the relevant code path; do not
recreate functionality merely because an older row still calls it a stub.

## Verified VB6 environment state on this computer

Re-verify the VB6 launch section in `docs/rapidpy_transition_readiness_2026-08-29.md` before changing it. It now records the English IDE repair; preserve that evidence and update only facts that you directly verify.

The following was verified directly on this computer without rebooting:

- English VB6 RTM `6.00.8176` is installed at `C:\Program Files (x86)\Microsoft Visual Studio\VB98\VB6.EXE`.
- The IDE opens normally to the English New Project dialog. The real `VB6/Paleomag v3.vbp` loads when VB6 is run elevated.
- DAO 3.5, `VB6.OLB`, `VB6EXT.OLB`, `MSDERUN.DLL`, `MSO97RT.DLL`, `MRT7ENU.DLL`, and the required Microsoft OCX files were restored or registered.
- Mixed Simplified Chinese VB6 templates, T-SQL components, wizards, Common Tools binaries, Designer satellites, and five system components were replaced with hash-matching English-media versions. The active VB6/Common/Designer trees had zero files reporting Chinese language metadata after repair.
- Displaced files are preserved at `C:\VB6-English-Repair-Backup-20260829-2`. Do not delete that backup.
- The Visual Studio setup front end that requested a reboot was terminated, and the machine was not rebooted.
- The Paleomag project must currently be launched as administrator to avoid `Error accessing the system registry` during design-time component loading.
- Visual Studio 6.0 Enterprise setup registered the product correctly, enabling
  `VB6.EXE /make`; `LINK.EXE` and `C2.EXE` are present.
- `C:\Windows\SysWOW64\MSCOMCTL.OCX` is the signed KB3096896 build
  `6.01.9846` and provides the project's required type library `2.2`.
- The authorized legacy archive supplied `vbSendMail_v3.0.dll` version
  `3.06.0005`, whose registered type library `5.7` matches the project.
- `VB6/Build-VB6Project.ps1 -NoFixups` successfully produced
  `build/vb6/PALEOMAG2013.exe` (1,736,704 bytes, 32-bit i386, version
  `3.01.0009`). Preserve `build/vb6/vb6-build-receipt.json` as evidence.
- The IDE remains RTM `6.00.8176`; the runtime is SP6-era `6.00.9848`. Base
  SP6 remains desirable for security fixes but is not a build blocker.

Do not regress or re-solve these closed gates. Re-run the repository preflight
and build only to confirm current state. The remaining VB6 gates are:

1. launch `PALEOMAG2013.exe` in no-communication mode against a copy of
   `VB6/Defaults.ini`, exercise the primary forms, and retain a smoke-test log;
2. run the signed physical RAPID-system procedure only after explicit operator
   authorization and area-safety confirmation.

Before changing VB6 source, re-run `VB6/Test-LaunchReadiness.ps1`, inspect the real project references, and distinguish IDE startup, project loading, and compilation as three separate gates. Do not install system software, edit registry permissions, actuate hardware, or reboot from Claude Code unless the user explicitly authorizes that action in the active session.

## Known safety defect and code already present

The archived defect is a discrete 2G flux-count step between the two zero readings that bracket four holder orientations. Observed one-count increments are approximately X `0.090`, Y `0.106`, and Z `0.066` in calibrated raw units. Linear interpolation across the step creates a monotonic ramp; rotating the four contaminated readings into the holder frame creates a false pattern separated by roughly 90 degrees.

The repository already contains:

- VB6 categorical rejection/re-latch/re-zero/retry logic in `VB6/modMeasure.bas`;
- RapidPy `BracketedMeasurementBlock`, `validate_zero_pair`, `reduce_bracketed_measurement`, and `FluxCountDiscontinuityError`;
- worker logic that invokes `recover_flux_count_discontinuity` only when a backend explicitly implements it;
- unit tests for the archived X step, valid reduction, monotonic staircase rejection, and recovery-hook ordering.

Do not weaken those rules or reintroduce “accept after N retries.” A discontinuous block must always be discarded. Verify that tests cover X, Y, and Z steps, down-orientation transforms, holder subtraction, retry exhaustion, no recovery hook, and proof that output/holder state is unchanged after rejection; add tests only for coverage that is genuinely missing.

## Existing P0 specification: preserve and verify the production bracketed SQUID path

Use the requirements below as an acceptance specification. First inspect the current acquisition service, transport evidence, recovery path, and tests. Preserve working implementations and implement only requirements that evidence shows are missing.

The production measurement backend must use an explicit, testable acquisition state machine that acquires:

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

Verify that the live transport/evidence model retains the atomic latch sequence, separate counter and DVM readings, raw replies, timestamps, and range. Reject partial, stale, malformed, mismatched, timed-out, or non-finite observations. Ensure all three axes belong to one documented latch/read cycle.

Verify the live backend implements `recover_flux_count_discontinuity`. It must move to a safe zero state, clear/re-latch/reset the relevant 2G path using the verified hardware commands, wait for stabilization, record the recovery, and restart the entire block. On retry exhaustion or recovery failure, halt through the shared safe-state path and produce a clear operator error. Never synthesize a successful result.

## Existing P0 specification: preserve and verify measured holder state

Verify that the queue implements `Holder` as a complete measured-state command, not merely changer motion. Where evidence is missing, complete the path so it can:

- acquire and validate a full bracketed block;
- persist holder ID/location, direction, four raw and adjusted vectors, zeros, range/calibration, timestamps, software version, and validation evidence;
- install the new correction atomically only after success;
- keep the previous correction on rejection, abort, timeout, or write failure;
- expose holder magnitude, induced/asymmetry metrics, age, identity, and validity in the operator UI;
- associate the exact holder record/version with each sample result;
- prevent sample measurement if required holder state is absent, stale, direction-incompatible, or otherwise invalid.

Add replay fixtures representing stable holders and the archived failure. Test queue-level behavior, not just math helpers.

## Existing P0 specification: preserve and verify fail-closed hardware mode

Audit every backend factory and `QueueHardwareBackend`. In explicit simulation/no-communication mode, simulated backends are allowed and must be visibly labeled. In hardware mode, any missing package, driver, port, calibration, adapter, or connection must fail preflight and prevent queue start. Remove broad exception-to-simulator fallback in hardware mode. Preserve diagnostic details and a safe operator message.

Guarantee that simulated results:

- carry an unmistakable simulation marker in memory, logs, UI, and files;
- cannot be written to production output paths unless explicitly enabled for a test;
- cannot be mistaken for hardware readiness evidence.

Add tests for missing imports, unavailable ports, constructor errors, failed identity checks, and mixed live/sim configurations.

## Existing P0 specification: preserve and verify metadata and persistence

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

## Remaining live acceptance work

After the P0 path is complete in software, work through the P1/P2 inventory in
`docs/rapidpy_transition_readiness_2026-08-29.md`: AF, IRM/ARM,
ADWIN/DAC/MCC, thermal, susceptibility, VRM, rockmag, motors, serial
robustness, installer/config migration, and physical acceptance. Vacuum
software behavior is complete as described above; its remaining work is
physical pump/valve acknowledgment and, if pressure gating is retained,
integration and acceptance of an independent pressure source. Plot/export,
settings backup/restore, and rollback software paths are already implemented;
revisit them only for a demonstrated regression or a specific acceptance gap.

For work that requires the physical RAPID system, provide an executable acceptance procedure and evidence schema instead of claiming success. Every hardware test record should include date/time, operator, machine/software commit, configuration hash, device/port identity, commands and raw replies, expected result, observed result, pass/fail, and artifact paths.

## Required verification

Run the smallest relevant tests while iterating, then the full RapidPy suite. Add type/lint checks already supported by the repository. Perform static checks on VB6 changes and re-run the existing no-fixup build to detect regressions. The compile gate is closed; the no-communication runtime smoke test is not. Do not connect or actuate hardware unless the operator explicitly authorizes it and confirms the physical area is safe.

Before finishing:

1. show `git status` and ensure no generated test/config artifacts are included;
2. run `git diff --check`;
3. update the readiness assessment with code-complete versus hardware-validated status;
4. produce focused commits, with tests paired with their implementation;
5. report exactly what is complete, what remains, tests and counts, commits, hardware evidence obtained, and the first safe test to run when the system is connected.

For the dialog pass, also provide a compact matrix with one row per reachable
dialog: launch route, shared-theme status, compact-size result, keyboard/focus
result, accessibility/status result, backend truthfulness, tests, and any named
hardware-only gate. Include paths to representative compact and normal
screenshots, but do not commit disposable captures unless the repository has an
intentional visual-baseline location.

Also report the VB6 gates separately in the final handoff:

- English IDE startup: verified or regressed;
- real `Paleomag v3.vbp` load: verified or blocked, with the exact reference/dialog;
- compiler edition and `/make`: verified or blocked;
- no-communication runtime smoke test: verified or blocked;
- physical RAPID-system test: not attempted unless explicitly authorized, otherwise include the exact safe procedure.

Acceptance for the replacement claim requires all P0 items to be code-complete and pass replay/fault tests, plus signed physical-system evidence for measurement sequence, rejection/recovery, motion/interlocks, safe halt, holder integrity, output parity, and restart behavior. Until then, label RapidPy as transition/testing software, not a full VB6 replacement.

---
