# Per-file queue measurement sequences

The Sample Queue's Treatment Steps now supplies the executable sequence for its
sample file. Both `NRM → AF20 → SUSC` and `NRM -> AF20 -> SUSC` are parsed in order;
mixed arrow styles work, and empty steps are rejected. JSON/CSV and session row
persistence retain the same treatment text. This does not make unsupported treatment
labels executable: native worker treatment preflight still validates actuator routes
and accepted calibration before scientific operations.

QueueSample and each compiled Meas command carry a tuple of measurement labels.
Programmatic samples without explicit labels receive a snapshot of the loaded
sequence before compilation. The actual resolved step count determines dual-side
eligibility: VB6 frmChanger disables the second pass when a file has multiple steps.
One-step doBoth/doUp files retain their second-pass command labels. Rows from the
same file must agree on steps, count, initial orientation and doBoth; conflicting
metadata, boolean counters/holes and nonboolean orientation flags fail validation.
Explicit file sequences can run without a global sequence.

Native startup journals the resolved sample and command sequences. Before the
measurement panel acquires worker leases or constructs its worker, MainWindow
requires the current compiled Meas command and sample identity. Native handoff also
checks the loaded command identity, original pending root/token, full linked history
and the exact journaled command list. Later global-sequence changes cannot change a
compiled command. Command mutations block measurement and retain native ownership.
Reviewed rockmag/thermal routine identity must match every file sequence before queue
startup, and is checked again at handoff.

MeasurementWorker receives those file labels. The measurement step counter and run
timer use the active run's steps while preserving the separately loaded sequence.
The existing native acquisition/treatment geometry and worker publication ownership
checks still apply.

Regressions cover distinct files, second-pass retention, multi-step eligibility,
default snapshots, malformed/conflicting metadata, mixed/empty arrows, direct panel
worker construction, pre-lease rejected handoffs, runtime estimates and native
journal tampering. The panel construction test does not start its worker. Native
window tests inject DLL/serial interfaces and substitute scientific/holder execution;
these checks do not prove final scientific artifacts or physical station acceptance.

File-registry source-file identity and step-progress/eligibility counters, operator editing/import of initial
orientation/doBoth (the queue UI currently defaults up/single-side), final scientific
bundle acceptance, earlier-stage/restart recovery, chain station support, remaining
VB6 auxiliary capabilities, portable rebuilding, performance and physical/scientific
qualification remain required for full-system completion.

Currently the queue's Sample Set value is used as file_id, and adding from the
sample selector derives it from formation/location. That does not establish the
original .sam file identity. Correct registry-to-queue source identity must be
completed before claiming legacy file-registry parity.

Full-suite verification exposed Windows access denial when publishing a journal
during native startup. Safety-store reads now capture and close bytes under a
shared per-path local lock before parsing; atomic replacement uses that same lock.
Windows errors 5/32/33 retry only the identical fsynced temporary snapshot, at most
six replacement attempts with 0.31 seconds of total backoff. Other errors fail
immediately; persistent denial leaves the original pending state and propagates an
error. OS transaction/operation locks, checksums and full history verification are
unchanged. No hardware operation or stage creation is replayed by this retry.
Microsoft documents the effect of open-handle sharing on rename in
[CreateFileW](https://learn.microsoft.com/en-us/windows/win32/api/fileapi/nf-fileapi-createfilew).

Regressions pause a real observer read while a second store finishes the journal,
then verify the old pending snapshot and new verified record. Injected transient
Windows errors must retry the exact source/destination; persistent denial must
exhaust its bound, retain the pending token and remove its temporary file. The
native-window pause test allows 30 seconds for its combined startup/holder/load/
QThread-exit observation while preserving its original no-start/no-reload assertions;
production hardware deadlines are unchanged.
