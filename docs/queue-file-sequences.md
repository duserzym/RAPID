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

File-registry step-progress/eligibility counters, final scientific
bundle acceptance, earlier-stage/restart recovery, chain station support, remaining
VB6 auxiliary capabilities, portable rebuilding, performance and physical/scientific
qualification remain required for full-system completion.

Index-backed queue rows now preserve their absolute source path independently of
the display Sample Set. Index File, Orientation and Both Sides are editable table
columns, also available when adding a sample. File settings in the row context menu
applies orientation settings to every row from that index. JSON/CSV/session rows
preserve these fields; old rows remain compatible as manually grouped rows with
default Up/single-side settings. The VB6 convention remains: a one-step file starting
Up can add a Down pass; starting Down runs only Down, and multiple steps suppress
the second pass.

SAM/CSV registrations carry the actual resolved source path. Queue startup requires
an absolute supported index, matching canonical identity and exactly one matching
specimen entry. It captures original index bytes SHA-256 and detached registrations
before any root, device lease or operator startup confirmation. Native roots link
this snapshot. Handoff verifies original source metadata against the journal and
rejects changed index bytes before acquiring measurement worker leases.

Rows receive a persistent identity when compiled. It survives JSON/CSV/session
restore and stays on duplicated second-side commands. Running/Done/Error updates
target that original row, so equal specimen names from different files remain
independent. Old programmatic commands without a row identity match file, slot and
name. Duplicate or malformed persisted identities are rejected, and completion for
another specimen fails the current queue rather than advancing its plan.

The panel resolves metadata and specimen headers from that index's directory and
captured registrations, rather than another selected index or the configured output
directory. Index-backed outputs have separate source-specific folders, so identical
specimen names from different indexes do not overwrite each other's bundles. File
identity is case-normalized on Windows, while actual resolved paths are preserved
for reads. Manually grouped rows without a source retain their prior metadata path;
this does not invent source-index provenance for old queues.

This proves source-index routing and integrity, not full legacy registry or final
scientific acceptance. Registry progress
counters, standalone selection provenance, complete scientific bundle acceptance,
earlier-stage/restart recovery and the other qualification gates remain required.

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

Index-backed queues also capture specimen metadata before native root creation.
The snapshot contains header selection, complete resolved metadata, registration,
defaulted fields and SHA-256 for each candidate source (or explicit absence).
Changes while preparing the snapshot fail admission. The original journal links
these values; handoff rejects changed/removed header files, newly created headers,
and edited in-memory metadata. Workers receive detached frozen metadata rather
than resolving headers again at measurement time. The digest conservatively covers
the whole source specimen file, including any measurement records after its header.
Manually grouped rows keep their existing resolution behavior.

This closes the queue's header snapshot/routing gap. Final scientific artifact
provenance/acceptance, standalone selection provenance, path containment, registry
progress/eligibility/.UP handling and all remaining qualification gates still apply.

The original captured scientific source is now exported into completed-run
provenance.json and every run's artifact_index.json (including aborted runs).
It carries source index path/SHA-256, compiled file/row identity and the complete
frozen specimen snapshot. The MainWindow verifies the original journal and source
bytes; the panel checks provenance/metadata equality before worker leases, and the
worker copies both inputs. Mismatched or nonfinite source metadata is rejected.
Completed-run provenance.json is a required manifest artifact with its own SHA-256.
Aborted runs retain source context without publishing the scientific bundle.
Manually grouped rows have no fabricated source provenance.

These links support source traceability; full scientific acceptance still requires
per-step/file progress and .UP combination/eligibility, artifact checks across real
formats, standalone provenance, containment and all other completion gates.

Specimen header candidates and output destinations now resolve within their
selected directory. Relative nested specimen names remain supported; absolute,
drive-relative, escaped, device/stream and reserved artifact names are rejected.
Queue validation occurs before native root/operator startup; panel resolution and
output admission precede device leases, and worker construction checks paths before
preflight. Captured metadata records its source root and rejects changed source
parents, including identical bytes redirected through an external junction.

Bundle/staging/sidecar/JSON-temporary output paths reject child links or junctions
and shared hard links, and publication rechecks destinations. Copied bundle
metadata cannot be redirected by caller edits. Nested legacy/RMG parent directories
are created within the admitted root, and resume copies preserve those paths.
These are admission/use checks, not an atomic defense against arbitrary concurrent
filesystem mutation. Complete scientific acceptance, registry/.UP/eligibility,
standalone provenance, recovery, portable/performance and physical gates remain.

Legacy `.UP` interchange now has a strict reader/writer in `io/legacy_up.py`,
derived from `VB6/Sample.cls` (`WriteUpMeasurements`/`ReadUpMeasurements`). A run
contains ten rows per block: two Z baselines, four S sample vectors, four H holder
vectors, plus specimen, direction, run block count, block/reading number and optional
row timestamp. Repeated specimens remain separate runs; exact specimen lookup
selects the latest complete Up run. Truncated or corrupt trailing runs fail closed
rather than falling back to older data. Seven-column records retain missing times.
The writer emits CRLF, ordered rows and seven-decimal scientific notation.

The file has no treatment, calibration or acquisition audit fields. Converting its
raw readings requires explicit external range/axis calibration and creates no
live observations or hardware audit. A full acquisition export must retain its
audit separately. VB6 places this file at the current-step path, not one filename
for all treatments. Queue wiring, durable per-step identity/eligibility, AvgSteps,
Up/Down assimilation/statistics and final scientific publication remain required;
the interchange checkpoint alone does not fix paired-run overwriting.

Repeated bracketed readings now produce collection statistics over all four
positions in every block. Saved error angle and the panel CSD/drift/holder/induced
values use this collection instead of a zero/last-block result. Component sample
SD, mean raw/calibrated vectors, signal ratios and directional subsets are retained
in provenance. Blocks with changed calibration/acquisition context and cycles that
mix structured blocks with unstructured tuples are rejected before publication.
Tuple-only steps clear earlier block quality and preserve their existing behavior.
The explicit historical exact-alignment Fischer convention still requires
scientific acceptance. This supports current repeated cycles; durable per-step
Up artifacts, queue AvgSteps and cross-run Up/Down assimilation remain unfinished.
