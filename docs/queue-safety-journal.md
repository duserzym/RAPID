# Queue lifetime and subordinate stage journal

`rapidpy_common.queue_safety` supplies the durable ownership model required for a
live queue which holds vacuum across acquisition and treatments. A queue persists
its original plan, run identity and fixed stage wiring/calibration profiles before
native I/O. Its OS lifetime lease excludes other helpers and processes. A worker
must claim that session before borrowing the existing treatment safety interface;
competing workers cannot borrow it or release its live owner.

AF, ARM, pulse IRM, RRM, motion, acquisition and vacuum are supported stage families.
Each stage persists a new identity and detached plan before it begins. Failed or
simulated stage completion remains pending and prevents another stage. A successful
treatment verifies that stage while the outer queue remains pending. Existing main
treatment journal hooks can use the stage adapter without reacquiring the OS lease
or clearing the queue. Outside the worker claim, the adapter reports the outer
queue fault so ordinary controls remain inhibited.

A native acknowledged vacuum enable can publish a `held` stage with raw replies.
It explicitly reports `safe_state_confirmed=false`; it permits subsequent stages
only while the queue retains ownership. Missing, simulated or failed acknowledgements
remain pending. A failed release cannot be disguised as a new held enable. A hold
record proves commanded state, not pressure or specimen grip. The queue driver must
still apply the configured pressure/command-only acceptance gate.

Every finished stage writes a separate immutable event, flushes it and links its
SHA-256 identity before updating the live journal. The journal retains one latest
stage and history head. The newest event is checked before beginning the next stage;
all linked events are checked during recovery and final queue completion. A failed
journal update after event publication leaves an unreferenced event and the original
pending stage; it cannot clear ownership. Missing, altered, misidentified or reordered
evidence prevents verified completion. File identities are restricted to generated
hexadecimal UUIDs and cannot select arbitrary paths.

The queue release record has its own schema with independent vacuum-off, motor-stop
and field-output-off booleans and cleanup errors. A treatment-success record cannot
serve as queue release evidence. All three checks must be physically verified,
simulation must be false, and cleanup errors must be absent before the parent clears.
Generic treatment store completion cannot clear the parent queue token.

Recovery obtains the OS lease and checks the exact original queue profile and full
event chain. A pending stage retains its family, token, specimen, run and plan for
stop/discharge/output recovery, without starting or replaying a new stage. A process
crash automatically releases the OS lease while retaining that durable identity.
Treatment-only main recovery refuses an outer queue and directs it to its coordinator.

Validation uses isolated journals, injected records and the actual native backend's
treatment journal hooks. Tests cover cross-owner exclusion, worker handoff, premature
release, strict booleans, vacuum holds, failed releases, publication failures, exact
station recovery, immutable history and a child-process crash. No instruments are
actuated.

This is the journal/coordinator primitive, not completed operator workflow wiring.
The main queue still needs a worker-owned vacuum start/release path, durable ordinary
motion/acquisition stage records and original-panel coordinated recovery. Integrate
those paths while preserving the legacy sample-transfer order and independent
treatment cleanup. Physical pressure/gripper, clearance and scientific acceptance,
remaining auxiliary tooling and a fresh portable build remain required.

The main native vacuum adapter now supplies a claimed queue transport path; see
`queue-vacuum-binding.md`. The transfer audit in `queue-transfer-audit.md` identifies
the distinct pump/valve phases, blank-hole resolution and operator tray flipping
which must be restored before enabling live queue startup.
