# Original scientific instruments in a native queue

Native startup now prepares a QueueInstrumentLifetime before creating the durable
root or opening ports. It captures the actual native SQUID adapter/reader/raw client
and, when configured, the native susceptibility adapter/client. The root records a
unique owner identity, both accepted instrument configurations and bridge presence.
An explicitly disabled bridge may retain the UI's unavailable placeholder; an
enabled bridge may not use that placeholder or a simulator. Existing scientific
serial handles prevent a new queue. Motor, vacuum and scientific serial circuits
must be distinct, including Windows device-prefix/case aliases.

Object creation is pure. The owner binds to the original claimed queue and checks
the full journal history, current configuration and original adapter/client identities
before transfer or instrument I/O. A first scientific serial handle is adopted only
inside its original pending acquisition stage and token. Its actual port, baud,
parity, data bits, stop bits and open flag must match. A mismatched handle is retained
for recovery rather than accepted or silently discarded. A later replacement or
unowned connection blocks the original queue. Acquisition evidence links the
instrument owner profile and validates the objects again during settlement.

Terminal cleanup still requires field cutoff, specimen return, stopped clear rod
and acknowledged vacuum OFF. Before any transport close, its immutable pending
stage now includes the original instrument profile and whether each instrument
ever acquired a serial handle. Vacuum, motors, SQUID and susceptibility receive
independent close attempts. An instrument close failure or a port left open retains
the exact original adapter/client/serial handle and leaves the stage pending.
Retries close only the unsettled original handles; they do not reconnect, read,
reset counters, move the rod, change vacuum outputs or reapply field treatments.

Both scientific clients and their retained original serial handles must be settled
before close-stage verification and root publication. Root/publication retries use
the exact original close evidence and instrument profile. MainWindow also checks
the retained startup owner's settled objects after actual worker exit before OS
and logical lease release. Queue reentry rejects replaced scientific objects.

Injected tests cover both native adapters with serial handles, pending-before-close
journaling, independent failure/close attempts, same-handle retries, changed objects
and settings, unstaged connections, wrong actual serial settings, root publication retries,
absent/enabled bridge handling, pre-start retained handles and serial collisions.
Connection-only fixture stages do not claim scientific specimen acceptance. Existing
MainWindow tests substitute the scientific panel/holder run as documented in
native-main-queue.md; these checks do not prove a complete physical station run.

MainWindow's Diagnostics menu provides **Retry Queue Shutdown** after a failed
terminal close. It requires the original queue, all six retained device leases,
the exact immutable close-stage token/plan and full verified journal history.
It starts a claimed worker only for unsettled original closes or publication.
An active queue, live worker, or unfinished field/motion/acquisition stage is
rejected. Rebinding the original backend preserves its terminal/tray services
and close token, including after some original transports have already closed.
Successful retry releases ownership only after actual worker exit and the same
verified root and instrument settlement checks. This is live terminal recovery;
restart recovery and recovery before a close stage remain unfinished.

Failed returned-handle opens and SQUID buffer initialization failures attempt
cleanup immediately. Failed cleanup retains the original handle and blocks reads
and replacement until its close succeeds. Failed diagnostic baseline acquisition
invalidates an earlier baseline. A serial constructor that raises without returning
a handle remains responsible for its own internal resource cleanup.

Original-panel recovery before terminal close and restart recovery, complete
per-file scientific step/eligibility/artifact acceptance, optional physically absent
field participation, chain transfer, remaining VB6 auxiliary capabilities, portable
rebuilding, performance gates and physical/scientific qualification remain required.
