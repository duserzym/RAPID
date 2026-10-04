# Claimed XY table transfer motion

The legacy station stores a separate X/Y controller-coordinate pair for each slot.
`import_vb6_ini` now imports `UseXYTableAPS`, `XYHomeX/Y` and `XY<n>X/Y` into
`motor_station`. Missing, fractional, nonfinite and out-of-range coordinate pairs
are omitted with warnings. A station reimport clears earlier native wiring and
calibration before loading the new file; it cannot synthesize a complete station
by mixing values from different files.

`QueueStationGeometry.from_config` requires accepted motor calibration, an explicit
matching station mode, and XY home/map entries for XY operation. It snapshots the
coordinate map as immutable tuples. `xy_target` rejects unregistered locations.
`slot_from_xy_counts` checks both signed controller readbacks within the legacy
0.02-slot tolerance derived from accepted `OneStep`; it requires exactly one match.
There is no fallback to SlotMin and no chain-coordinate proof for an XY slot.
`verify_empty_xy_readback` additionally requires the calibrated empty location.

`QueueXYTableMotion` requires the actual routed native motor client and all four
registered axes. Its profile includes the full controller calibration, immutable
geometry and original axis ports/addresses. A worker must own the original
`QueueWorkflowSession.claim`; its root profile must bind this exact motion profile.
The caller must establish the live XY reference and verify field outputs off.
Imported home coordinates alone do not establish a live reference.

`QueueXYReference` now supplies a recorded, bounded native two-edge reference;
see `queue-xy-reference.md`. Table transfers can borrow its typed state and persist
the exact original reference ID/connection/anchor proof. Automatic coordinator
wiring must use this proof; existing explicit operator-attestation callers retain
their strict Boolean contract.

Before native I/O, `move_to_slot` persists a motion stage and specimen identity.
It independently issues Stop and checks two position/velocity-register samples on
all four motors. Lift position must be within the calibrated clearance envelope and
the top-reference switch must be set. It commands absolute X and Y moves with native
XY acceleration on their own ports, waits cooperatively, then independently stops
and reads all axes again. Both final coordinates must identify the requested slot;
lift clearance remains required. It never homes, changes vacuum or rotates the rod.

Missing stop replies remain failures even if fallback Halt and zero velocity appear
to succeed. A failed Y command still stops X and every other axis. Cancellation
after the X command prevents starting Y and retains a pending stage. Evidence is
published through the existing immutable queue event chain; a publication failure
leaves the original stage pending even after apparently successful physical motion.
Recovery sessions cannot replay transfer motion. They need the original coordinator's
stop/support/field recovery and terminal vacuum-release path, which remains open work.

Tests use the actual routed controller and injected serial position/status responses.
They verify per-port native commands, before-I/O journaling, both-axis slop checks,
clearance, failed acknowledgements, cancellation, publication failure and no replay.
No instruments are actuated. This worker path remains to be wired into the complete
pickup/grip/home/empty-hole/read/return/drop/release lifecycle and qualified on the
physical station before automatic live execution.
