# Queue transfer parity audit

The active VB6 sources define transfer and operator phases which the current Python
queue still needs to restore before live automatic startup. The worker and durable
queue journal are prerequisites; they do not prove transfer parity.

## Source evidence and required phases

`VB6/SampleCommand.cls:Execute` powers the pump for Holder/Meas commands and dispatches
specimens through `VB6/modChanger.bas:Changer_ProcessSample`. For a specimen, that
routine moves to the registered slot, discharges axial IRM before loading in the
enabled rockmag case, performs the gentle pickup and sample-height calculation,
connects the vacuum valve, settles for 0.3 seconds, lifts/home-references the sample,
then moves the table to its nearest empty hole before treatment/read. Returning a
specimen raises to zero, moves back to the original slot, performs dropoff, disconnects
the valve, waits the configured dropoff delay, then lifts clear. Pump and valve
lifetimes are distinct; requiring the gripper valve ON across every queue command
cannot represent this sequence.

The holder path processes an empty changer hole and measures the blank rod. It does
not pick up the specimen occupying the most recently measured slot. Periodic holder
markers now use the station-resolved empty-hole sentinel zero; the native coordinator
still must resolve and verify the actual empty location before lowering. The existing
native positive-hole holder path and current Goto/Meas dispatch do not establish that
blank-holder/transfer invariant.

`VB6/SampleCommand.cls:Execute` treats InitUp as a preprocessing marker. Its Flip
case moves to a tray-access position, asks the operator to place the selected samples
with arrows down, then changes that file's orientation. A 180-degree rod rotation
does not implement tray inversion. Restore operator confirmation, per-file orientation
and access-position/vacuum phases; do not enable a live queue using the current native
`flip` implementation. Existing native InitUp homing also needs to be reconciled with
the marker semantics.

`VB6/SampleCommands.cls:Preprocess` appends Flip/Holder and repeated Meas commands
to the original collection. It does not insert a Flip before the original first pass
or duplicate each specimen read adjacently. Python preprocessing now preserves the
whole original pass, appends those second-pass commands in the legacy order, retains
periodic holder checks and uses the appropriate XY/non-XY flip marker. Repeat commands
are independent copies. `test_queue_pass_order.py` checks complete command traces,
multiple files, periodic blanks, an existing Flip and final return placement.

## Remaining integration work

`rapid_main.queue_station.QueueStationGeometry` now resolves blank-holder markers
from the full accepted motor calibration and imported `HoleSlotNum`. XY stations
have one explicit empty slot; chain stations use multiples of `HoleSlotNum`, circular
distance and the legacy upper-on-tie rule. Unsupported chain origins fail validation.
The resolver rejects specimen/blank confusion. Chain readbacks use counts per slot;
XY readbacks require both coordinates from the imported `XYTable.XY<n>X/Y` map.
A single chain count can never prove an XY slot. Unmatched or ambiguous coordinates
fail validation, including a large Y mismatch which the legacy asymmetric comparison
could mistakenly accept. Home and slot coordinates use strict signed integer counts.
New station imports clear earlier native calibration, wiring and geometry so missing
entries cannot silently inherit another station's settings.

`rapid_main.queue_table_motion.QueueXYTableMotion` now supplies claimed native XY
motion. It binds all four original ports/addresses, full motor calibration and the
coordinate map into the queue motion profile. It persists the stage before any
commands, checks motor stops and live lift top-switch/clearance readbacks, commands
the independent X/Y targets with native absolute motion, then checks both final
positions and stopped telemetry. Stop checks run independently for all four motors
after a failed axis command or cancellation. Failed checks/publication remain pending;
restart recovery cannot replay motion. See `queue-xy-motion.md` for its boundaries.

The resolver and claimed motion path are not yet wired into the automatic transfer
coordinator. That coordinator must establish a fresh XY home reference, verify field
outputs off and persist specimen transfer state; aligned coordinates alone do not
prove specimen support or safe grip release.

`rapid_main.queue_lift_transfer.QueueLiftTransfer` now implements claimed pickup,
loaded top-reference, raise-for-return, supported original-slot dropoff and
post-release clearance phases using the actual routed motor client. Typed specimen
identity, original slot, raw pickup position and measured height are persisted before
motion and in linked phase evidence. The height calculation subtracts the live home
offset from pickup position minus SampleBottom; dropoff uses that measured height
without mutating the accepted motor calibration. A pending later stage makes the
recovered context unverified while retaining original identity and geometry.

Grip must be acknowledged before loaded homing/return/dropoff; original-slot and
stopped-motor readbacks must verify support before valve release is authorized by the
coordinator. Post-release lift clearance requires acknowledged valve OFF and the
accepted DropoffVacuumDelay. Native failure, cancellation and failed publication
preserve the original journal/outputs. Recovery never replays transfer. See
`queue-lift-transfer.md` for phase requirements and remaining integration work.

`QueueHardwareBackend.bind_specimen_geometry` now consumes the verified loaded
context under the original worker claim. SQUID zero/measurement, AF/ARM, pulse IRM,
RRM and susceptibility targets use measured height while baseline calibration remains
unchanged. The binding includes the original scientific settings and XY home pair.
Wrong/returned specimens, changed settings and unrelated pending stages fail closed;
own field stages retain immutable geometry in their durable plans. Acquisition cannot
borrow an unfinished field stage. See `queue-specimen-geometry.md`. The coordinator
still must invoke this binding and stage remaining transfer I/O. Public bound-specimen
SQUID/susceptibility reads and flux recovery now stage ordinary acquisition under the
original claim, validate measurement evidence and independently verify all four motor
stops before publication/handoff. Failures retain the journal and specimen grip. See
`queue-acquisition.md`; the worker/coordinator integration remains open.

Wire the separate pump/valve phases under the durable parent queue and record each
remaining native motion/acquisition/transfer stage. Wire accepted empty-hole resolution and
fresh readback verification into those stages, persist the specimen's original slot
and transfer state, and verify clearance/stop/support before changing vacuum. Add the original-panel recovery path
which preserves grip when specimen support or field/motion safety is unknown. Wrap
acquisition and treatment workers in the queue claim without moving hardware waits
back onto the GUI thread. Restore initial loading and Flip operator interventions
and file-specific orientation before accepting the second pass.

These findings are explicit open full-system requirements. Unit tests and simulated
command traces do not qualify physical clearance, pressure, grip, sample transfer or
scientific station acceptance. The portable pilot predates these changes and must be
rebuilt after integration.
