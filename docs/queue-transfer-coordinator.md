# Native specimen transfer coordinator

`QueueTransferCoordinator` composes original `QueueLiftTransfer`, `QueueXYReference`,
`QueueFieldOutputs` and queue vacuum ownership on the same claimed worker. It requires
a live recorded XY reference, unchanged native motor/field profiles, an acknowledged
queue-owned pump and no active blank-holder rod. Recovery sessions cannot replay it.

`load(original_slot, sample_id, file_id=...)` rejects another active specimen before
I/O, verifies participating field cutoff, moves to the original specimen slot, picks
up gently, records fresh stopped pickup telemetry, connects grip, waits/references
the loaded lift and records measured height. It then parks the XY table at the
calibrated empty hole before handing off the lifted specimen context.

`return_specimen()` requires that original verified loaded context, verifies field
cutoff, raises the specimen at the empty hole, restores the original slot, lowers
to the measured-height supported dropoff position, records fresh stopped support
telemetry, releases the valve while retaining pump power, waits the accepted delay
and raises the rod to verified clearance. The parent journal remains owned throughout.

`QueueLiftTransfer.verify_vacuum_pose` stages all four independent stop/stable telemetry
checks before commands and publishes immutable evidence before returning a typed
`QueueVacuumPoseProof`. Pickup requires the exact recorded pickup count; dropoff
requires the original slot, measured-height target tolerance and unset top switch;
clearance requires the top switch and accepted zero envelope. The original vacuum
state is checked at both boundaries. Vacuum consumes this typed proof only while
its motion record is the latest verified stage, retains the entire pose evidence in
its plan and revalidates live field-off proof before changing outputs. Modified
records, a changed motor connection or any intervening stage invalidate pose authorization.

Failures never automatically replay transfer or release grip. A failed valve release
does not perform post-release rod motion; failed home/pose/publication retains the
original specimen identity and pending stage. Software support evidence comes from
calibrated position/stop/switch telemetry, not an independent specimen-support force
sensor. Physical pressure, contact and dropoff qualification remain required.

Injected native DLL/serial regressions exercise the complete load/return sequence,
height/empty-hole handoff, support evidence, valve/home/cutoff/publication failures,
stale/modified evidence, reconnect/recovery, changed calibration and repeated specimens.
This coordinator is not yet attached to MainWindow/MeasurementWorker. Startup pump/
empty-rod reference, blank-holder composition, measured backend binding, operator
interventions, terminal ownership/recovery and chain station transfers remain open
full-system integration requirements.
