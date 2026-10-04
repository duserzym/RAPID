# Native blank-holder geometry

Active VB6 `Sample.SampleHeight` defaults to zero, and `Measure_Read` uses a blank
holder correction. `modMeasure.bas` computes `Int(ZeroPos + SampleHeight / 2)` and
`Int(MeasPos + SampleHeight / 2)`; `modSusceptibility.bas` uses the same rule for the
coil. The empty rod therefore uses the configured positions without nominal
specimen-height shifts. It must not be treated as a pickup at a positive holder hole.

`QueueBlankHolderMotion.prepare` verifies an already positioned calibrated empty
hole under the original claimed queue. It requires live reference/field-off proof,
the original acknowledged pump ON/valve OFF binding, and no active specimen. Its
motion stage is persisted before independent four-axis Stop/readback and both XY
coordinate checks, including top-switch/lift-clearance checks. It performs no pickup,
grip command, table move or homing. The distinct `QueueHolderState` records zero
height, original queue token and `holder-NNN` identity in immutable linked evidence.
`latest_holder_context` preserves this context separately from specimen transfer
identity and makes pending contexts unverified.

`QueueHardwareBackend.bind_holder_geometry(session, vacuum)` borrows that verified
pose, original calibration/scientific settings and exact original vacuum owner.
Lost valve evidence, changed station bindings, an active specimen, wrong identity
or pending stage cannot authorize acquisition. Specimen geometry cannot overwrite
the blank binding. Blank geometry cannot authorize specimen field treatments.

Public SQUID/susceptibility reads use the durable acquisition API. Before instrument
I/O, blank acquisition stops/checks all axes and confirms the empty XY hole again;
it repeats those checks before handoff. SQUID audit must identify a holder block,
and its holder correction is blank. A bridge acquisition accepts zero height only
with the explicit `blank_holder=True` configuration and `is_holder=True` call.
Existing specimen acquisitions retain their nonzero-height requirement. Holder
bridge results retain scaled raw values without subtracting an earlier correction.

`measure_bound_holder` acquires the bridge first when enabled, publishes its holder
evidence, then acquires/reduces magnetic blocks and atomically installs a replacement
correction through `HolderMeasurementService`. Each block/recovery has its own durable
stage. Failed acquisition or reduction retains the prior correction. The exact owned
acquisition token may borrow already connected motor transports; it cannot reconnect
after a lost connection or bypass a foreign pending operation.

After acquisition, `QueueBlankHolderMotion.return_to_clearance` persists its motion
intent, verifies the original empty XY hole, raises the rod with the calibrated fast
lift tier and independently verifies motor stops/top-switch/clearance. It retains
pump power and valve OFF. `clear_holder_geometry` requires that same holder's verified
clear context before restoring the earlier specimen identity. Recovery does not replay
the return or release the outer queue lifetime.

The automatic coordinator still must establish the live reference, resolve/move to
the accepted empty hole, invoke these APIs under its worker claim, manage operator
tray inversion/per-file orientation, and implement original-panel stop/field/support
recovery and terminal release. The legacy `holder()` path is not replaced or enabled
by these APIs. These injected tests do not prove physical rod clearance, table origin,
vacuum pressure or scientific station acceptance; portable rebuilding remains required.
