# Verified measured specimen geometry

`QueueHardwareBackend.bind_specimen_geometry(session)` borrows immutable geometry
from the original verified loaded specimen. A live worker must hold the queue claim;
recovery and NoComm cannot borrow it. Motor routing, full controller calibration,
accepted XY home/map geometry and the fixed journal must match the original queue.
The root's acquisition profile is provided by `_acquisition_safety_profile`: helper
`rapid_main_queue_acquisition` plus immutable motion, SQUID, calibration and
susceptibility settings. Field stage profiles retain the usual actuator safety
profile. A configuration change cannot silently shift a loaded sample's targets.

The context must name the active specimen and a verified `lifted` phase, with positive
height inside accepted lift travel. Claim, original binding and current linked
transfer context are checked again when height is consumed. A returned, uncertain or
different specimen cannot reuse a cached bracketed acquisition or nominal fallback.
The backend's accepted config and native controller's nominal height stay unchanged.

Bracketed SQUID zero and measurement positions use `floor(base_position + height/2)`.
AF/ARM planning and execution, pulse positioning/execution, RRM planning/execution
and susceptibility configuration use the same measured height. RRM accepts the
explicit height without changing native spin-speed calibration. Susceptibility's
negative odd-height centre now uses floor, matching active VB6
`modSusceptibility.bas:Susceptibility_Measure` and VB6 `Int`; truncation toward zero
would otherwise shift the target by one count.

Native field-stage plans persist the specimen geometry before I/O. During that exact
owned pending stage, its matching immutable geometry can still be used for execution
and cleanup; another pending stage cannot borrow that exception. Measured queue
treatments require the original borrowed child store. Acquisition cannot start during
an unfinished field stage. Ordinary acquisition still needs its durable worker-stage
wrapper when the complete queue coordinator is wired.

`clear_specimen_geometry` discards the binding and bracketed cache only after the
same specimen has a verified `clear` return to its original slot. It does not release
the outer queue or prove vacuum OFF. Blank-holder geometry and operator tray inversion
still need their distinct coordinator phases; a blank is not a loaded specimen.

Regression tests obtain height from actual native pickup/top-offset commands with
injected serial responses, then inspect SQUID, AF/ARM, pulse, RRM and susceptibility
consumers. They check unchanged calibration, wrong/stale specimens, changed scientific
settings/home coordinates, original-store enforcement, pending-stage isolation and
negative odd-height rounding. Actuator services are injected for argument/evidence
checks; no physical instruments are actuated. The API remains to be invoked by the
automatic queue coordinator, with live reference/field/support verification, complete
acquisition ownership, original-panel recovery and physical/scientific qualification.
