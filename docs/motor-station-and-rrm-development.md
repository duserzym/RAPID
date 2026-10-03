# Motor station routing and rotational remanence

Source checkpoint, 3 October 2026. These are software verification results;
physical motion, field calibration and scientific acceptance remain required.

## Station migration

The Berkeley VB6 profile assigns address 16 to four controllers on separate
ports: changer X COM3, changer Y COM4, lift COM5 and turning COM6. RapidPy now
persists these bindings in `motor_station` and routes logical axis operations
through their actual serial connections. Shared ports open once; duplicate
address/port pairs are rejected. Request/reply and composite operations use a
reentrant lock. Raw commands cannot infer a port in the routed client.
Connection failure closes ports already opened. Communication evidence includes
the actual port. Construction opens no serial connections.

The queue and manual motor diagnostics both use imported native calibration:
turning -8000 counts/turn and 16000000 native velocity units/rps, lift speeds
25000000/35000000/35000000 and acceleration 50000, changer speed 10000000,
slot limits 1..100 and fractional one-step -1010.1010101. Transport is
57600,N,8,2, as specified by `frmDCMotors`. Lift coordinates use current motion
settings. Fractional negative specimen-centre positions use VB6 floor semantics.
Slot number 46 is retained as a slot setting, rather than a lift-speed percentage.
Pickup torque percentages convert using the configured native maximum torque
(28% of 32000 produces 8960, rather than sending 28).

Incomplete or invalid imported native calibration blocks queue preflight and
leaves its motor transport unavailable. The Settings and manual motor panels
show imported wiring; percentage speed and single-port controls cannot silently
override the imported native station profile.

## RRM

Executable labels carry both AF field and signed spin speed:

- `RRM50/5`: transverse 50 mT, +5 rps.
- `RRMZ50/-5`: axial 50 mT, -5 rps.
- `RRM50/5@0.05`: transverse 50 mT, +5 rps, independent 0.05 mT ARM bias.

Historical speed-only labels remain readable in old artifacts; live execution
rejects them because the AF field is unspecified. Requests require accepted
AF calibration and imported motor wiring/calibration. Nonzero signed speeds
must lie within the VB6 +/-40 rps bound and native signed 32-bit target/velocity
limits. The source bound does not establish the station's physically qualified
operating speed. Bias requests also require accepted, enabled ARM circuitry.

The lifecycle verifies home and coil positioning, orients the specimen, starts
asynchronous turning, verifies movement direction, and runs the calibrated AF
ramp with cooperative cancellation and rotation checks. Readback evidence keeps
native positions, deltas and elapsed time. This verifies continuing movement;
it does not claim independently calibrated rotational velocity or magnetic field.

Cleanup independently clears optional bias, resets AF and verifies stationary
turning. Unverified field/stop withholds reference and lift motion and inhibits
further queue treatments/motion and diagnostic leases. Recovery only clears
outputs and stops turning; it does not restart a spin or apply AF. Reference
restoration normalizes completed turns before rotation, avoiding a return across
the entire spin travel. A verified stop is required before lift return.

Immutable records use `rapidpy.rrm.treatment.v1` or `rapidpy.rrm.recovery.v1`,
retain both parameters and all failed phases, and publish through the current
run's AF artifact directory with required SHA-256 indexing. The manual Zero
Field recovery button also publishes RRM recovery records, including when ARM
and pulse IRM are disabled.

## Remaining system work

The full-system goal remains active. Remaining work includes zero-field queue
IRM semantics, routine/editor exposure of typed RRM parameters, persistence of
unsafe hardware state across process restarts, calibration creation/review,
probe motion, alternative legacy instrument routes, complete migration coverage,
fresh release packaging and physical/scientific acceptance. The existing portable
pilot predates the new treatment and motor source changes.

Verification: the complete main-app suite passed 673 tests in 84.794 seconds.
The additional native reference-normalization regression (thousands of turns,
both rotation signs) and the current motor/RRM subset passed 23 tests. The
source UI smoke passed six panels and all six helpers. No physical instrument
was connected or actuated by these checks. The full-system goal remains active.
