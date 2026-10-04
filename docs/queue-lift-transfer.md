# Claimed specimen lift transfer

`QueueLiftTransfer` uses the actual routed native client, accepted XY geometry and
the original queue-owned vacuum transport. Its motion profile additionally binds
the legacy 0.3-second grip settling time and accepted dropoff delay. The legacy
`Vacuum.DropoffVacuumDelay` now imports into `vacuum.dropoff_delay_s`; missing values
use VB6's one-second default, and invalid/negative/nonfinite values remain unaccepted.

Every lift phase persists its intent and typed specimen context before motor I/O.
The context names the original queue token, specimen/file identity and original slot;
it retains raw pickup position and referenced sample height. Finished evidence uses
`rapidpy.queue_lift_transfer.v1` in the existing immutable queue history. The reader
validates the full linked history even after unrelated acquisition/field events. Any
pending stage returns an unverified transfer phase rather than a completed pose.
Missing publication leaves the original intent pending. Context never authorizes
replaying pickup or homing after restart.

Pickup requires the original slot, lift-top clearance, acknowledged pump ON/valve
OFF and independently verified fields off. Native gentle pickup retains the legacy
torque/motion/readback behavior. The phase checks stationary axes, original XY slot,
an unset top switch and a positive specimen height within accepted lift travel.

Loaded home requires that recorded pickup position and acknowledged grip. It waits
0.3 seconds cooperatively, homes the lift with the native top-switch sequence, and
checks stopped motors and top clearance. Sample height is pickup position minus
SampleBottom minus the live home offset, matching the VB6 calculation. It remains
per specimen; the accepted baseline motor configuration is not changed.

Returning from an acquisition pose raises to zero at the fast lift tier only above
a verified empty XY location with grip acknowledged and fields verified off. After
the table returns to the original slot, supported dropoff uses the measured height
in the XY target `SampleBottom + 0.9 * height` at slow speed. Fresh position, switch,
stationary-axis and original-slot checks qualify that commanded support pose.

After the coordinator obtains acknowledged valve OFF while retaining pump power,
lift clearance requires the same original support pose. It waits the accepted
dropoff delay cooperatively, moves to zero at the medium tier, then verifies top
switch, position and stopped motors. It leaves the pump owned and the outer queue
pending; it is not terminal queue cleanup.

Each failure independently stops all four motors, preserves failed acknowledgements
and leaves the original stage pending. No lift error path changes vacuum outputs.
Failures during loaded phases preserve commanded grip; failures after an acknowledged
release preserve pump ownership without pretending grip is still ON. Cancellation
during settling does not start the next motion. Field-off proof, live XY reference,
vacuum/pressure acceptance and physical specimen support remain the coordinator's
independent requirements; commanded poses do not prove laboratory acceptance.

Tests use actual native routed motor/vacuum clients with injected serial contact,
top-offset, switch and register responses. They compose pickup, grip, home, empty-hole
positioning, acquisition pose, return, supported dropoff, valve release and clearance.
They verify dynamic height, identity persistence, before-I/O journaling, wrong-slot
and unsafe-phase rejection, delays, cancellation, publication failure, corrupt history
and no replay. The acquisition pose is injected; these tests do not qualify real
measurements, grip or sample transfer. No physical instruments are actuated.

Still wire these phases into the automatic queue coordinator and measurement worker,
propagate measured height into bracketed SQUID/AF/ARM/IRM/RRM/susceptibility plans,
establish/verify the live XY reference, restore operator tray inversion and per-file
orientation, and implement original-panel stop/field/support recovery and terminal
vacuum release. Automatic live startup remains gated until that integration and
physical/scientific station acceptance are complete.
