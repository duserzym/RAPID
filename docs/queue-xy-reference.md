# Recorded native XY reference

`QueueXYReference` establishes a live reference under the original claimed
`QueueWorkflowSession`. It uses the active VB6 HomeToCenter sequence: negative
edge X bit 4/Y bit 5 with relative travel and stop masks -1/-2, then positive edge
X bit 5/Y bit 6 with masks -2/-3. Both axes use accepted controller travel/speed and
the native XY acceleration of 483184. It issues each axis command once per pass.

The root must bind the original four-axis motion profile and vacuum transport.
The caller must confirm the empty rod and verify field outputs off. Acknowledged
valve OFF, no active specimen/blank-holder context, connected original motors,
live lift top-switch and calibrated lift-clearance readbacks are required. This
service does not home the lift, discharge fields, release a specimen or change
vacuum outputs. The coordinator must establish those preconditions first.

The motion intent is written before Stop, switch reads or sweep commands. Every
axis receives independent Stop and stable register 1/7 checks before homing, at
each edge boundary and after the operation. Complete binary switch observations
must verify the edge and remain consistent with the opposing switch. All-zero
opposing switch states fail closed. Relative count overflow is rejected before
motion. Loss of lift clearance, failed Stop acknowledgement, cancellation,
deadline expiration, changed calibration/routing or changed connection identity
prevents remaining commands and leaves the original stage pending.

Only after both stopped edge boundaries are verified are X and Y targets zeroed.
Final positive edges, lift clearance and stationary coordinates must verify the
accepted imported XYHome pair within the same 0.02-slot count tolerance used for
table geometry. Accepted calibration is not rewritten. The final reference
context is published in an immutable linked queue event before verification.
Publication failure prevents handing out a usable reference.

Each successful `RoutedMotorSerialClient.connect` creates a new connection ID.
Disconnect attempts invalidate the ID, including failed closes. A failed native
close retains its original handle/client for retry and prevents reconnect from
replacing it. A partially closed station does not report all motors connected.
The typed `QueueXYReferenceState` includes the original queue token, unique
reference ID, connection ID and actual stationary home counts. Reconnect,
another queue, tampered evidence or an unfinished later stage invalidates it.
Restart recovery never replays the reference sweep.

`QueueXYReference.require` retrieves a usable original reference. `move_to_slot`
and blank-holder `prepare` can accept that typed state as `reference_verified`,
validate it against linked history/current transport and persist its proof in
the operation plan. Existing explicit operator-attestation callers still accept
strict Boolean True; automatic coordinator wiring must use the typed reference.

The complete automatic coordinator, field/support verification, operator load/
tray-inversion workflow, original-panel recovery and terminal release remain open.
Tests inject serial switch/position replies into the actual native routed controller;
they do not prove physical sensor polarity, switch separation, table origin, rod
clearance or station calibration. These need physical qualification before live
automatic execution. The portable build must be rebuilt after integration.
