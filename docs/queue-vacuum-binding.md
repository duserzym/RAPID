# Native vacuum borrowing under a durable queue

The main `VacuumBackendAdapter` exposes `queue_station_binding`, `queue_connect`,
`queue_set_pump` and `queue_disconnect` for the queue coordinator. These methods
borrow an exclusive `QueueWorkflowSession` worker claim rather than acquiring a
second diagnostic OS lease. They require the same fixed journal, original queue
token, `rapid_main_queue` helper identity, port, baud and acceptance binding before
opening a transport or issuing outputs.

The stage profile includes the configured pressure threshold. A positive threshold
requires available pressure telemetry; this serial controller does not provide it.
Threshold zero is the explicit command-only mode. Changing that choice while the
queue owns hardware cannot change its latched binding. Connecting does not reset or
enable outputs and does not interpret the new connection's cached OFF flags as proof.
Saved Auto Pump does not actuate the queue connection.

Enabling persists a vacuum stage before the native `E`, `10MFF`, `O`, `10VFF`
commands. The adapter requires both exact command replies and the acknowledged
controller state before publishing a held stage. Evidence retains the raw commands
and terminated replies accepted by the legacy serial driver. It establishes
commanded state; it does not establish measured pressure or specimen grip.

The transfer API `queue_set_outputs` now distinguishes pump power from gripper
valve connection. Pump ON / valve OFF persists a `pump_ready` stage before native
`C`, `10V00`, `E`, `10MFF` commands. Both fresh replies and both requested controller
states are required. Its phase record explicitly reports no acknowledged grip and
no safe terminal state. The durable queue treats this as retained output ownership,
allows the next claimed stage, and forbids transport detach or queue completion.
The adapter reports pump and valve command states separately.

Connecting grip through the transfer API requires strict verified motor stop,
specimen pickup position and field outputs off. Disconnecting the valve while
leaving the pump powered requires verified stop, specimen support and fields off.
Pump OFF / valve ON is rejected. A failed valve-OFF acknowledgement does not proceed
to pump enable. Restart/lost-link recovery refuses pump-ready as well as grip enable;
full OFF recovery remains the only allowed output transition. Failed evidence
publication retains the pending stage and unknown state for that OFF recovery.

Release requires strict motor-stop, specimen-support and field-output-off checks
from the coordinator before any release I/O. Native release attempts valve OFF
(`C`, `10V00`) and pump OFF (`D`, `10M00`) independently, even after one fails.
Both fresh replies and both controller OFF states are required. Publication failure,
partial enable, missing replies and failed release retain the original pending stage
and unverified owner. An explicit OFF retry recovers the same token without replaying
ON. Diagnostic Pump/Release/Disconnect cannot replace queue ownership.

Restart recovery obtains the original queue profile and exclusive lease. Recovered
sessions and a lost/reopened live connection refuse enable and permit OFF recovery
only. A reused connection cache is unverified. Disconnect requires persisted original
queue release evidence and keeps the queue journal pending for independent final
motor/field verification. Failed transport closure retains the original handle and
owner for an explicit retry, including closure failures after the port state changed.

`test_queue_vacuum.py` uses the actual native controller with an injected serial
interface to check commands, replies, before-output journaling, lease exclusion,
release preconditions, partial failures, publication failure, recovery and handle
retention. No physical instruments are actuated.

This is the native transport path for the coordinator, not automatic queue startup.
The complete queue transfer and operator-intervention phases remain to be wired;
see `queue-transfer-audit.md`. Keep live startup gated until those phases are
implemented. Pressure/gripper, motion clearance, scientific acceptance and portable
rebuilding remain required full-system work.
