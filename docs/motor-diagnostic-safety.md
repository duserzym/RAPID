# Lift diagnostic ownership and restart recovery

Up/Down Control uses the shared `HardwareSafetyStore` journal and OS lifetime
lease for native lift movement and measurement Z scans. The `motion_diagnostic`
family blocks main-app treatments and other diagnostics until verified completion
or explicit recovery. A pending diagnostic is recovered in the original helper;
the main app does not replay a lift motion or substitute its own station profile.

The physical profile contains the actual motor COM port, axis identity/address,
native motor settings and lift position calibration. The settings file location
is excluded so moving an unchanged file does not change station identity. Connecting
to a pending lift diagnostic requires this exact profile. Motor connection setup
holds the operation lease, so even its broadcast ACK-delay command cannot race an
active treatment. Motor and SQUID connection changes, settings changes, baseline
replacement and additional motion are blocked while a scan is queued or active.

Every movement records its request/result. A complete Z scan holds one lease across
all movements, settling intervals and SQUID acquisition. Evidence includes specimen
height, signed position calibration, safe bounds, baseline, SQUID calibration,
calibrated scan points and native raw XYZ readings when supplied by the native
reader. Cancelling retains collected points and failure evidence. Scan settling is
cooperative, and cancellation cannot produce a successful final scan signal.

Cleanup issues a native Stop and reads Quicksilver registers 1 (actual position)
and 7 (both native velocity words) twice, with a short interval. Both complete
velocity registers must be zero, and actual positions must be identical. Stable
position alone does not prove a motor has stopped. A failed Stop acknowledgement
attempts Halt independently and keeps the recovery pending even if subsequent
telemetry appears stationary. Invalid telemetry, drifting position or a nonzero
velocity word retains the latch. Cleanup never homes, zeroes or relabels the axis.

Records and their required SHA-256 artifact indexes are published immutably under
`hardware_diagnostics/motion-<id>` before clearing the latch. Evidence publication
failure retains it even when physical stop verification succeeded. An explicit
stop verification without an existing latch creates one first, so failed recovery
also survives restart. Restore the original motor port and settings, connect, and
use **Recover: Verify Stopped In Place**. Inspect specimen clearance before choosing
a subsequent movement; recovery does not automatically return it home.

Closing a live scan requests cancellation and retains the window and connections
until its thread actually finishes checked cleanup. It does not terminate a worker,
discard it after a timed wait or open a failure dialog while closing. Compact
screens provide Connections & Status, Settings & Console, Axis Profile and Motion
& Scan tabs; the recovery button remains visible at 736x720.

Development validation uses injected motor/SQUID interfaces, including restart,
cross-owner exclusion, cancellation, readback drift, failed stop and failed evidence
publication. Physical instruments were not actuated. Vacuum/gripper lifetime
ownership, other motor/direct diagnostics, durable main-app non-treatment motion
and acquisition recovery, calibration/probe tooling, the portable rebuild and
physical/scientific station acceptance remain required full-system work.
