# Held vacuum and lift diagnostics

Up/Down Control holds one `station_diagnostic` journal and OS lifetime lease while
vacuum is commanded on. A vacuum-only session remains supported. A participating
lift can join the session later; its original port, address and calibration are
persisted before its first connection setup or motion. Existing resource bindings
cannot be replaced. This allows lift moves and complete Z scans under a held
vacuum without releasing the specimen between clicks or worker operations.

Each lift operation verifies stopped-in-place readback before reporting completion,
then publishes a station checkpoint while retaining vacuum ownership. A checkpoint
with intentional held outputs is indexed as `held` and cannot clear the journal.
Other helpers and main-app treatments remain blocked. Main-app native vacuum
connection/output commands also check the shared journal and lifetime lease.

Releasing vacuum first verifies that every participating lift has stopped. If that
verification fails, no valve/pump-off command is issued and the session stays owned.
With the lift stopped, valve-off and pump-off acknowledgements are attempted
independently. Failure of the first action does not suppress the second. Missing
acknowledgements retain ownership and durable recovery state. Immutable station
records include original resource bindings, movement/acquisition observations,
raw vacuum command/reply evidence and required SHA-256 artifact indexes.

The legacy vacuum protocol (`VB6/frmVacuum.frm`) supplies terminated command
responses; it does not supply pressure telemetry. `safe_state_confirmed` for this
diagnostic means that participating lift stop readbacks passed and both vacuum-off
commands received terminated replies. It does not prove pressure, specimen grip
or the physical valve state independently. Those require station acceptance. UI
status uses *commanded ON/OFF*, *unverified* or *recovery pending*. A saved checkbox
never proves an output state and cannot re-enable vacuum on startup.

Connecting the vacuum helper no longer resets outputs. While a held station needs
recovery, only its original bindings may reconnect. Reconnecting vacuum requires
the participating lift to be stopped and marks vacuum command state unverified;
new motion is inhibited until recovery. A faulted lift may reconnect at its original
binding without releasing vacuum. Recovery never enables vacuum, homes an axis,
relabels position or replays a scan.

After restart, reconnect the original vacuum controller and any participating lift,
then use **Recover Held Vacuum / Lift Outputs** in Connections & Status. A live
scan must finish cooperative cancellation and checked cleanup first. Closing a
window with held outputs runs release; if lift stop or vacuum release remains
unverified, the window and connections remain open for recovery. A publication
failure after physically acknowledged release preserves the pending journal but
does not retain a live hold. On disk, recovery requires the original ports and
profiles before any release command.

Development checks use injected native interfaces. They cover held movement from
a worker thread, vacuum-only use and later lift participation, stale/rebound
resource rejection, stop failure retaining vacuum, independent release attempts,
acknowledgement failure/retry, original-binding restart recovery, evidence failure,
main-app exclusion and close behavior. Physical instruments were not actuated.
Other embedded output lifetimes, other motor helpers, durable main-app acquisition
and non-treatment motion recovery, calibration/probe tooling, portable rebuilding
and physical/scientific station acceptance remain full-system work.
