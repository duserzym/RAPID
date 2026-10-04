# Held vacuum diagnostics and integrated main control

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

## Diagnostics > Vacuum

The integrated native backend now constructs without opening a port or enabling
outputs. **Connect** runs in a worker and opens the configured transport without
resetting the valve or pump. Saved **AutoPump** settings never enable live outputs
on startup or connect. Output state remains **unverified** until this session has
acknowledged explicit enable or release commands; the cached flags initialized by
serial connection do not prove physical outputs off. The monitor and shared vacuum
snapshot distinguish unverified state from acknowledged off state.

**Pump On** persists `station_diagnostic` ownership with helper identity
`rapid_main_vacuum`, the original COM port and baud before output. It retains an OS
lifetime lease and pending journal for the entire commanded hold. Other hardware
panels and automated operations cannot use that held station. **Release / Verify
Off** uses the shared native release path, independently acknowledging valve-off
and pump-off commands before publishing immutable indexed evidence and clearing the
latch. Failed release retains the live owner and connection; failed publication
after acknowledged off preserves durable recovery without retaining a live hold.
The integrated panel does not borrow lift motion into its vacuum hold; the Up/Down
helper supports that explicitly participating resource workflow.

Restart recovery requires the original main panel and its original configured port
and baud. An Up/Down hold cannot be recovered through the main vacuum panel merely
because a port matches. Recovery never enables outputs or replays movement. The
main recovery action routes unfinished integrated vacuum operations back to
Diagnostics > Vacuum. Corrupt journals fail before output and produce visible faults.

Connection, pump commands, release and disconnect run in workers. Closing during a
command keeps the worker and connection, then requests checked release before
disconnect. Failed release keeps the window open for explicit retry and cancels
parent-window shutdown. Shutdown does not repeatedly issue release commands. The
parent requests each owned dialog's closure once and waits for actual completion.
Mode/operator changes cannot replace a live diagnostic owner. Pending recovery
blocks entry into No-Comm while permitting hardware mode for original-panel recovery.

Pump, Release / Verify Off and Close remain outside the scroll area on compact
displays. Thresholds, connection information, readings and evidence notes scroll;
wrapped faults retain sufficient height. Labels say **commanded** on/off, never
infer pressure or specimen grip from native replies, and explicitly describe the
no-pressure-telemetry evidence basis. Pressure-based automated acceptance remains
a separate station requirement.

Each new shared hold clears prior observations and lift participation. Immutable
checkpoints link to the previous record instead of growing live observation history
indefinitely. Native acknowledgement buffers are bounded and reset for each new
diagnostic lifetime; published records retain the relevant raw replies. Recovery
records identify the original pending token and previous artifact identity.

Other embedded output lifetimes, other motor helpers, durable main-app acquisition
and non-treatment motion recovery, calibration/probe tooling, portable rebuilding
and physical/scientific station acceptance remain full-system work.
