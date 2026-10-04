# MainWindow native queue lifetime

Hardware-mode QueueHardwareBackend queues now use the native startup, measured
transfer, blank-holder and terminal services. MainWindow reserves measurement,
changer, AF/field, vacuum, SQUID and susceptibility resources for the whole run.
The ownership manager reserves the group atomically; conflicts leave existing
ownership intact, and releasing a group twice cannot consume another reservation.

The named operator confirms an empty control rod before preparation. The durable
parent queue then captures commands, samples, table options and explicit per-file
initial orientations. Startup runs in the original queue command worker. After
verified startup, a bounded negative-edge XY sweep parks the empty, clear rod's
table at the VB6 loading corner, without zeroing the XY reference or rotating the
rod. The operator confirms loading and clearing their hands. Fresh field cutoff,
four-axis stop/clearance, both negative limit edges and valve OFF evidence precede
publication of the operator confirmation.

InitUp remains the VB6 preprocessing marker and issues no motor command. Holder
resolves the calibrated empty hole, records a distinct zero-height blank pose,
binds owned acquisition, measures/install its correction, returns the blank rod
to verified clearance and clears that geometry. It never picks up a specimen.
Holder uses VB6's up orientation and restores the previously selected file direction
after its verified clearance. A failed blank acquisition/return retains the bound
geometry and original queue owner.

Meas first uses an owned worker to load the original command slot, specimen and
file identity. It selects the acknowledged direction for that file and binds
measured height before starting MeasurementPanel. Pausing at this handoff leaves
the loaded pose owned; resume starts that same measurement without reloading or
skipping the command. Native preflight composes/validates acquisition without
testing/opening SQUID or issuing instrument commands. Connection/reset/read I/O
belongs to the pending acquisition stage.

Goto parks the native XY table at the loading corner. Flip parks the empty rod,
prompts the operator to reverse the selected file's tray specimens and publishes
fresh confirmation. It never substitutes an automatic 180-degree rod rotation.
Only that file's direction changes. Duplicate confirmation, unknown files,
changed switches, active specimens and unfinished stages cannot authorize a
new direction. Per-file orientations survive in SHA-linked operator evidence.

Diagnostic vacuum ON and diagnostic SQUID test calls are not queue preconditions.
The queue uses its original acknowledged pump/valve state and original acquisition
owner. Reentrant queue leases require that same native backend/session, full group,
matching scientific profiles, a completed startup and a verified/held child stage,
with no live command worker. Other panels/helpers remain blocked.

Terminal cleanup runs after the actual measurement/command worker exits. Only the
original verified queue finish and closed native motor/vacuum bindings permit OS
and logical group release. Cutoff, acquisition, return, transport/publication failure
or an invalid journal retains ownership. Shutdown stays open for original recovery
instead of repeatedly closing or discarding its unresolved owner.

Injected native DLL/serial and isolated Qt checks cover startup, operator confirmation,
blank acquisition composition, measured slot/file handoff, device conflicts,
pause/resume, original reentrancy, successful terminal release and failed shutdown.
The MainWindow orchestration fixture substitutes the measurement panel's scientific
run and holder command; separate acquisition/holder/worker tests cover those services.
It does not prove a complete physical scientific station run.

Remaining full-system work includes original-panel live/restart queue recovery,
SQUID/susceptibility original transport identity and terminal settlement, complete
per-file step/eligibility and scientific artifact acceptance, optional physically
absent field circuitry, chain-station transfers, remaining VB6 auxiliary tools,
portable rebuilding, performance and physical/scientific station qualification.
No physical instruments were actuated by these checks.
