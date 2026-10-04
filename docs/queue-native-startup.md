# Native queue startup

`QueueHardwareBackend.prepare_queue_lifetime` constructs the original native XY
table, field cutoff, vacuum and scientific stage profiles without opening ports or
changing outputs. It requires a named operator's explicit empty-rod confirmation,
a queue plan and run identity. It validates accepted motor calibration, all four
axis bindings, 57600 baud, the original vacuum configuration and original ARM/pulse
configurations. Retained motor/vacuum handles, diagnostic holds, unavailable
components and already-borrowed owners require their original settlement first.

Preparation persists the complete parent queue and timestamped empty-rod attestation
under its OS lifetime lease, then lends its child store to the original backend.
The owner retains the returned startup/session object through worker exit and
eventual terminal settlement. It must not release that lifetime lease while a
worker claim remains active.

`start_queue_lifetime` runs under `QueueCommandWorker`'s original session claim:

1. Verify participating native ARM, AF and capacitor/relay field cutoff.
2. Begin a motion stage before opening original motor ports or issuing their native
   connection broadcast. Independently Stop and verify all four axes, then publish
   the original connection identity and linked empty-rod stop evidence.
3. Connect the original native vacuum. The typed empty-rod stop proof permits only
   valve OFF; it cannot authorize grip. Persist acknowledged valve OFF/pump ON.
4. Begin an empty-lift reference motion stage. Verify all four stops. If the top
   switch is unset, issue one bounded native upward command with the original top
   switch stop condition. Require live top detection and independent stopped
   telemetry before zeroing. Verify zero/clearance/top and unchanged valve OFF.
   There is no unguarded second homing command.
5. Perform the existing two-edge XY reference, preserving lift clearance and the
   original field/vacuum proof. Bind the measured transfer coordinator and terminal
   cleanup to the same native objects and durable owner.

Changed configuration, invalid telemetry, absent switch detection, cancellation,
connection/publication failure and unfinished startup retain parent ownership.
Startup cannot replay: any second call requires explicit original-stage recovery.
Recovery sessions cannot use this live startup path. No specimen, nominal height,
grip ON or measurement reading is fabricated during startup.

Injected native DLL/serial tests cover preparation, claim boundaries, before-I/O
journaling, cutoff/stop failures, calibration/retained handle rejection, successful
top referencing, cancellation before zero and startup-to-terminal settlement.
These tests do not qualify physical top/XY switches, pressure, field strength or
mechanical clearance. MainWindow now supplies whole-run device leases, operator
confirmation/loading stages and dispatch through this startup API; see
`native-main-queue.md` for the integration and remaining recovery/qualification work.
Physically present participating circuits are currently mandatory; absent-hardware
configuration must be distinguished from disabled treatments before optional
circuits can be supported.

Startup also captures original scientific adapters/clients and accepted serial
settings without opening ports, rejects retained scientific handles and serial
circuit collisions, and binds that ownership to the durable root. See
`queue-instrument-settlement.md` for acquisition and terminal participation.
