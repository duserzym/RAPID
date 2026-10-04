# Claimed specimen acquisition

A `QueueHardwareBackend` bound to `QueueSpecimenGeometry` stages public SQUID
and susceptibility reads under the original `QueueWorkflowSession` worker claim.
The acquisition binding includes the original motion, SQUID, calibration and
susceptibility settings. The stage records specimen identity, measured height and
count targets before connection, reset, measurement or motion I/O.

The owned stage token permits its immutable specimen geometry to remain available
during acquisition and cleanup. A pending foreign stage, changed station/settings,
foreign journal, returned specimen or recovery session cannot borrow that geometry.
Count targets must fit the signed 32-bit motor range. Native preflight uses the
owning worker directly; individual adapters retain their bounded I/O deadlines.

SQUID settlement requires the actual returned block to match the completed
`BracketedAcquisition`, original sample/run and count targets, physical audit flags,
successful command evidence and six validated coherent observations. Transport
recovery records accompany the block. Susceptibility requires a new typed record,
the original measured height and target, successful acquisition phases, finite
matching result and verified safe return. Failed bridge records are retained as
failure evidence. A flux-count recovery hook also begins a separate acquisition
stage before motion/reset and cannot replay an unfinished acquisition.

All four original axes receive independent Stop and register 1/7 checks after
acquisition, including failed instrument I/O. Failed Stop acknowledgement remains
unsafe even if fallback Halt and later telemetry show zero velocity. Identity and
configuration are checked again after cleanup; cancellation before settlement keeps
the stage pending. The immutable linked queue event is published before the stage
can become verified. Publication failure prevents a completed result handoff.

These operations retain the specimen grip and outer queue lifetime. Stage verification
does not establish field outputs off, specimen support, transfer clearance, pressure
telemetry or final release. Scientific block reduction/acceptance remains a separate
requirement. A failed acquisition requires original-panel recovery before another
stage; no recovery replay or grip release is added here.

The full queue coordinator and MeasurementWorker must still borrow this original
claim and invoke geometry binding. Blank-holder geometry, initial station reference,
operator tray inversion, per-file orientation and terminal support/field/stop/vacuum
recovery remain open. No physical instruments were actuated by these tests, and
portable and scientific station qualification remain required.
