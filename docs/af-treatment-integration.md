# Calibrated AF treatment integration

Source development checkpoint: 3 October 2026. Physical acceptance is pending.
The previously built portable pilot does not contain this checkpoint; rebuild
it after the treatment development milestone is complete.

`QueueHardwareBackend` now executes live AF through `AfTreatmentService`.
It plans the entire treatment before moving a motor. Normal AF runs axial,
transverse at 0 degrees, and transverse at 90 degrees, with the legacy
three-second pause before the first transverse pass. AFZ runs axial only.
AFMAX runs both transverse orientations followed by axial, using each coil's
independent maximum. The lift centres the specimen at the imported AF position;
VB6's `Int` rounding is preserved for odd specimen heights. Verified motion
outcomes are required. Field reset and return home are attempted independently
on completion, cancellation, and failure.

Calibration imports use the declared table count and units, preserving separate
coil frequencies, monitor/ramp limits, relay outputs and ramp timing. Fields
are interpolated with a zero anchor or extrapolated from the final interval
within the configured maximum. Invalid or missing calibration blocks execution.
Unlike the legacy silent clamp, an out-of-range request is rejected. The
transverse voltage conversion includes the source's factor of one half.
Production calibrated ramps use ADwin mode 2; mode 3 is a clip test. The driver
checks cancellation during process polling and stops the process on cancellation,
timeout or status-read failure. Coil limits are passed into the native process.

Each attempt produces an immutable `rapidpy.af.treatment.v1` record with the
planned ramps, specimen/run identity, completed passes, timestamped phases,
primary failure and cleanup failures. The measurement worker publishes only
the current run's records under `af_treatments/`, references them in the workflow
summary, and indexes them with SHA-256 digests. Publication failures are reported
and mark the run aborted.

Implementation evidence comes from `VB6/RockmagStep.cls`, `VB6/frmAF_2G.frm`,
`VB6/frmADWIN_AF.frm`, and `VB6/settings/Paleomag_v3.INI`. The 613-test main-app
suite passes, including injected tests for multi-pass ordering, accepted
calibration, source-profile import, cancellation, failed ramps, failed cleanup,
changed calibration, cross-board relay rejection, immutable records and run
artifact indexing. These tests do not actuate instruments.

Remaining AF work includes calibration creation/review controls, explicit
waveform channel binding verification and bench qualification of ramp quality,
relay transitions, motor clearances, cancellation and scientific outputs.
Independent ARM bias, pulse IRM, backfield polarity, synchronized RRM and the
MCC/DAC routes remain separate development requirements. This checkpoint does
not establish that the full system or its portable release is complete.
