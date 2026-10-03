# Pulse IRM development checkpoint

3 October 2026. The full-system goal remains active.

The legacy implementation charges a capacitor through a DAC, reads capacitor
voltage through an ADC, controls trim and fire switches, and selects the coil and
polarity through separate ADwin relays. An AF waveform cannot implement that
treatment. `VB6/frmIRMARM.frm` supplies the conversion, charge/trim/fire and relay
behavior, including separate Old, ASC and Matsusada implementations.

Completed source work:

- A separate persisted `PulseIrmConfig` imports the IRMPulse section, module
  enable flags, centering position, independent capacitor/field calibration
  tables, field limits, capacitor limits, control/feedback conversion factors,
  trim polarity and ASC boost values.
- `plan_pulse_irm` plans Old and ASC pulse voltages with source interpolation
  and bounded extrapolation. It converts imported gauss to mT, enforces accepted
  calibration, and requires explicit backfield enablement. ASC boost follows
  `CalculateAscBoostMultiplier`. Invalid field or boosted DAC targets are
  rejected before hardware output instead of silently clipped.
- The MCC driver has checked ADC input and engineering-unit conversion calls,
  using the `cbAIn`/`cbToEngUnits` interface declared in `VB6/CBW.BAS`. Read and
  conversion errors cannot produce usable voltage evidence.
- Tests verify unit conversion, imported module disablement, coil-specific
  capacitor values, ASC boost, backfield gating, invalid calibration, persistence
  and checked ADC behavior. No instrument is opened or actuated by these tests.

The station profile selects ASC and disables both IRM coils and backfield. Import
preserves those flags. A field of 100 mT maps to approximately 33.195 capacitor
volts using that profile's accepted axial table, rather than an AF peak voltage.
The imported capacitor maximum of 420 V is distinct from the maximum field of
1200 mT. This corrects a distinction the earlier generic IRM configuration loses.

## Circuit and workflow integration

`PulseCircuit` now executes checked capacitor charge/readback, trim, active-low
fire and discharge. It verifies the ADwin coil/polarity word before charging,
requires three stable readings plus a fresh pre-fire reading, checks capacitor
limits, and applies finite charge/discharge deadlines. Cleanup attempts every
output independently. It releases coil relays only after discharge is verified
and all cleanup outputs were acknowledged. Enabled temperature channels are
checked against the imported slope/offset, zeroed-sensor condition and hot limit
before charging, during charging and before firing.

Matsusada planning now includes all `ScaleUp` intervals. Old/ASC/Matsusada runs
must satisfy the requested capacitor tolerance; the workflow does not accept
the legacy plateau fallback as proof of the requested field. That deliberate
stricter behavior and timeout policy require bench/scientific qualification.

`PulseTreatmentService` verifies the load position and requested orientation,
performs the source's two residual zero-charge pulses at load, centres the
specimen at the imported IRM position, and returns it after verified circuit
cleanup. It withholds motor return when capacitor/relay safe state is unknown.
The queue attempts a separate discharge-only recovery before subsequent motion
or general relay reset. That recovery never charges or asserts fire. An
unresolved pulse fault blocks subsequent treatments, motor commands and other
measurement/changer/AF ownership until a verified recovery succeeds.

Live queue IRM now uses this workflow. The old adapter cannot substitute an AF
waveform for a capacitor pulse, and its private generic ramp implementation was
removed. Negative IRM fields require enabled backfield and change the polarity
relay word while retaining non-negative DAC voltage. Z uses axial; X/Y use the
transverse coil at 0/90 degrees. Legacy axis names are normalized to the UI values,
and capacitor voltage is no longer imported as a field in mT.

Immutable `rapidpy.irm.pulse.v1` and `rapidpy.irm.treatment.v1` records retain
calibration/binding snapshots, capacitor readings, relay results, phases,
primary/cleanup failures and mechanical return outcomes. The measurement worker
publishes current-run records under `pulse_treatments/`, indexes SHA-256 digests,
and references identities in the workflow summary. Manual IRM uses the same
worker/ownership/evidence flow as manual ARM; signed fields, specimen identity
and axis are explicit. Zero Field retains discharge-recovery evidence, including
failed recovery, and does not conceal publication errors.

Verification: all 651 main-app tests pass in 88.038 seconds. Coverage includes
actual injected queue/circuit execution, orientation and centering, stable/fresh
charge proof, relay mismatch, temperature faults, deadline/cancellation failures,
unverified-discharge motor inhibition, discharge-only recovery, manual records,
worker artifact indexing, and restored queue context. The source six-panel and
six-helper UI smoke check also passes. No instrument was opened or actuated.

Remaining development includes calibration creation/review controls, zero-field
queue semantics, completion of motor/station configuration migration and other
treatment routes. Physical relay behavior, capacitor thresholds, native vendor
drivers, discharge behavior, coil/motor clearances, pulse waveforms and scientific
comparisons remain unqualified. The full goal is active; the portable pilot
predates these source changes and must be rebuilt after development stabilizes.
