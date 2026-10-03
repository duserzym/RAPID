# Zero-field IRM and explicit treatment units

Source checkpoint, 3 October 2026. Hardware and scientific acceptance remain open.

## Zero-field queue IRM

The VB6 RockmagStep axial routine discharges twice with the specimen at load,
moves it to the coil, and skips FireIRMAtField when the field is zero. RapidPy's
queue now follows that sequence for IRM0 (or typed IRMZ0G / IRMX0MT / IRMY0MT).
The routine never commands a nonzero DAC charge voltage. It verifies capacitor
readback before each residual pulse, inhibits fire afterward, moves only after
the load-position discharges, and verifies final discharge/relay clearing before
reference and lift return. Cancellation or unsafe discharge retains failed-phase
evidence and follows the existing motor-inhibition/recovery rules.

Zero-field planning requires an enabled coil, explicit station wiring and valid
capacitor feedback/voltage limits. It does not use or require a magnetic field
calibration table, because no charged pulse is applied to the specimen. The
manual Zero Field recovery command remains discharge-only and does not fire.

Queue zero treatments use rapidpy.irm.zero_treatment.v1 and their circuit evidence
uses rapidpy.irm.zero_field.v1. Evidence includes completed residual pulses and
an explicit charged_pulse_fired value. These files are immutable, indexed with
required SHA-256 digests, and separated semantically from positive/backfield
pulse treatments and emergency recovery. Residual voltage below the configured
threshold does not prove an independently measured zero magnetic field.

## Field labels and routine generation

The typed parser converts explicit units once at the actuator boundary:

- IRM100G applies a 10 mT request; IRM100MT applies 100 mT.
- IRMZ / IRMX / IRMY select axial / transverse 0 degrees / transverse 90 degrees
  without changing global axis settings.
- IRM-100G is a numeric -10 mT backfield request and requires the polarity gate.
- ARM100MT_0.5G keeps a 100 mT AF peak and 0.05 mT DC bias independent.
- AF100G plans 10 mT rather than falling back to a default peak.

Unsuffixed labels retain the existing RapidPy mT execution contract. Archived
models remain readable. Newly generated IRM/ARM and manual treatment labels carry
explicit units; legacy VB6 routines require review of their source units before
execution. No conversion is inferred solely from an old unlabeled step string.

The sequence panel now uses its previously ignored RRM AF-field and ARM AF-peak
controls. RRM generation carries field, signed speed, axial/transverse choice,
and optional independent bias. ARM step controls represent bias in gauss. IRM
logarithmic fields carry G suffixes, and the shared minimum step is converted
from G to mT for the AF series. Backfield generation produces numeric signed
steps rather than the non-executable IRM-BF placeholder. The Works template
also retains input gauss units and explicit ARM peak/bias parameters.

The Hawaiian preset follows the documented button-caption series 25, 50, 100,
200, 400, 800 G (2.5 to 80 mT). VB6's loop starts at i=1, which disagrees with its
caption; RapidPy retains the documented caption series and makes units explicit.
Field formatting preserves small decimal values instead of rounding them to zero.

The controls scroll vertically and stack above the preview at compact widths.
Layout checks cover short windows and value-cell visibility. Visual inspection
used the Windows Segoe UI font loaded explicitly for offscreen rendering.

## Remaining development

The full-system goal remains active. Unsafe-state persistence across process
restarts, calibration creation/review, probe motion, alternative instrument
routes, complete legacy import coverage, release rebuilding and physical/
scientific acceptance remain required. The portable pilot predates this source.

Verification: the isolated main-app suite passes 696 tests in 90.593 seconds.
Source UI smoke passes all six panels and six helper applications. Zero-field
and pulse interlock checks include cancellation, unsafe discharge, temperature
rise between residual pulses, and a hot sensor after specimen positioning but
before charge. Typed-unit tests cover conversion, explicit axis selection,
independent ARM units, native record values and configuration immutability.
The sequence layout was inspected at wide and compact sizes; compact value cells
fit the viewport without horizontal scrolling. Physical acceptance remains open.
