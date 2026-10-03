# Durable hardware treatment recovery

Live queue and manual AF, ARM, pulse IRM and RRM now persist an unfinished
operation before treatment motion or field output. The latch is stored at
`~/.rapid/hardware_safety.json`, independently of operator configuration and
measurement destinations. `RAPID_SAFETY_STATE` is an explicit isolated runtime
override used by development tests; it is not an operator setting.

The snapshot retains specimen/run identity, immutable treatment plan, station
wiring/calibration, operation token and latest physical safety evidence. Writes
flush and fsync a temporary file before atomic replacement. A checksum and
strict schema validation make corrupt or unreadable state a blocking error.
Transaction locks serialize journal updates. A separate OS lifetime lease
prevents another app instance from recovering an operation still owned by a
running treatment; the operating system releases this lease after a crash.

Verified physical cleanup clears the pending state. Treatment failure with
verified cleanup may also clear it; failure without verified cleanup, process
interruption and unsuccessful publication keep it pending. Stale tokens and
changed station profiles cannot clear a newer or differently wired operation.
NoComm leaves the physical latch untouched.

On restart, queue preflight, actuator connection and main-app diagnostic leases
block ordinary motion/field work. The manual IRM/ARM recovery control remains
available when its physical adapter is connected. Recovery uses the original
specimen/run identity and publishes its records through the existing worker or
manual artifact index.

Pulse recovery first inhibits firing, zeroes charge, enables bleed, qualifies
capacitor discharge and checks relay clear. Only then may it connect the motor
transport, verify stationary turning, restore reference and home the lift.
The latch remains pending if any required motion return fails. No recovery
charges the capacitor or fires a residual/specimen pulse.

AF/ARM/RRM recovery stops and checks ADwin process status using the existing
VB6 `Get_Par(-100 + ProcessNo)` contract, zeroes the ramp DAC and checks relay
readback on the existing board. It does not boot the ADwin board. Independent
ARM bias cleanup, stationary turning and reference/lift return must succeed
before the operation is verified. Recovery never reapplies AF, ARM bias or spin.
Controller status and output acknowledgements are hardware evidence; physical
field-zero and scientific acceptance still require station validation.

The journal is now shared with AF Tuner; see [AF diagnostic safety](af-diagnostic-safety.md).
Remaining full-system work includes extending ownership/persistence to the other
hardware helper processes and embedded direct diagnostics, non-treatment motion/acquisition interruption,
calibration acceptance tooling, probe workflows, rebuilding the portable pilot
from current source, and physical/scientific acceptance. A missing first-run
latch is not evidence that an externally operated instrument is physically safe.
Corrupt state is never automatically deleted or silently replaced by recovery.

Validation: 724 isolated main-app tests pass (92.420 seconds), including 28
safety-store, restart, treatment-ordering, ownership and recovery regressions.
Six-panel/six-helper source UI smoke passes. These checks use injected physical
interfaces and do not actuate the installed station.
