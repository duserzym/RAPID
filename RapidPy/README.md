# RapidPy Subsystem Apps

RapidPy now contains multiple subsystem control apps to support a staged VB6 -> Python transition.

## Apps

- `vrm_logger`: VRM logging and SQUID live plotting
- `af_tuner`: AF coil tuning panel based on `frmAFTuner`
- `changer_xy_control`: hole/sample list and queue prep based on `frmChanger`
- `updown_control`: vertical axis controls based on `frmDCMotors` up/down panel sections
- `dc_motor_control`: general motor panel based on `frmDCMotors`
- `system_shell`: operator launcher for all converted subsystem panels

## Shared Layer

- `rapidpy_common/ui.py`: shared UMN maroon/gold liquid-glass style
- `rapidpy_common/hardware.py`: VB6-aligned Quicksilver motor protocol, movement, and conversion utilities
- `rapidpy_common/adwin_af.py`: ADWIN AF ramp backend (`boot/load/set params/start/readback`)

Implemented parity highlights:

- Up/down `HomeToTop` switch-stop behavior and sample pickup/dropoff torque choreography
- XY table `HomeToCenter` and `MoveToCorner` limit-switch guided routines
- AF relay selection through ADWIN digital output bit setting before AF ramps

## Transition Plan

1. Keep each subsystem app familiar and standalone for operators.
2. Validate protocol-accurate hardware behavior on Windows against machine limits and safety interlocks.
3. Merge app backends into a single orchestrated control app once workflows are validated.
