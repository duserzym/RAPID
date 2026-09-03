# Replacing the PCI-DAS6030 with a USB device — 2026-09-03

The lab computer no longer has a PCI slot. This document specifies two ways to
put the same signals on USB, in enough detail to build or buy either one.

Nothing here has been built or wired. Three measurements are still outstanding
(see [Before anything is wired](#before-anything-is-wired)); the safe idle state
of the design depends on them.

## 1. What the card actually does

Every analog and digital operation in the VB6 application funnels through one
dispatcher, `frmDAQ_Comm.DoDAQIO` in `VB6/frmDAC_Comm.frm:578`, which calls into
`VB6/Board.cls`. Enumerating its callers gives a complete, closed inventory —
there is no second path to the hardware.

Against the live site configuration in `VB6/settings/Paleomag_v3.INI`:

| INI channel | Type | Physical | Function | Live |
|---|---|---|---|---|
| `ARMVoltageOut` | AO | `AO-0-CH0` | ARM bias voltage | **yes** |
| `ARMSet` | DO | AUXPORT bit 0 | TTL relay, ARM box → ARM coils | **yes** |
| `VacuumToggleA` | DO | AUXPORT bit 6 | Valve, quartz rod → vacuum | **yes** |
| `DegausserToggle` | DO | AUXPORT bit 2 | Degausser cooler | **yes** |
| `MotorToggle` | DO | AUXPORT bit 5 | Vacuum pump request | not needed — building vacuum |
| `IRMVoltageOut` | AO | `AO-0-CH1` | IRM charge setpoint | no |
| `IRMFire`, `IRMTrim` | DO | bits 1, 3 | IRM fire / trim | no |
| `AnalogT1`, `AnalogT2` | AI | `AI-0-CH3/4` | Coil temperatures | no |
| `IRMPowerAmpVoltageIn` | AI | `AI-0-CH1` | IRM amp monitor | no |
| `ALTAFMONITOR` | AI scan | `AI-0-CH1`, 100 kS/s | AF LC current monitor | **no** — see below |
| `IRMMONITOR` | AI scan | `AI-0-CH4`, 100 kS/s | IRM discharge monitor | no |
| `VacuumToggleB` | DO | bit 7 | — | declared, never driven |

The IRM and temperature rows are off because `[Modules]` sets
`EnableAxialIRM`, `EnableTransIRM`, `EnableIRMBackfield`, `EnableIRMMonitor`,
`EnableIRMReturn`, `EnableT1` and `EnableT2` all to `False`.

`ALTAFMONITOR` needs a second look, because `EnableAltAFMonitor = True` makes it
appear live. Its only consumer is `frmAF_2G.frm`, which `modAF.bas:54` reaches
only when `AFSystem = "2G"`. The site runs `AFSystem = ADWIN`, so the AF ramp,
AF monitor and AF/IRM relay TTL are all on the ADwin-light-16 over its own link.
The flag is vestigial.

**Nothing timed or buffered remains on the card.** What is left is one DC level
and three on/off lines.

### The ARM analog channel in detail

From `[ARM]` in the live INI:

```
ARMMax       = 10.1     ' Gauss, full scale
ARMVoltGauss = 0.208    ' volts per Gauss
ARMVoltMax   = 2.1      ' volts, hard ceiling
ARMTimeMax   = 600      ' seconds the bias may stay energised
```

`SetBiasField` (`VB6/frmIRMARM.frm:3363`) is the whole driver:

1. zero the DAC, wait 0.5 s;
2. if the target field is positive, close `ARMSet`, wait 0.5 s, write
   `ARMVoltGauss × Gauss` to the DAC, wait 1 s to settle;
3. if the target is zero, zero the DAC, wait 0.5 s, then open `ARMSet`.

Voltage is clamped to `[0, ARMVoltMax]` in software before it reaches the DAC.

So the requirement is **0–2.1 V DC, written about once every half second and
then held constant for the duration of a demagnetisation step**. The card's
16-bit DAC on its `UNI10VOLTS` range resolves 152 µV, or 0.73 mG. That is the
bar, and it is a low one — the card was sized for the 100 kS/s work that is no
longer running.

### The overheat timer is software-only, and that matters

`frmIRMARM.frm:3647` zeroes the bias field once `ARMTimeMax` has elapsed. That
timer exists because the ARM coil overheats if left energised. It runs in the
VB6 process.

On the PCI card, a host crash leaves the DAC holding its last voltage and the
`ARMSet` relay closed, indefinitely, with nothing left running to open them.
A USB replacement should not reproduce that. Both designs below are judged
partly on whether they improve it.

## 2. Before anything is wired

Three readings, taken with the 68-pin cable **unplugged from the breakout**, so
the boxes are seen unloaded:

1. **`ARMSet` input on the ARM box — open-circuit voltage.** Pulled high, pulled
   low, or floating? This decides whether an undriven line leaves the ARM coil
   *connected* or *disconnected*, and therefore what "safe state" means for the
   whole design. Everything else can be adapted later; getting this wrong can
   energise a coil at power-up.
2. **Vacuum valve and degausser cooler inputs — same measurement.** For the
   valve specifically: does de-energising it *release* a held sample? If so the
   safe state for that line is "hold", not "off", and it must be excluded from
   the watchdog's drop-out set.
3. **ARM box control-voltage input — open-circuit voltage and input impedance.**
   Determines whether the output needs a buffer or can be driven through a
   series resistor.

Photographs that would settle the rest: the ARM box front and rear panels with
connectors and labels legible; the breakout terminal block where the 68-pin
cable lands, with wire labels visible; and the valve and cooler control inputs.

Note that `ARMSet` is **active low** in software — `DoDAQIO ARMSet, , False`
means *connected*. Whether that inversion is in the box or in the wiring is one
of the things the photographs should resolve.

## 3. Path A — a USB DAQ that speaks the same API

`Board.cls` calls the Measurement Computing Universal Library directly:
`cbAOut`, `cbAIn`, `cbDConfigBit`, `cbDBitOut`, `cbDBitIn`, `cbAInScan`,
`cbAOutScan`. MCC still sells USB hardware driven by that same library, so a
USB board is a driver-level swap rather than a code change.

### Part

| Board | AO | AI | DIO | Notes |
|---|---|---|---|---|
| **USB-3101** | 8 × 16-bit | none | 8 | Cheapest clean fit. Verify the AO range is software-selectable to unipolar 0–10 V so `RangeType0=100` is preserved. |
| **USB-1608GX-2AO** | 2 × 16-bit | 16 × 16-bit, 500 kS/s | 8 | Full superset of the 6030, including the 100 kS/s scan channels. Choose this if IRM or the 2G AF path might return. |

### Changes required

All in `VB6/settings/Paleomag_v3.INI`; no source edits.

| Key | From | To |
|---|---|---|
| `BoardName0` | `PCI-DAS6030` | the new board's name |
| `RangeType0` | `100` (`UNI10VOLTS`) | unchanged if the board supports unipolar; `1` (`BIP10VOLTS`) otherwise |
| `DOutPortType0` | `1` (`AUXPORT`) | `10` (`FIRSTPORTA`) on most USB boards |

`DOutPortType0` is the easy one to miss. `Board.cls:702` passes the port type
straight through to `cbDConfigBit`/`cbDBitOut`, so it is configurable — but USB
boards generally expose their DIO as `FIRSTPORTA`, not the `AUXPORT` the PCI
card used. If it is wrong the digital lines silently fail to configure.

On `BIP10VOLTS`, resolution over the 0–2.1 V span halves to 305 µV ≈ 1.5 mG,
still far finer than any field step in use. `SetBiasField` already clamps
negatives, so a bipolar DAC cannot be commanded below zero — but check what it
outputs at power-up before it is first written, which on some boards is
mid-scale rather than zero.

### Assessment

Fastest and lowest software risk: InstaCal, three INI keys, done in an
afternoon, calibration constants untouched, and `mcculw` makes the same board
reachable from RapidPy later.

It does not improve the overheat exposure. The DAC still holds its last value
through a host crash. If Path A is chosen, add an external timer relay in series
with the ARM enable line so the coil cannot stay energised past a fixed limit
regardless of what the PC is doing.

## 4. Path B — build it: Pico + 16-bit DAC

~$60 in parts. Delivers better resolution than the card, a real fail-safe, and
a documented serial protocol that suits RapidPy's typed-backend architecture
better than a proprietary 32-bit DLL does.

### Block diagram

```
   shielded room boundary
            |
  PC        |                      ARM box            ARM coils
 (USB) --[USB isolator]--+                 
            |            |
            |         [ Pico ]
            |            |  I2C/SPI
            |            +---> [ AD5693R 16-bit DAC ]
            |            |         internal 2.5 V ref, gain x2 -> 0-5 V
            |            |              |
            |            |         [ OPA192 buffer ] -- RC 16 Hz --> ARM Vctl
            |            |
            |            +---> [ 3x opto ] --> ARMSet          --> ARM relay
            |            |                     VacuumToggleA   --> valve
            |            |                     DegausserToggle --> cooler
            |            |
            |            +--10 Hz--> [ charge pump + MOSFET ] --> relay coil rail
            |                        (hardware watchdog: stops toggling,
            |                         rail collapses in ~200 ms)
            |
      keep the box on this side
```

Put the enclosure **outside** the shielded room and run only the DC lines
through the wall. A USB microcontroller is an RF source and its output is about
to drive a field coil next to a SQUID.

### Parts

| Part | Purpose | Approx |
|---|---|---|
| Raspberry Pi Pico | MCU, native USB CDC | $4 |
| AD5693R | 16-bit DAC, I²C, internal 2.5 V reference, ×2 gain → 0–5 V | $8 |
| OPA192 | Unity-gain output buffer, 5 µV offset, rail-to-rail | $5 |
| 1 kΩ + 10 µF | Output RC filter, 16 Hz corner | — |
| 4-ch opto-isolated output module | `ARMSet`, `VacuumToggleA`, `DegausserToggle`, one spare | $8 |
| ADuM3160 USB isolator | Galvanic break at the room boundary | $25 |
| 74HC123 or TPL5010 + MOSFET | Hardware watchdog on the relay coil rail | $3 |
| 5 V supply | Relay coil rail, separate from USB | $8 |
| Metal enclosure | Shielding, single-point ground | $15 |

**On the DAC.** 0–5 V at 16 bits is 76 µV, or 0.37 mG — twice the card's
resolution, with 2.4× headroom over the 2.1 V ceiling, and no ±12 V rail
needed because the ARM input never goes above 2.1 V. `AD5686R` is the SPI
equivalent if I²C is unwelcome; `MCP4725` (12-bit, widely stocked on breakouts)
is a defensible fallback at 1.2 mV ≈ 6 mG, still finer than any field step in
use, if sourcing is difficult. Do not use PWM into an RC, and do not use the
Pico's own PWM-based "analog out" — the ARM bias is a field-setting DC level,
and its noise lands directly in the measurement.

**On the digital outputs.** Whether these should be dry relay contacts or
open-collector TTL depends on what the ARM and vacuum boxes expect, which the
photographs will settle. The help file's language ("toggles relay") suggests
contacts. Opto-isolation is required either way: these are separate chassis
with their own grounds, and tying them together through a USB device invites a
ground loop straight into the magnetometer.

### Serial protocol

Line-oriented ASCII, `\n` terminated, so it can be driven from a terminal for
bring-up and fault-finding.

```
*IDN?           -> RAPID-AUX v1 <serial>
AO0 <volts>     -> OK | ERR <reason>
AO0?            -> <volts>
DO<n> <0|1>     -> OK | ERR <reason>          n in 0..3
DO<n>?          -> <0|1>
SAFE            -> OK
PING            -> PONG
STAT?           -> <safe|live> <ao0> <dobits> <uptime_ms> <last_fault>
```

Rules, all enforced in firmware:

- Boots into the safe state and stays there until commanded.
- `PING` must arrive at least every 1000 ms once any output is non-safe.
  Timeout → safe state, `last_fault = heartbeat`.
- USB disconnect or reset → safe state, immediately.
- `AO0` above a compile-time ceiling of 2.5 V is rejected outright. This is a
  second limit independent of `ARMVoltMax` in the INI — the point is that a
  corrupted INI or a bad edit cannot exceed it.
- Any unparseable input → safe state.

The safe state is DAC 0 V and `ARMSet` open. Whether the vacuum valve belongs
in that set depends on measurement 2 above; if de-energising it drops a sample,
it must be excluded and driven only by explicit command.

### Hardware watchdog

The firmware timeout covers a hung host. It does not cover hung firmware, so
the relay coil rail is gated in hardware: the Pico toggles a pin at 10 Hz into
a diode charge pump holding a MOSFET on. Stop toggling — crash, reset, halt,
unplugged — and the rail collapses within about 200 ms, dropping every relay
including `ARMSet`. Five components, no firmware in the path.

This is the design's main advantage over both the PCI card and Path A: the ARM
coil cannot remain energised through any single failure of the PC or its
software.

### VB6 changes

`Board.cls:43` already declares a protocol enum:

```vb
MCC_UL    = 1
ADWIN_COM = 2
```

Adding `USB_SERIAL = 3` is the whole insertion point. `AnalogOut`,
`AnalogIn`, `DigitalOut_MCC` and `DigitalOut_ADWIN` all branch on
`mvarCommProtocol`, so new `*_Serial` cases slot in beside them and every
caller — `frmIRMARM`, `frmVacuum`, `frmDAQ_Comm` — is untouched. The project
already uses MSComm throughout for the vacuum box and motors, so the transport
is an established pattern rather than a new dependency.

| File | Change |
|---|---|
| `VB6/Board.cls` | `USB_SERIAL = 3`; `AnalogOut_Serial`, `DigitalOut_Serial`, `AnalogIn_Serial`; dispatch cases |
| `VB6/modConfig.bas` | Read the aux COM port from `[COMPorts]` |
| new hidden form | MSComm instance plus a `Timer` sending `PING` every 300 ms |
| INI | `CommProtocol0=3`, aux port number |

### RapidPy changes

Follows the structure already established in Phase 1:

- `rapid_main/aux_io.py` — a typed `AuxOutputController` protocol
  (`set_arm_voltage`, `set_digital`, `safe`), a `SerialAuxBackend`, and an
  `UnavailableAuxBackend` matching the fail-closed pattern in
  `diagnostic_services.py`. No silent simulation fallback.
- Heartbeat owned by the worker thread, not the GUI thread.
- Tests against a fake transport, following `tests/acquisition_fakes.py`.

## 5. Recommendation

**Path B**, on the assumption that the outstanding measurements do not turn up
a surprise.

The scope has collapsed to one DC level and three bits. The build is a day's
work, it costs a tenth of the commercial board, it resolves finer than the card
it replaces, and it is the only option that fixes the overheat exposure rather
than inheriting it. The serial protocol is also a better fit for RapidPy than
`cbw32.dll`, which would otherwise drag an InstaCal dependency into the Python
rewrite.

**Path A** is the right answer instead if there is any real prospect of IRM or
the 2G AF path returning — the 100 kS/s channels come back for free, and no
firmware exists to maintain. Pair it with an external timer relay on the ARM
enable line.

Either way, the three measurements in section 2 come first.
