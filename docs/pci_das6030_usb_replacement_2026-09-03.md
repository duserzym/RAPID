# Replacing the PCI-DAS6030 with a USB device — 2026-09-03

The lab computer no longer has a PCI slot. This document specifies what has to
cross the gap and how to build or buy it.

**Scope reduced, same day.** ARM is being abandoned. That removes the only
analog channel and leaves **two digital outputs**. The ARM analysis is kept in
[appendix A](#appendix-a--if-arm-is-ever-restored) because restoring it later
changes the design materially.

Nothing here has been built or wired. Two questions are outstanding — see
[Before anything is wired](#6-before-anything-is-wired).

## 1. What the card is actually doing

Every analog and digital operation funnels through one dispatcher,
`frmDAQ_Comm.DoDAQIO` in `VB6/frmDAC_Comm.frm:578`, which calls into
`VB6/Board.cls`. There is no second path to the hardware, so enumerating its
callers gives a closed inventory. Read against
`VB6/settings/Paleomag_v3.INI`:

| INI channel | Physical | Function | State |
|---|---|---|---|
| `VacuumToggleA` | AUXPORT bit 6 | Valve, quartz rod → vacuum | **required** |
| `DegausserToggle` | AUXPORT bit 2 | Degausser cooling air | **required** |
| `MotorToggle` | AUXPORT bit 5 | Vacuum pump request | not needed — building vacuum |
| `ARMVoltageOut` | `AO-0-CH0` | ARM bias voltage | abandoned |
| `ARMSet` | AUXPORT bit 0 | TTL relay, ARM box → coils | abandoned |
| `IRMVoltageOut` | `AO-0-CH1` | IRM charge setpoint | module off |
| `IRMFire`, `IRMTrim` | AUXPORT bits 1, 3 | IRM fire and trim | module off |
| `AnalogT1`, `AnalogT2` | `AI-0-CH3/4` | Coil temperatures | module off |
| `ALTAFMONITOR` | `AI-0-CH1` @ 100 kS/s | AF LC current monitor | unreachable |
| `IRMMONITOR` | `AI-0-CH4` @ 100 kS/s | IRM discharge monitor | module off |
| `VacuumToggleB` | AUXPORT bit 7 | — | declared, never driven |

The IRM and temperature rows are off because `[Modules]` sets
`EnableAxialIRM`, `EnableTransIRM`, `EnableIRMBackfield`, `EnableIRMMonitor`,
`EnableIRMReturn`, `EnableT1` and `EnableT2` all to `False`.

`ALTAFMONITOR` needs a second look, because `EnableAltAFMonitor = True` makes it
appear live. Its only consumer is `frmAF_2G.frm`, which `modAF.bas:54` reaches
only when `AFSystem = "2G"`. The site runs `AFSystem = ADWIN`, so the AF ramp,
monitor and relay TTL all sit on the ADwin-light-16 over its own link. The flag
is vestigial.

Setting `EnableARM = False` makes `SetBiasField` return at its first line
(`frmIRMARM.frm:3370`), so neither ARM channel is ever touched.

**What remains is two relay-grade digital outputs.** No analog, no buffered
acquisition, no timing requirement.

### The serial path to the vacuum box is dead

`frmVacuum.ValveConnect` looks like it has two independent control paths — it
sends `O`/`C` and `10VFF`/`10V00` over `MSCommVacuum` *and* toggles the DAQ
line. On this system only the second one does anything:

```
[COMPorts]
COMPortVacuum = -1
```

`Connect` (`frmVacuum.frm:208`) requires `COMPortVacuum > 0`, so the port never
opens; `SendCommand` (`frmVacuum.frm:391`) then falls through its `Else` branch,
which is *also* guarded by `COMPortVacuum > 0` and so does not even raise the
"Vacuum Comm Port Not Open" message. The serial commands are silently discarded.

This matters because it rules out the cheapest possible answer. A USB-to-serial
adapter on the vacuum box would not restore valve control — the box is not being
driven that way here. **The DAQ digital line is the only control**, and it has
to be reproduced.

## 2. The hazard, which is not what it was

With ARM in scope the governing risk was thermal: `ARMTimeMax = 600` zeroes the
bias field after ten minutes because the coil overheats, and that timer lives
entirely in the VB6 process. The design conclusion was a watchdog that
de-energises everything when the host stops talking.

**Dropping ARM inverts that.** The remaining risk is mechanical: the vacuum
valve is what holds a sample on the quartz rod. The application treats losing it
as serious enough to gate behind an operator prompt —

> `frmMeasure.frm:1410` — "You are logging out, remove the sample before the
> vacuum switch off."

— and `frmMagnetometerControl` carries a "Keep the vacuum on" checkbox so an
operator can prevent the automatic release. If vacuum drops while a sample is
attached, the sample falls into the magnetometer bore.

So a fail-open watchdog copied from the ARM design would *cause* the accident it
was meant to prevent. For this scope the correct failure behaviour is:

| Line | On loss of USB or host | Why |
|---|---|---|
| `VacuumToggleA` | **hold** current state | Releasing drops the sample into the bore |
| `DegausserToggle` | hold, or default on | Air on is the benign state during AF |

Neither line benefits from a watchdog. Both benefit from a relay that keeps its
contact position when the controlling logic disappears.

### Wire through the normally-closed contact

The cheap and robust way to get that: wire the valve through the relay's **NC**
contact, so the de-energised relay holds vacuum and the energised relay releases
it. A USB dropout, a host crash, or an unplugged box then leaves the sample
held — the only lost capability is the ability to *release*, which fails safe.

That inverts the logic the software expects, and the codebase already has the
pattern for exactly this. `TrimOnOff` (`frmIRMARM.frm:3652`) reads a
`TrimOnTrue` flag from the INI and returns `Not TrimOn` when a line is wired
active-low, with a comment explaining that the wiring convention differs per
system. Add the equivalent — `VacuumOnTrue` — rather than inverting silently in
firmware, so the convention is visible in the settings file where someone
debugging the valve will find it.

## 3. Options

The software work is now the dominant cost, and it is the same for every option
that is not a Measurement Computing board.

| | Hardware | Firmware | VB6 changes | Approx |
|---|---|---|---|---|
| **A** — USB relay module | 2-ch USB relay, dry contacts | none | `USB_SERIAL` case | ~$25 |
| **B** — Pico + optos | Pico + 2 opto outputs | small | `USB_SERIAL` case | ~$15 |
| **C** — MCC USB DIO | USB-1024LS or USB-DIO24 | none | none — three INI keys | ~$150–200 |

**Recommended: A.** At two dry-contact outputs there is nothing for a
microcontroller to add. An off-the-shelf USB relay module presents as a COM port
with a trivial text protocol, gives real relay contacts — almost certainly what
the vacuum box wants — and needs no soldering and no firmware to maintain.
Choose one whose relays are de-energised at power-up and wire the valve through
NC per section 2.

**B** is the fallback if the boxes turn out to want logic levels rather than
contacts, or if a spare channel and a defined power-up sequence are worth the
build.

**C** buys away the VB6 work entirely: `Board.cls` already calls the Universal
Library, so it is three INI keys and no source edits — `BoardName0`, and
`DOutPortType0` from `1` (AUXPORT) to `10` (FIRSTPORTA), which USB boards use.
That last key is the easy one to miss; if it is wrong the digital lines silently
fail to configure. Worth it only if avoiding legacy VB6 edits is worth ~$150,
and note it still needs the NC wiring, since an MCC board's DIO powers up as
high-impedance inputs.

## 4. Software changes

Common to A and B.

`Board.cls:43` already declares a protocol enum:

```vb
MCC_UL    = 1
ADWIN_COM = 2
```

Adding `USB_SERIAL = 3` is the whole insertion point. `DigitalOut_MCC` and
`DigitalOut_ADWIN` branch on `mvarCommProtocol`, so a `DigitalOut_Serial` slots
in beside them and every caller — `frmVacuum`, `frmDAQ_Comm` — is untouched. The
project already uses MSComm throughout for the changer, motors and squids, so
the transport is an established pattern rather than a new dependency.

| File | Change |
|---|---|
| `VB6/Board.cls` | `USB_SERIAL = 3`; `DigitalOut_Serial`; dispatch case |
| `VB6/modConfig.bas` | Read the aux COM port from `[COMPorts]`; read `VacuumOnTrue` |
| `VB6/frmVacuum.frm` | Apply `VacuumOnTrue` the way `TrimOnOff` applies `TrimOnTrue` |
| new hidden form | MSComm instance for the aux device |
| INI | `CommProtocol0=3`, aux port number, `VacuumOnTrue` |
| `rapid_main/aux_io.py` | New. Typed `AuxOutputController` protocol, `SerialAuxBackend`, and an `UnavailableAuxBackend` matching the fail-closed pattern in `diagnostic_services.py` — no silent simulation fallback |
| `tests/` | Fake transport following `acquisition_fakes.py` |

No heartbeat and no watchdog, per section 2. State is held in the relay, not in
the host.

## 5. Inside the control box

A photograph of the chassis where the 68-pin cable lands (2026-09-04) settles
the physical approach.

The cable terminates on a passive breakout, **Vertical System / Measurement
Computing Interface, 100-0627 Rev A**, which fans the DAQ signals into three
small headers along its lower edge. A second board, **Vertical Systems ARM,
100-0629**, carries the `ARMSet` relay; with ARM abandoned that board and
whatever feeds it are out of scope. The rest of the chassis is a mains inlet,
a transformer, a 1600 VDC film capacitor and a switching supply.

**If the vacuum lines come out on their own header, the replacement never
touches the 68-pin connector.** Unplug that header and land the relay contacts
on the same pins. The interface board, the ARM board and the power section stay
as they are, and the change reverses by plugging the original connector back in.

### Safety

Mains at the inlet, a transformer, and a 1600 VDC-rated capacitor. Power off is
not the same as discharged: unplug, wait, and verify across the capacitor
terminals before probing.

### Identifying the wires

A five-pin header is consistent with the `[Vacuum]` INI section, which lists
exactly three consecutive AUXPORT bits — `MotorToggle=5`,
`VacuumToggleA=6`, `VacuumToggleB=7` — so three signals plus ground and
supply. Of those, only bit 6 is needed: bit 5 is the pump, replaced by building
vacuum, and bit 7 is never driven anywhere in the code. `DegausserToggle` is
bit 2 and will be on one of the other two headers.

Do not derive the mapping from the MCC 68-pin pinout. The Vertical Systems
board defines it, and it need not pass signals straight through.

The unambiguous method uses the old computer, which still has the card:

1. Confirm `EnableVacuum = True` and that the program is not in no-comm mode.
2. Open the vacuum form and click Vacuum Connect on and off while metering each
   candidate wire against ground. The wire that swings is bit 6.
3. Repeat with the degausser cooler buttons to find bit 2.
4. Record the idle voltage and identify the supply pin. Open-collector lines
   pulled to +5 V mean the relay only has to short signal to ground; lines
   switching 12 or 24 V constrain the contact rating.

Also trace where each cable physically lands. A header whose cable terminates
on the ARM board (100-0629) is the ARM relay drive, not vacuum.

## 6. Before anything is wired

Section 5 covers identifying the wires. One question remains, and it decides
the wiring rather than the parts.

**Is the valve normally open or normally closed?** The test needs no
instrument: with the control connector unplugged and building vacuum supplied,
is the rod under vacuum or vented? Vented means the solenoid energises to hold,
and the NC wiring in section 2 is mandatory. Under vacuum means the hardware
already fails safe and the relay can be wired the obvious way.

A secondary question, only if the tap point turns out to be outside the
chassis: an 11-pin circular connector was reported as carrying the vacuum
controls. If the replacement lands there instead of on the interface board
header, check the open-circuit voltage on each pin and source the mating part,
so the swap stays a plug change with no disturbance to the existing harness.
The vintage suggests an Amphenol/Bendix MS or a Cinch-Jones plug.

## Appendix A — if ARM is ever restored

Restoring ARM re-introduces one analog output and changes the safety analysis
back, so it is not a small addition.

From `[ARM]` in the live INI: `ARMVoltGauss = 0.208` V/G, `ARMVoltMax = 2.1` V,
`ARMMax = 10.1` G, `ARMTimeMax = 600` s. `SetBiasField`
(`VB6/frmIRMARM.frm:3363`) zeroes the DAC, waits 0.5 s, closes `ARMSet`, waits
0.5 s, writes `ARMVoltGauss × Gauss`, then waits 1 s to settle — so the
requirement is 0–2.1 V DC written about twice a second and then held. The card's
16-bit DAC on `UNI10VOLTS` resolves 152 µV, or 0.73 mG.

A Pico with an `AD5693R` — 16-bit, I²C, internal 2.5 V reference at ×2 gain —
covers 0–5 V at 76 µV (0.37 mG) with no ±12 V rail and no gain stage, buffered
by an `OPA192` through a 16 Hz RC.

Use an external DAC, not whatever the board offers. Arduino's `analogWrite()`
is PWM on every AVR board — an Adafruit Metro 328 and a Pico have no DAC at
all — and the boards that do (Metro M0, one 10-bit DAC; Metro M4, two 12-bit)
reference the 3.3 V rail rather than a precision source, so accuracy and drift
track the regulator. A 12-bit DAC over 3.3 V is 806 µV, about 3.9 mG at this
calibration, against the card's 0.73 mG. The ARM bias is a field-setting DC
level and its noise lands directly in the measurement.

Restoring ARM also restores the thermal hazard, and with it the case for a
hardware watchdog on the ARM relay — which must be scoped to the ARM line only,
never the vacuum valve, for the reason in section 2.

Whichever way it goes, keep the box outside the shielded room and run only the
DC lines through the wall. A USB microcontroller is an RF source and its output
would be driving a field coil next to a SQUID.
