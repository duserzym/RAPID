# Arduino aux output bench test

Toggles D7, D8, D12 and D13 individually from a browser page so each can be
metered against GND. The point is to confirm a 5 V AVR board can stand in for
the PCI-DAS6030 AUXPORT lines that carry the vacuum controls — see
`docs/pci_das6030_usb_replacement_2026-09-03.md`.

## Move the Arduino off COM3 first

On this computer the Arduino Uno enumerated onto **COM3**, which the RAPID
configuration already uses:

```
[COMPorts]
COMPortChanger = 3
```

The PCI serial card claims COM3 through COM10, and those numbers are assigned
to the changer, up/down, turning, susceptibility and IRM. An Arduino sitting on
top of one of them will collide as soon as both are open.

Device Manager → Ports (COM & LPT) → Arduino Uno → Properties → Port Settings →
Advanced → COM Port Number. Pick **COM11 or higher**. Ports marked "in use" that
belong to absent hardware can be reused, but do not take one the INI lists.

## Load the sketch

Open `aux_test.ino` in the Arduino IDE, select the board and the port, upload.

It answers a line protocol at 115200 8N1:

```
*IDN?           -> RAPID-AUX-TEST v1
PIN <n> <0|1>   -> OK <n> <0|1>        n in {7, 8, 12, 13}
ALL 0           -> OK ALL 0
STAT?           -> 7=0 8=0 12=0 13=0
```

Every managed pin is driven LOW in `setup()`, with the PORT latch cleared before
the driver is enabled so no pin glitches high on the way up.

## Run the page

Web Serial needs Chrome or Edge and a secure context, so serve it from
localhost rather than double-clicking the file:

```bash
python -m http.server 8000 --directory hardware/arduino-aux-test
```

Then open **http://localhost:8000/control.html** in Chrome and click Connect.

The page will not work as a published artifact or inside an embedded preview
pane — those sandboxes do not carry the `serial` permission.

Opening the port asserts DTR, which resets the Uno, so the page waits about two
seconds for the bootloader before sending anything. It then sends `ALL 0` and
reads the state back. A switch only shows HIGH once the board has confirmed it,
so the display reflects the hardware rather than the click.

## Measuring

Black probe on any GND pin of the POWER header, red probe on the pin under test
— not pin to pin.

| | Expected |
|---|---|
| HIGH | 4.8 – 5.1 V on USB power |
| LOW | under 0.1 V |
| D13 | also lights the onboard LED |

Under 4.5 V on a HIGH means the pin is loaded or the 5 V rail is sagging. A few
hundred millivolts on a LOW means something else is still driving the line.

Per-pin current is 20 mA as a design figure and 40 mA absolute, with roughly
100 mA per port group and 200 mA for the whole chip. That is enough for a TTL
input or an opto-isolator LED, and not enough for a relay coil or a solenoid —
those want a ULN2803 or an opto-isolated relay module in between.

## What this does not test

Boot state. The pins are high-impedance from power-on until `setup()` runs,
which on a stock bootloader is one to two seconds after every reset and every
USB reconnect — including the reset this page causes when it opens the port.
Watch the meter while clicking Connect to see it. On the real system that window
is what decides whether a reset drops the sample, and the fix is a pull-down or
pull-up on the wiring, not in the sketch.
