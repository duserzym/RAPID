# Arduino aux output bench test

Toggles D7, D8, D12 and D13 individually from a browser page so each can be
metered against GND. The point is to confirm a 5 V AVR board can stand in for
the PCI-DAS6030 AUXPORT lines that carry the vacuum controls — see
`docs/pci_das6030_usb_replacement_2026-09-03.md`.

## Keep the Arduino above COM10

The Arduino Uno first enumerated onto **COM3**, which the RAPID configuration
already uses:

```
[COMPorts]
COMPortChanger = 3
```

The PCI serial card claims COM3 through COM10, and those numbers are assigned
to the changer, up/down, turning, susceptibility and IRM. An Arduino sitting on
top of one of them will collide as soon as both are open.

It was moved to **COM11** on 2026-09-04 and COM3 returned to the serial card.
If it ever re-enumerates, put it back above COM10: Device Manager → Ports
(COM & LPT) → Arduino Uno → Properties → Port Settings → Advanced → COM Port
Number.

The card occupies COM3-COM10. Of those the INI claims 3 (changer), 4 (changerY),
5 (up/down), 6 (turning), 7 (susceptibility) and 9 (IRM), with COM1 the squids;
8 and 10 are spare but still belong to the card.

## Load the sketch

Open `aux_test.ino` in the Arduino IDE, select the board and the port, upload.

The IDE on this machine is **1.0.5-r2**, which predates CLI upload — its
`arduino.exe` ignores `--upload` and just opens the GUI. `build_upload.sh`
compiles against the bundled 1.0.5 core and flashes with the bundled avrdude,
using the recipe from `boards.txt`:

```bash
bash hardware/arduino-aux-test/build_upload.sh COM13
```

Confirm it took: the board should answer `*IDN?` with `RAPID-AUX-TEST v1`. A
board running something else will not answer at 115200 at all — the first one
tried here was still carrying an Adafruit LSM303 demo at 9600 baud, which reads
as framing garbage rather than as a wrong reply.

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

## Re-enable the serial card when you are done

Isolating a COM number collision here meant disabling the PCI serial card's
ports. **The RAPID software needs them back**: `[COMPorts]` assigns COM1 to the
squids and COM3-COM10 to the changer, changer Y, up/down, turning,
susceptibility and IRM. Re-enable them in Device Manager, then confirm the
device map is right before starting Paleomag:

```bash
python hardware/arduino-aux-test/bench_server.py --list
```

`\Device\Sbser0` must read COM3 again. If it does not, its driver has a stale
registration and needs the device disabled and re-enabled, or a reboot.
