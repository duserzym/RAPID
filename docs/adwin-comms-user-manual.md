# RapidPy ADwin Communication Tester — User Manual

> **Status:** Entry-level hardware test tool. Use this app to verify ADwin board communication, relay wiring, and DAC/ADC signal paths **before** connecting amplifiers or coils.

---

## Table of Contents

1. [Overview](#1-overview)
2. [System Requirements & Driver Installation](#2-system-requirements--driver-installation)
3. [Safety — What This App Is For](#3-safety--what-this-app-is-for)
4. [Launching the App](#4-launching-the-app)
5. [Application Layout](#5-application-layout)
6. [Left Panel — Controls (Card by Card)](#6-left-panel--controls-card-by-card)
   - [Board Configuration](#board-configuration)
   - [Relay Test (Digital Outputs)](#relay-test-digital-outputs)
   - [Direct DAC / ADC](#direct-dac--adc)
   - [Process Control](#process-control)
   - [PAR / FPAR I/O](#par--fpar-io)
   - [Sine Loopback Test](#sine-loopback-test)
   - [Console](#console)
7. [Right Panel — Loopback Plot](#7-right-panel--loopback-plot)
8. [Typical Test Workflows](#8-typical-test-workflows)
9. [ADwin Driver Installation](#9-adwin-driver-installation)
10. [Python Package vs ctypes: How This App Talks to ADwin](#10-python-package-vs-ctypes-how-this-app-talks-to-adwin)
11. [Building from Source](#11-building-from-source)
12. [Troubleshooting](#12-troubleshooting)

---

## 1. Overview

The **ADwin Communication Tester** is the entry-level RapidPy module for working with the Jäger ADwin real-time data acquisition board used in the RAPID paleomagnetics system.

The ADwin board serves two key roles in RAPID:

| Role | Mechanism |
|---|---|
| **AF sine wave generation** | ADbasic process outputs sine wave to DAC → power amplifier → AF coils |
| **Relay control** | Digital output bits (0–5) drive relay boxes that select axial/transverse coils and IRM/AF/ARM modes |

This app lets you test **both** before connecting any high-voltage hardware:

1. **Board boot & connection** — verify the board responds, check the version.
2. **Relay testing** — toggle digital output bits, physically verify the relay box LEDs respond.
3. **Direct DAC/ADC testing** — write a voltage to a DAC channel, read it back from an ADC channel (with a loopback cable).
4. **Sine loopback test** — generate a configurable sine wave on DAC, read it back on ADC, plot both traces in real time.
5. **Process control** — load, start, stop, and clear ADbasic (.abp) processes.
6. **PAR/FPAR read/write** — inspect and set process parameters for debugging.

---

## 2. System Requirements & Driver Installation

| Requirement | Detail |
|---|---|
| OS | Windows 10 / 11 (ADwin driver is Windows-only) |
| ADwin hardware | ADwin Gold II (or compatible) connected via USB or Ethernet |
| ADwin driver | **`adwin32.dll`** must be installed on the PC — see [Section 9](#9-adwin-driver-installation) |
| Python | 3.10+ with `paleomag` conda environment (dev path) |

The app will launch without the driver installed, but all hardware buttons will be disabled and a warning appears in the console.

---

## 3. Safety — What This App Is For

> **This app is designed to be used WITHOUT the power amplifier or AF coils connected.**

The sine loopback test generates voltages directly on the ADwin DAC output pin. With nothing connected, or with a short DAC→ADC cable for testing, this is completely safe.

**Do NOT** connect the AF coil amplifier to the DAC output while using this app. The AF ramp process in the full `af_tuner` module manages the amplifier safely. This app is purely for testing the board electronics.

For relay testing, the digital output bits drive 5V logic signals to the relay box — these are safe to toggle freely.

---

## 4. Launching the App

**From conda environment (development):**
```powershell
conda activate paleomag
cd RapidPy\adwin_comms
python main.py
```

**From built executable:**  
Double-click `RapidPyADWin.exe` in the application folder.

Settings are saved to `~/.rapidpy_adwin_comms.json` and restored on next launch.

---

## 5. Application Layout

The window is divided by a resizable horizontal splitter:

| Side | Content |
|---|---|
| **Left** (scrollable) | Seven control cards: board, relay, DAC/ADC, process, PAR/FPAR, signal gen, console |
| **Right** | Live loopback signal plot (dark background, two traces) |

Drag the splitter handle to give more space to the plot or the controls.

---

## 6. Left Panel — Controls (Card by Card)

### Board Configuration

Before any hardware operation, the board must be **booted**.

| Control | Description |
|---|---|
| **Board #** | ADwin device number (default: 1). Change only if multiple boards are installed. |
| **Boot file** | The firmware file to load: `ADwin9.btl` for Gold II, `ADwin5.btl` for older boards, etc. |
| **Bin folder** | Directory containing the `.btl` and `.abp` files. Leave blank to use the current working directory. |
| **Boot Board** | Boots the ADwin firmware. First checks if already booted (via `ADTest_Version`); if not, runs `ADboot`. The status label turns green on success. |

The boot step is required once per session (or after the board loses power). It loads the real-time operating system firmware onto the board.

---

### Relay Test (Digital Outputs)

The ADwin board has 6 digital output bits (0–5) that drive the relay switching logic in RAPID:

| Bit | Default label | Function |
|---|---|---|
| Bit 0 | Axial Relay | Routes power amplifier output to the axial AF coil |
| Bit 1 | Trans Relay | Routes power amplifier output to the transverse AF coil |
| Bit 2 | IRM/AF Select | Switches between IRM capacitor discharge and AF sine mode |
| Bit 3 | ARM Polarity | Controls DC bias field polarity for ARM acquisition |
| Bit 4 | Bit 4 | (spare — lab-specific use) |
| Bit 5 | Bit 5 | (spare — lab-specific use) |

Each bit is represented by a toggle button. When toggled ON (checked), that bit is set HIGH; the full bitmask is written to `Set_Digout()` immediately.

**Read** — reads the current digout register from the board and syncs all toggle buttons.  
**All OFF** — sets the entire digout register to 0x00 (all relays off). Always do this before leaving the app.

> The raw digout word is shown in hex (e.g. `Digout: 0x03` = bits 0 and 1 both ON = both coil relays active).

**Test procedure:** Toggle each bit in isolation, then walk to the relay box and confirm the corresponding LED changes. This verifies the relay box wiring.

---

### Direct DAC / ADC

Write a voltage directly to any DAC channel, or read a voltage from any ADC channel, without a loaded process.

| Control | Description |
|---|---|
| **DAC Ch** | DAC output channel number (1–8) |
| **V** | Voltage to write (±10.0 V range) |
| **Write DAC** | Writes the voltage. Internally: `count = int(V × 32768 / 10) + 32768`, then calls `Set_DAC(chan, count, device)`. |
| **ADC Ch** | ADC input channel number (1–8) |
| **Read ADC** | Reads the ADC and displays the voltage. Internally: reads 16-bit count, converts `V = (count − 32768) × 10 / 32768`. |

The conversion formula uses the ±10 V range convention consistent with the ADbasic globals (`MAX16BIT = 65536`, `MAX15BIT = 32768`).

**Loopback cable test:** Connect DAC Ch 1 output → ADC Ch 1 input with a short BNC cable. Write 2.5 V to DAC 1, read ADC 1 — you should get approximately 2.5 V back.

---

### Process Control

Manage ADbasic real-time processes (.abp files) on the board.

| Control | Description |
|---|---|
| **.abp path** | Path to the compiled ADbasic process file. Browse with the **…** button. |
| **Process #** | Which process slot to start/stop (1–10). ADwin supports up to 10 concurrent processes. |
| **Load** | Loads the .abp file into the board (`ADBload`). Must boot first. |
| **Start** | Starts the process in the selected slot (`ADB_Start`). |
| **Stop** | Stops the process (`ADB_Stop`). |
| **Clear All** | Stops and clears all 10 process slots. Use before loading a new process. |

The AF ramp process (`AF_Ramp_System.abp` or `sineout.T91` for the older system) can be loaded here for manual testing.

---

### PAR / FPAR I/O

Read or write individual PAR (integer) or FPAR (float) process parameter registers.

These map directly to `Set_ADBPar` / `Get_ADBPar` and `Set_ADBFPar` / `Get_ADBFPar` in the DLL.

| Control | Description |
|---|---|
| **PAR[idx]** | Integer parameter at index idx (1–80). The AF ramp process uses PAR 31–38 for ramp control. |
| **FPAR[idx]** | Float parameter at index idx (1–80). FPAR 31–37 carry slope, frequency, peak voltage etc. |

Clicking **Read** updates the spinbox value from the board. Clicking **Write** pushes the spinbox value to the board register.

This is useful for inspecting a running process's state (e.g. checking ACPHASE in PAR 4 to see where the ramp is), or for setting parameters before starting a process manually.

---

### Sine Loopback Test

Generates a sine wave at configurable frequency, amplitude, and duration, writing it sample-by-sample to the DAC and simultaneously reading back from the ADC.

| Control | Description |
|---|---|
| **Frequency** | Sine wave frequency in Hz (0.1–10000 Hz) |
| **Amplitude** | Peak voltage (0.001–9.999 V). Keep below 1 V for initial tests. |
| **Duration** | How long to run the test (seconds) |
| **IO Rate** | Target sample rate in Hz (1–2000). PC-side timing — actual rate may be lower than requested due to OS scheduling. |
| **DAC Ch / ADC Ch** | Which channels to use. Connect DAC out → ADC in for loopback. |
| **Run Sine Loopback** | Starts the background worker thread. |
| **Stop** | Signals the worker to stop early. DAC is returned to 0 V. |

The worker runs in a background thread so the UI stays responsive. Progress is visible in the plot and status label.

> **Note:** This is PC-timed, not ADwin real-time. Timing jitter is 1–5 ms typical (Windows scheduler). For real-time AF generation, the full `af_tuner` module loads an ADbasic process which runs at 10 kHz deterministically on the board.

---

### Console

All operations are logged here with timestamps (`[HH:MM:SS]`). Board boot results, digout writes, DAC/ADC values, process load/start/stop events, and errors all appear here. Use the **Clear** button to reset it.

---

## 7. Right Panel — Loopback Plot

The dark-background plot shows two traces updated in real time during a sine loopback test:

| Trace | Colour | Content |
|---|---|---|
| **DAC out** | Bright lime `#4DFF91` | The commanded sine voltage written to the DAC |
| **ADC in** | Gold `#FFCD34` | The voltage measured back from the ADC |

Both traces show voltage vs elapsed time in seconds.

- **Pan**: right-click + drag
- **Zoom**: scroll wheel (both axes simultaneously)
- **Clear Plot**: removes all accumulated data without stopping the test

Comparing the two traces lets you verify:
- Correct cable connection (ADC reads what DAC writes)
- Phase delay (ADC lags DAC by the cable/circuit propagation time)
- Amplitude match or attenuation (if going through any passive components)
- Noise floor (ADC reads noisy signal → check grounding or shielding)

---

## 8. Typical Test Workflows

### Workflow 1 — Verify board communication
1. Set **Board #** and **Bin folder** to match your ADwin installation.
2. Click **Boot Board**.
3. Confirm the status label turns green: "● Booted ✓".
4. Check the console for the success message.

### Workflow 2 — Relay box verification
1. Boot the board (Workflow 1).
2. Click **Bit 0 (Axial Relay)** to toggle it ON.
3. Walk to the relay box and confirm the axial relay LED is lit.
4. Click **Bit 0** again to toggle OFF; confirm LED goes out.
5. Repeat for each bit in sequence.
6. Click **All OFF** when done.

### Workflow 3 — DAC/ADC loopback
1. Boot the board.
2. Connect a short BNC cable: DAC 1 output → ADC 1 input.
3. In **Direct DAC / ADC**, set DAC Ch=1, V=2.500, click **Write DAC**.
4. Set ADC Ch=1, click **Read ADC**. Expected result: ~2.500 V.
5. Test several voltages: −5.0 V, 0.0 V, +5.0 V.

### Workflow 4 — Sine signal loopback plot
1. Boot the board.
2. Connect DAC 1 → ADC 1 with a cable.
3. In **Sine Loopback Test**, set Frequency=100 Hz, Amplitude=1.0 V, Duration=5.0 s.
4. Click **Run Sine Loopback**.
5. Watch the plot: both traces should overlay (lime=commanded, gold=measured).
6. After the run, compare amplitudes and phases to verify hardware integrity.

---

## 9. ADwin Driver Installation

The ADwin system uses `adwin32.dll` (Windows 32-bit COM DLL) to communicate with the board via USB or Ethernet. This DLL is supplied by **Jäger Messtechnik**, the manufacturer of ADwin hardware — it is not available on PyPI and cannot be replaced by any Python package.

> **ADwin product page:** [www.adwin.de/us/products/products.html](https://www.adwin.de/us/products/products.html)  
> **Python for ADwin info:** [www.adwin.de/us/products/python.html](https://www.adwin.de/us/products/python.html)

### Dependency chain

```
ADwin hardware (Gold II / Pro II)
        ↓
ADwin driver installer  (from adwin.de → Downloads)
        ↓  installs adwin32.dll → C:\Windows\System32\
RapidPy / adwin_af.py   (ctypes.WinDLL("adwin32.dll"))
        ↓
RapidPyADWin.exe  /  af_tuner  /  full RAPID orchestrator
```

**A pip-installable Python package alone is not sufficient.** The driver installer is required on every PC that will communicate with the board.

### What is needed
- The **ADwin driver/runtime package** from Jäger Messtechnik.
- Typically installed as part of the **ADwin Development Environment** (includes ADbasic IDE and USB/Ethernet drivers).
- After installation, `adwin32.dll` is present in `C:\Windows\System32\` and the boot file `ADwin9.btl` (or equivalent for your board model) is placed in the ADwin program directory.

### Installation steps
1. Go to [www.adwin.de/us/products/products.html](https://www.adwin.de/us/products/products.html) and navigate to **Support → Downloads**.
2. Download the ADwin software package appropriate for your board model (Gold II, Pro II, etc.).
3. Run the installer as Administrator.
4. Accept the USB driver installation when prompted (registers the board's USB device with Windows).
5. Note the install path — the `.btl` boot file location is needed in the Board Configuration panel.
6. Verify the DLL is accessible by running in a command prompt:
   ```
   python -c "import ctypes; ctypes.WinDLL('adwin32.dll'); print('OK')"
   ```
   If it prints `OK`, communication is ready. If it raises `OSError`, the driver is not installed or the DLL is not on the system path.

### Do you need to install the driver to use this Python app?

**Yes, to communicate with hardware.** Without `adwin32.dll`, the app launches in degraded mode: all hardware buttons are disabled and the console shows:

```
[WARNING] adwin32.dll not found: ...
  → Hardware buttons disabled.  Install ADwin driver to enable.
```

This is intentional — the app is safe to open on any machine for inspection or documentation purposes.

### Python package alternative?

There is a PyPI package called `adwin` (and similar community wrappers) that wrap the same `adwin32.dll` — they do not bypass the driver requirement, they simply provide a higher-level Python API on top of it. Since `rapidpy_common/adwin_af.py` already implements all necessary ctypes bindings directly, adding the PyPI package would be redundant. Our implementation is lighter, fully under our control, and matches the exact function signatures used by the legacy VB6 RAPID code.

---

## 10. Python Package vs ctypes: How This App Talks to ADwin

The existing VB6 RAPID code declares functions from `adwin32.dll` using VB6's `Declare Function` mechanism:

```vb
Declare Function Get_ADC Lib "adwin32.dll" (ByVal NADC As Integer, ByVal DeviceNo As Integer) As Long
Declare Function Set_DAC Lib "adwin32.dll" Alias "Set_DAC" (ByVal ndac As Integer, ByVal value As Long, ByVal DeviceNo As Integer) As Integer
```

Our Python equivalent in `rapidpy_common/adwin_af.py` uses ctypes:

```python
self.Get_ADC = self._dll.Get_ADC
self.Get_ADC.argtypes = [ctypes.c_int, ctypes.c_int]
self.Get_ADC.restype = ctypes.c_long

self.Set_DAC = self._dll.Set_DAC
self.Set_DAC.argtypes = [ctypes.c_int, ctypes.c_long, ctypes.c_int]
self.Set_DAC.restype = ctypes.c_int
```

This is **exactly equivalent** to the VB6 bindings. Both paths:
1. Load `adwin32.dll` from the Windows DLL search path.
2. Call the same functions by name.
3. Pass the same parameter types (int channel, long value, int device number).

The `AdwinAFController` class in `rapidpy_common/adwin_af.py` wraps these calls with Python-friendly methods:

```python
ctrl.set_dac(channel=1, voltage=2.5)   # writes 2.5 V to DAC 1
v = ctrl.get_adc(channel=1)            # reads ADC 1 in volts
ctrl.set_digout(0b00000011)            # sets bits 0 and 1 HIGH
ctrl.boot_board()                      # boots the board firmware
ctrl.load_process()                    # loads the .abp process file
ctrl.run_ramp(request)                 # executes a full AF ramp
```

---

## 11. Building from Source

```powershell
conda activate paleomag
cd RapidPy\adwin_comms
build_windows.bat
```

The batch script:
1. Generates the icon (`tools\generate_icon.py`).
2. Runs PyInstaller in one-file mode → `dist\RapidPyADWin.exe`.

> **Note:** `adwin32.dll` is **not** bundled in the exe — it must be installed on the target PC. This is intentional: the DLL is a system-level driver component.

---

## 12. Troubleshooting

### "adwin32.dll not found" in console at startup

Install the ADwin driver package from the manufacturer. See [Section 9](#9-adwin-driver-installation).

### Boot Board fails

- Verify the ADwin board is powered on and connected via USB or Ethernet.
- Check **Board #** matches the board's hardware address.
- Check **Bin folder** contains the correct `.btl` file for your board model.
- Try re-running the ADwin development environment once to initialize the USB driver binding.

### Relay toggle works in software but relay box doesn't respond

- The 5V digital output lines need to reach the relay box driver circuit.
- Check the IDC cable from the ADwin digital I/O port to the relay box.
- Verify power to the relay box (+5V supply).
- Confirm the bit assignment matches the physical relay box wiring (may differ from defaults).

### ADC reads wrong voltage during loopback

- Verify the cable is BNC-to-BNC, not BNC-to-bare-wire.
- Confirm both ends are plugged in (DAC 1 out → ADC 1 in).
- Try a different cable (shield continuity matters at >1 kHz).
- Note that PC-timed loopback has timing jitter — the two traces may show slight phase offset that is normal and not a hardware problem.

### Sine loopback plot shows flat line on ADC

- Board not booted.
- ADC channel not connected.
- Amplitude too low (try 2.0 V).

### App freezes during sine loopback

The worker runs in a background thread so the GUI should never freeze. If it does, check whether the DLL call to `Set_DAC` or `Get_ADC` is blocking (rare — happens if the board is unresponsive). Click **Stop** to interrupt. If the thread doesn't stop within 3 seconds, close the app normally — the DAC will return to 0V on the next board boot.
