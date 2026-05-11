# RapidPy ADwin Communication Tester — User Manual & Test Procedure

> **Purpose:** Verify ADwin board communication, relay wiring, and DAC/ADC signal paths **before** connecting amplifiers or coils. All testing is done at low voltage with a patch cable — no high-voltage hardware required.

---

## Table of Contents

1. [What This App Does (and Doesn't Do)](#1-what-this-app-does-and-doesnt-do)
2. [System Requirements](#2-system-requirements)
3. [Safety Rules](#3-safety-rules)
4. [Launching the App](#4-launching-the-app)
5. [Application Layout](#5-application-layout)
6. [Left Panel — Control Cards (Detailed)](#6-left-panel--control-cards-detailed)
7. [Right Panel](#7-right-panel)
8. [What Is PAR/FPAR? (Why Is It Not Shown?)](#8-what-is-parfpar-why-is-it-not-shown)
9. [ADwin Driver & Board Setup](#9-adwin-driver--board-setup)
10. [Standard Testing Procedure](#10-standard-testing-procedure)
11. [Interpreting the Console](#11-interpreting-the-console)
12. [How RapidPy Talks to ADwin (Technical Background)](#12-how-rapidpy-talks-to-adwin-technical-background)
13. [Building the Executable from Source](#13-building-the-executable-from-source)
14. [Troubleshooting](#14-troubleshooting)

---

## 1. What This App Does (and Does Not Do)

### What it does

The **ADwin Communication Tester** is a low-risk hardware verification tool. It tests every major communication channel between the PC and the ADwin board:

| Channel | How tested |
|---|---|
| USB / Ethernet link to board | Boot Board + Test_Version check |
| DAC output (analog voltage out) | Write a voltage, measure it |
| ADC input (analog voltage in) | Read back a voltage from a DAC->ADC patch cable |
| Digital outputs (relay control bits) | Toggle each bit, read back the register |
| Integer process parameters (PAR) | Write/read round-trip (used by self-test internally) |
| Float process parameters (FPAR) | Write/read round-trip (used by self-test internally) |
| ADbasic process management | Load, start, stop .abp files |

### What it does NOT do

- It does **not** drive AF coils or a power amplifier.
- It does **not** perform actual alternating field demagnetization.
- It does **not** set paleomagnetics measurement parameters.
- It is **not** a replacement for the full `af_tuner` module.

After passing all tests in this app, the AF coil tuner and full RAPID orchestration system can be used safely.

---

## 2. System Requirements

| Requirement | Detail |
|---|---|
| Operating system | Windows 10 / 11 |
| ADwin hardware | ADwin Gold II (or Pro II / T9 series) via USB-B or Ethernet |
| ADwin driver | `adwin64.dll` (64-bit Python) or `adwin32.dll` (32-bit) in `C:\Windows\` |
| Python (dev) | Python 3.10+, conda environment `paleomag` |
| Test cable | Short BNC cable, DAC output -> ADC input (for loopback tests) |

The app launches without the driver installed. All hardware buttons are disabled and the console explains what to install. No data is lost.

---

## 3. Safety Rules

> **Never connect the power amplifier or AF coils while using this app.**

1. The sine loopback test outputs real analog voltages (up to +-10 V) on the selected DAC channel.
2. Only a short BNC patch cable should be connected between the DAC output and ADC input pins during loopback tests.
3. The relay test outputs 5 V logic signals to relay control lines. It is safe to toggle relays freely.
4. Always click **All OFF** at the end of a relay test session.
5. The board must be **booted** (firmware loaded) before any hardware operation. Skipping this step is not possible — the Boot button must succeed before other controls become active.

---

## 4. Launching the App

**From the built executable:**

```
Double-click:  E:\Github\RAPID\dist\RapidPyADWin.exe
```

**From the conda development environment:**

```powershell
conda activate paleomag
Set-Location E:\Github\RAPID\RapidPy\adwin_comms
python main.py
```

Settings are saved to `~/.rapidpy_adwin_comms.json` and restored automatically on next launch. Window geometry (position and size) is also saved.

---

## 5. Application Layout

The window is a resizable horizontal splitter with two sides:

```
+---------------------------------------+------------------------------------------+
|  LEFT: Control Cards (scrollable)     |  RIGHT: Monitoring Panels                |
|                                       |                                          |
|  +------------------+---------------+ |  +--------------------------------------+ |
|  | Board Config     | Relay Test    | |  |  Loopback Signal Plot (live)         | |
|  |                  | (Dig. Out.)   | |  |  green = DAC out, gold = ADC in      | |
|  +------------------+---------------+ |  +--------------------------------------+ |
|  | Direct DAC/ADC   | Process Ctrl  | |  |  Hardware Self-Test Results          | |
|  +------------------+---------------+ |  |  (per-step PASS/FAIL list)           | |
|  | Sine Loopback Test (spans both)   | |  +--------------------------------------+ |
|  +------------------------------------+ |  |  Console (timestamped log)           | |
+---------------------------------------+------------------------------------------+
```

- Drag the **horizontal splitter** handle (vertical bar between the two sides) to re-size.
- Drag the **vertical splitter** handles on the right side to trade space among the three right panels.
- The left panel scrolls vertically if the window is too short to show all cards.

---

## 6. Left Panel — Control Cards (Detailed)

### Board Configuration

This is the primary card. It must be configured first — no hardware operation is possible until the board is successfully booted.

#### Board #

ADwin device number. Default is `1`. If multiple ADwin boards are installed on the same PC, each has a unique device number assigned in the ADwin hardware configuration utility. For RAPID, it is always `1`.

#### Boot file

The firmware file (`*.btl`) to load onto the board at startup.

| Board model | Correct BTL file |
|---|---|
| ADwin Gold II / T9 | `ADwin9.btl` |
| ADwin Pro II | `ADwin9.btl` |
| ADwin 5 (older) | `ADwin5.btl` |

The dropdown shows the most common options.

#### Bin folder

The directory containing the `.btl` firmware file and optionally `.abp` process files. **Leave blank** — the app automatically finds the right folder from the Windows registry (see Section 9). A grey hint text shows the auto-detected path. Only fill this field if auto-detection fails. Use **Browse...** to navigate to the folder manually.

#### Boot Board

Loads the real-time operating system firmware onto the ADwin board. This is required once per session (or after the board loses power). The operation calls `ADboot()` in the DLL, which takes 2-5 seconds. The status label shows the result:

| Status label | Meaning |
|---|---|
| `● Not booted` (grey) | No boot attempt yet |
| `● Booted ✓` (green) | Firmware loaded successfully |
| `● Boot failed` (red) | ADboot() returned an error — see console |

**What "booting" means:** The ADwin board is an embedded real-time computer. Its firmware (the real-time OS) is not stored permanently — it must be uploaded from the PC each time the board powers on. The `.btl` file contains this firmware image. Once booted, the board runs its OS and waits for ADbasic process programs to be loaded.

#### Run Self-Test

Starts the automated 8-step hardware verification sequence. Requires the board to be booted and a DAC->ADC patch cable connected. Results stream into the Self-Test Results panel on the right.

---

### Relay Test — Digital Outputs

The ADwin board has 6 digital output bits (bits 0-5) that drive the relay box in the RAPID instrument rack.

#### Default relay assignments

| Bit | Default label | Physical function |
|---|---|---|
| Bit 0 | Axial Relay | Routes power amp output to the axial AF Helmholtz coil |
| Bit 1 | Trans Relay | Routes power amp output to the transverse AF coil |
| Bit 2 | IRM/AF Select | Switches between IRM capacitor discharge and AF sine wave path |
| Bit 3 | ARM Polarity | Controls DC bias field polarity for ARM acquisition |
| Bit 4 | Bit 4 | Spare — lab-specific use |
| Bit 5 | Bit 5 | Spare — lab-specific use |

#### Toggle buttons

Each button is a toggle. When pressed (checked), that bit is driven HIGH (5 V), energizing the relay. The full 6-bit output word is written to the board immediately on each toggle via `Set_Digout()`.

#### Digout display

Shows the current output word in hexadecimal: `Digout: 0x00` means all bits off. `Digout: 0x01` means only bit 0 is ON (Axial Relay energized).

Do not set bits 0 and 1 simultaneously (0x03) with the amplifier connected — that would route the amplifier to both coils at once.

#### Read State

Reads the current digout register from the board via `Get_Digout()` and syncs all toggle buttons to match. Use this to verify the actual hardware state.

#### All OFF

Sets the entire output word to `0x00` in a single call. Always click this at the end of a relay test session.

---

### Direct DAC / ADC

Allows direct voltage reads and writes to individual DAC and ADC channels without needing a loaded ADbasic process. Requires only a booted board.

#### DAC (Digital-to-Analog Converter)

| Control | Description |
|---|---|
| **DAC Ch** | DAC output channel (1-8). Channel 1 is the primary AF output in RAPID. |
| **V** | Voltage to output (+-10.000 V, 3 decimal places). Converted to a 16-bit unsigned count: `count = int(V * 32768 / 10) + 32768` — the same formula as VB6 RAPID. |
| **Write** | Calls `Set_DAC(channel, count, device)`. The DAC holds this voltage until another write or re-boot. |

After testing, always write `0.0 V` to return the DAC output to zero.

#### ADC (Analog-to-Digital Converter)

| Control | Description |
|---|---|
| **ADC Ch** | ADC input channel (1-8). Channel 1 is the primary AF feedback input in RAPID. |
| **Read** | Calls `Get_ADC(channel, device)`, converts the 16-bit count to voltage: `V = (count - 32768) * 10 / 32768`, and displays the result. |

**Loopback test:** Connect a short BNC from DAC Ch 1 output to ADC Ch 1 input. Write any voltage and immediately read it back. A healthy board reads within +-0.05 V of the commanded value.

---

### Process Control

Manages ADbasic real-time process programs (`.abp` files) on the board. The board supports up to 10 concurrent processes in numbered slots.

| Control | Description |
|---|---|
| **.abp** | Path to the compiled ADbasic process file. Browse with **Browse...** |
| **Process #** | Which process slot to target (1-10). AF ramp process uses slot 1 by default. |
| **Load** | Transfers the `.abp` file to the board via `ADBload()`. Process is ready but not running. |
| **Start** | Starts the process via `ADB_Start()`. The ADbasic code runs on the board's real-time CPU at the configured rate (typically 10 kHz for the AF process). |
| **Stop** | Stops the process via `ADB_Stop()`. Leaves it loaded but halted. |
| **Clear All** | Stops and unloads all 10 process slots. Use before loading a new process. |

Once started, an ADbasic process runs entirely on the board's dedicated real-time CPU, independent of the PC. The PC can monitor process state via PAR/FPAR registers but has no timing influence.

---

### Sine Loopback Test

Generates a sine wave entirely in Python (PC-timed), writing each sample to the DAC and reading it back from the ADC. Results are plotted in real time. Tests the full DAC->ADC communication path.

> **Important:** This is PC-side software timing, not ADwin real-time timing. Timing jitter of 1-5 ms is normal (Windows scheduler). The actual sample rate may differ from IO Rate. This is acceptable for communication testing — real AF measurements use the ADwin real-time process which runs at precisely 10 kHz on the board.

| Control | Description |
|---|---|
| **Frequency** | Target sine frequency (0.1-10000 Hz). Use 10-100 Hz for initial tests. |
| **Amplitude** | Peak voltage (0.001-9.999 V). Use 1.0 V for initial tests. |
| **Duration** | How long to run (0.1-3600 s). 5 seconds is sufficient for a check. |
| **IO Rate** | Target sample rate (1-2000 Hz). 500 Hz is a good default. |
| **DAC Ch / ADC Ch** | Channel selection. Must match the patch cable connections. |
| **Run Sine Loopback** | Starts the background worker thread. UI stays responsive. |
| **Stop** | Signals the worker to stop early. DAC returns to 0 V. |
| **Status label** | Shows: `Idle`, `Running... N Hz, N V, N s`, `Done.`, or `Error.` |

---

## 7. Right Panel

### Loopback Signal Plot

A live `pyqtgraph` plot updated every 50 ms during a sine loopback test.

| Trace | Color | Content |
|---|---|---|
| DAC out | Bright green | Commanded voltage sent to the DAC |
| ADC in | Gold | Voltage measured back from the ADC |

**Navigation:** Right-click + drag to pan; scroll wheel to zoom; **Clear Plot** to reset data.

**Reading the plot:**

- Traces overlapping: DAC->ADC path working correctly
- ADC trace flat / near 0: Cable not connected or wrong channel
- ADC trace attenuated: Signal through resistive elements (check connections)
- ADC trace noisy: Check cable shielding and grounding

### Hardware Self-Test Results

Shows per-step results from the automated self-test (green = PASS, red = FAIL).

The summary banner shows the final verdict:
- `All N tests PASSED` (green) — board communication verified
- `K of N tests FAILED` (red) — check connections and driver

### Console

Scrolling timestamped log of every operation. Format: `[HH:MM:SS] message`.

**Clear Console** erases the log without affecting board state.

---

## 8. What Is PAR/FPAR? (Why Is It Not Shown?)

### What PAR and FPAR are

PAR (integer parameters, indices 1-80) and FPAR (float parameters, indices 1-80) are communication registers that serve as shared memory between the PC and a running ADbasic process:

- The PC writes values via `Set_ADBPar()` / `Set_ADBFPar()`
- The running ADbasic process reads these values to get its operating parameters
- The process can write back into PAR/FPAR to report status to the PC

### How they are used in RAPID

The AF ramp ADbasic process reads these registers at startup to configure the ramp:

| Register | Content |
|---|---|
| PAR31 | Ramp mode (1=Normal AF, 2=Debug, 3=Clipping test, 4=AF Tuning sweep) |
| PAR32 | DAC output channel (PORT_SINEOUT) |
| PAR33 | ADC input channel (PORT_ACCUR — field monitor input) |
| PAR34 | Process timing constant (computed from IO rate) |
| PAR35 | Noise level (16-bit count distance from target peak to trigger ramp-down) |
| PAR36 | Number of sine periods to hold at peak field |
| PAR37 | Number of periods for ramp down |
| PAR38 | Ramp-down mode (0 = slope V/s, 1 = period count) |
| FPAR31 | Ramp-up slope (V/s) |
| FPAR32 | Ramp-down slope (V/s) |
| FPAR33 | Target peak field (V, from field monitor calibration) |
| FPAR34 | Sine frequency (Hz) |
| FPAR35 | AC amplitude limit (V) |
| FPAR36 | Maximum ramp voltage absolute limit (V, default 10) |
| FPAR37 | Maximum peak monitor voltage absolute limit (V, default 10) |

### Why PAR/FPAR is hidden in this app

These registers are **set automatically** by the `af_tuner` module based on measurement protocol parameters (field, frequency, coil calibration). Direct editing by hand is:

- **Confusing** — register indices and units are not obvious without reading the ADbasic source
- **Error-prone** — writing a wrong value can corrupt a running process
- **Unnecessary** — the af_tuner handles this completely automatically

The PAR/FPAR functionality is still fully implemented in the code and is used internally by the self-test routine (which tests PAR[79] and FPAR[79] — safely above the range used by any production process, which tops out at index 38).

To re-enable the PAR/FPAR card for debugging, add `ctrl_grid.addWidget(self._build_par_card(), row, col)` to `_build_ui()` in `app.py`.

---

## 9. ADwin Driver & Board Setup

### Driver Installation

ADwin communicates via `adwin64.dll` (64-bit Python) or `adwin32.dll` (32-bit). These are proprietary DLLs from Jager Messtechnik — they cannot be replaced by a PyPI package.

**Installation steps:**

1. Download the ADwin software package from [www.adwin.de](https://www.adwin.de) -> Support -> Downloads.
2. Select the package matching your board model (Gold II, Pro II, T9).
3. Run the installer as Administrator.
4. Accept USB driver installation when prompted.
5. Note the install directory — the `.btl` boot file is inside it.

After installation: `adwin64.dll` is in `C:\Windows\` and `ADwin9.btl` is in `C:\ADwin\BTL\`.

**Verify the DLL:**
```powershell
conda activate paleomag
python -c "import ctypes; ctypes.WinDLL('C:/Windows/adwin64.dll'); print('OK')"
```

### BTL Boot File Auto-Detection

When Bin folder is blank, the app searches automatically:

1. Windows registry: `HKLM\SOFTWARE\Jager Mestechnik GmbH\ADwin\Directory` -> BTL value
2. Common paths: `C:\ADwin\BTL\`, `C:\ADwin9\BTL\`, `C:\ADwin\`
3. Current working directory (last resort)

**Verify the registry path:**
```powershell
Get-ItemProperty -Path "HKLM:\SOFTWARE\Jäger Meßtechnik GmbH\ADwin\Directory" -EA SilentlyContinue
```

### USB / Ethernet Connection

**USB:** Connect the board with a USB-B cable before powering it on. After driver install, it appears in Device Manager as "ADwin Gold II" under Universal Serial Bus devices. A yellow warning icon means the driver needs repair — re-run the ADwin installer.

**Ethernet:** Use the ADwin Ethernet configuration utility to assign a static IP. The DLL handles Ethernet transparently.

**Verify connection without RAPID:**
```powershell
conda activate paleomag
python -c "
import ctypes
dll = ctypes.WinDLL('C:/Windows/adwin64.dll')
dll.ADTest_Version.restype = ctypes.c_int
dll.ADTest_Version.argtypes = [ctypes.c_int]
ver = dll.ADTest_Version(1)
print(f'Test_Version = {ver}  (' + ('board live' if ver else 'not booted') + ')')
"
```

---

## 10. Standard Testing Procedure

Perform all steps in order. Do not proceed to the next step if the current step fails.

**Equipment needed:**
- ADwin Gold II board, powered on, USB connected
- ADwin software (driver) installed on the PC
- One short BNC-to-BNC patch cable (for DAC->ADC loopback tests)
- Access to the RAPID relay box (for relay verification, Step 4)

---

### Pre-Test Checklist

Before starting:

- [ ] ADwin board powered on (front panel LED on)
- [ ] USB cable connected from board to PC
- [ ] Board visible in Device Manager without yellow warning icon
- [ ] `C:\Windows\adwin64.dll` exists
- [ ] ADwin BTL file present (auto-detected or known path)
- [ ] `RapidPyADWin.exe` launches without immediate crash

---

### Step 1 — Driver & DLL Verification

**Objective:** Confirm the ADwin DLL loads correctly at startup.

1. Launch `RapidPyADWin.exe`.
2. Observe the Console panel within 2 seconds of launch.

**PASS:** `[HH:MM:SS] adwin64.dll loaded successfully.`

**FAIL:** `[HH:MM:SS] [WARNING] adwin64.dll not found: [error]`
`  -> Hardware buttons disabled. Install ADwin driver to enable.`

If FAIL: Install the ADwin driver package and relaunch.

---

### Step 2 — Board Boot

**Objective:** Load firmware onto the board and confirm communication.

1. Confirm Board # = `1`, Boot file = `ADwin9.btl`, Bin folder = blank (auto-detect).
2. Click **Boot Board**.
3. Wait 2-10 seconds.

**PASS:** Status label shows `● Booted ✓` (green). Console shows: `Board booted successfully.`

**FAIL:** Status label shows `● Boot failed` (red). Possible console messages:
- `BTL boot file not found` -> Browse to the correct folder
- `ADWIN boot failed (return code -1)` -> Board not connected or powered off

---

### Step 3 — Automated Self-Test

**Objective:** Run all 8 hardware verification tests with a single button press.

**Prerequisite:** Step 2 complete + BNC patch cable connected DAC Ch 1 -> ADC Ch 1.

1. Connect the BNC patch cable: DAC Ch 1 output -> ADC Ch 1 input on the ADwin panel.
2. Confirm DAC Ch = `1` and ADC Ch = `1` in the Direct DAC/ADC card.
3. Click **▶ Run Self-Test** in the Board Configuration card.
4. Watch the Hardware Self-Test Results panel.

**The 8 tests:**

| # | Test | Pass criterion |
|---|---|---|
| 1 | DLL / driver | DLL loaded at startup |
| 2 | Board Test_Version | ADTest_Version() returns non-zero |
| 3 | DAC->ADC loopback +5V | Read within +-0.5 V of +5 V |
| 4 | DAC->ADC loopback -5V | Read within +-0.5 V of -5 V |
| 5 | DAC->ADC loopback 0V | Read within +-0.3 V of 0 V |
| 6 | Digout set/get | Each read-back matches the written pattern |
| 7 | PAR[79] write/read | Read = written value exactly |
| 8 | FPAR[79] write/read | Read within 0.001 of written value |

**PASS:** `All 8 tests PASSED — board communication verified.` (green banner)

**FAIL:** `K of 8 tests FAILED — check connections and driver.` (red banner)

If loopback tests fail: verify the BNC cable connects the correct DAC and ADC pins. Try a different cable.

---

### Step 4 — Relay Box Verification

**Objective:** Confirm each relay bit drives the correct physical relay.

**Prerequisite:** Steps 1-2 complete. Physical access to the relay box.

1. Click **All OFF**. Confirm all relay box LEDs are off and `Digout: 0x00` is shown.

2. For each bit 0 through 5:
   - Toggle the button ON
   - Confirm the Digout label updates
   - Walk to the relay box and confirm the correct LED turns on
   - Toggle the button OFF
   - Confirm the LED goes off

3. At the end: Click **All OFF**. Confirm all LEDs off, `Digout: 0x00`.

4. Document any mismatches between expected and observed relay behavior.

**PASS:** Every bit toggles the correct physical relay with no cross-triggering.

---

### Step 5 — DAC / ADC Spot Check

**Objective:** Verify voltage accuracy across the full DAC/ADC range.

**Prerequisite:** Steps 1-3 complete. BNC cable DAC Ch 1 -> ADC Ch 1.

Set DAC Ch = `1`, ADC Ch = `1`. For each commanded voltage: click Write, then click Read.

| Commanded (V) | Acceptable read result |
|---|---|
| +5.000 | +4.900 to +5.100 V |
| -5.000 | -5.100 to -4.900 V |
| +2.500 | +2.450 to +2.550 V |
| -2.500 | -2.550 to -2.450 V |
| 0.000 | -0.030 to +0.030 V |
| +9.000 | +8.900 to +9.100 V |
| -9.000 | -9.100 to -8.900 V |

After testing, write 0.000 V.

**PASS:** All measured values within the acceptable range.

Deviations > +-0.15 V indicate the board may need factory calibration.

---

### Step 6 — Sine Loopback Plot

**Objective:** Verify the DAC->ADC path handles a continuous waveform.

**Prerequisite:** Steps 1-5 complete. BNC cable DAC Ch 1 -> ADC Ch 1.

1. Set: Frequency = `100 Hz`, Amplitude = `1.000 V`, Duration = `5.0 s`, IO Rate = `500 Hz`, DAC Ch = `1`, ADC Ch = `1`.
2. Click **▶ Run Sine Loopback**.
3. Observe the Loopback Signal Plot.

**PASS criteria:**
- Both green (DAC) and gold (ADC) traces are visible and oscillating
- ADC peak amplitude within 5% of DAC amplitude (0.95 to 1.05 V peak)
- No severe noise or clipping

**FAIL indicators:**
- ADC trace flat -> cable or channel problem
- ADC severely attenuated -> cable resistance or ADC gain issue
- ADC heavily noisy -> grounding/shielding issue

4. Repeat with Frequency = `1000 Hz`. At 1000 Hz with IO Rate 500 Hz, PC-timed undersampling is expected — this is acceptable for communication testing.

5. Click **Clear Plot** after each test.

---

### Step 7 — Process Load / Start / Stop

**Objective:** Verify ADbasic process management works end-to-end.

**Prerequisite:** Steps 1-2 complete. A compiled `.abp` file available.

1. In Process Control: click Browse and select an `.abp` file. Set Process # = `1`.
2. Click **Load** -> console: `Process loaded: [path]`
3. Click **Start** -> console: `Process #1 started.`
4. Wait 2 seconds.
5. Click **Stop** -> console: `Process #1 stopped.`
6. Click **Clear All** -> console: `All processes cleared.`

**PASS:** All 4 operations complete without `[ERROR]` in the console.

---

### Pass / Fail Criteria Summary

| Test | Required for AF demagnetization |
|---|---|
| Step 1 — DLL load | Must PASS |
| Step 2 — Board boot | Must PASS |
| Step 3 — Automated self-test (all 8) | Must PASS |
| Step 4 — Relay verification | Must PASS |
| Step 5 — DAC/ADC spot check | Must PASS (within +-0.15 V) |
| Step 6 — Sine loopback plot | Must PASS (< 5% amplitude error) |
| Step 7 — Process management | Recommended |

---

## 11. Interpreting the Console

| Message pattern | Meaning |
|---|---|
| `adwin64.dll loaded successfully.` | Driver installed and DLL accessible |
| `[WARNING] adwin64.dll not found:` | Install ADwin driver |
| `Board booted successfully.` | Firmware loaded, board responding |
| `[ERROR] BTL boot file not found:` | BTL path is wrong — use Browse |
| `[ERROR] ADWIN boot failed (return code -1)` | Board not connected or powered off |
| `DAC ch1 -> +5.000 V` | Successful DAC write |
| `ADC ch1 -> +4.983 V` | Successful ADC read |
| `Bit 0 (Axial Relay) -> ON  (word=0x01)` | Relay bit toggled |
| `Process loaded: C:\...` | Process file sent to board |
| `  Testing: [name]...` | Self-test step in progress |
| `[SELFTEST] Complete: 8/8 passed.` | All self-tests passed |
| `[ERROR] ...` | Any operation failure — read the full message |

---

## 12. How RapidPy Talks to ADwin (Technical Background)

### Communication stack

```
Python (RapidPy adwin_af.py)
    |  ctypes.WinDLL("adwin64.dll")
adwin64.dll  (Jager Messtechnik Windows DLL)
    |  USB driver / TCP socket
ADwin Gold II board
    |  ADwin real-time OS
ADbasic process (.abp)
    |  DAC / ADC / Digital I/O hardware
Coils, sensors, relays
```

### Key DLL functions

| Function | Purpose |
|---|---|
| `ADTest_Version(device)` | Returns 0 if not booted; non-zero version code if booted |
| `ADboot(btl_path, device, 0, 1)` | Loads firmware; returns > 0 on success |
| `Set_DAC(channel, count, device)` | Sets a DAC output (16-bit count) |
| `Get_ADC(channel, device)` | Reads an ADC input; returns 16-bit count |
| `Set_Digout(word, device)` | Sets all 6 relay control bits at once |
| `Get_Digout(device)` | Reads the current digital output register |
| `ADBload(path, device)` | Loads a compiled ADbasic process |
| `ADB_Start(process_num, device)` | Starts a loaded process |
| `ADB_Stop(process_num, device)` | Stops a running process |
| `Set_ADBPar(index, value, device)` | Writes to an integer parameter register |
| `Get_ADBPar(index, device)` | Reads from an integer parameter register |
| `Set_ADBFPar(index, value, device)` | Writes to a float parameter register |
| `Get_ADBFPar(index, device)` | Reads from a float parameter register |

### Voltage conversion

```python
# PC -> DAC: volts to 16-bit unsigned count
count = int(volts * 32768 / 10) + 32768   # clipped to [0, 65535]

# ADC -> PC: 16-bit unsigned count to volts
volts = (count - 32768) * 10 / 32768
```

This is identical to the VB6 RAPID code conversion formula.

### Matching VB6 RAPID

The legacy VB6 code declares the same DLL functions with `Declare Function`. RapidPy's `adwin_af.py` uses ctypes with the exact same function names and argument types, ensuring full behavioral compatibility with the VB6 reference implementation.

---

## 13. Building the Executable from Source

```powershell
conda activate paleomag
Set-Location "E:\Github\RAPID\RapidPy\adwin_comms"
& "C:\Users\[username]\anaconda3\envs\paleomag\Scripts\pyinstaller.exe" `
    --noconfirm "..\..\installer\rapid_adwin_comms.spec"
Copy-Item ".\dist\RapidPyADWin.exe" "..\..\dist\RapidPyADWin.exe" -Force
```

The spec file is at `E:\Github\RAPID\installer\rapid_adwin_comms.spec`.

**Note:** `adwin64.dll` is NOT bundled in the EXE — it must be installed on the target PC via the ADwin driver package. The EXE is approximately 64 MB (one-file PyInstaller bundle). First launch takes 2-5 seconds to extract.

---

## 14. Troubleshooting

### DLL not found on launch

Console shows `[WARNING] adwin64.dll not found`. All hardware buttons are disabled.

**Fix:** Install the ADwin software package from [www.adwin.de](https://www.adwin.de). After installation, relaunch the app.

---

### Boot fails — BTL file not found

Console shows `[ERROR] BTL boot file not found: C:\ADwin\BTL\ADwin9.btl`.

**Fix:**
1. Click Browse next to Bin folder and navigate to the folder containing `ADwin9.btl`.
2. Verify: `Test-Path "C:\ADwin\BTL\ADwin9.btl"` in PowerShell.
3. If not found: ADwin software is not installed or was installed to a non-standard path.

---

### Boot fails — return code -1

Console shows `[ERROR] ADWIN boot failed (return code -1)`.

**Causes:** Board not powered on; USB cable disconnected; wrong device number; board driver not installed (yellow icon in Device Manager).

---

### ADC reads wrong value during loopback

Self-test tests 3-5 fail; ADC reads more than 0.5 V from commanded value.

**Causes:** Patch cable not connected; connected to wrong channel; wrong channel numbers in DAC Ch / ADC Ch; poor BNC connector (try a different cable); board needs factory calibration.

---

### Relay LED does not respond

Toggle buttons work (Digout label updates) but relay box LED does not change.

**Causes:** Relay control cable between ADwin and relay box is disconnected; wrong bit assigned to the relay; relay box fuse blown.

---

### App window is blank or crashes on open

**Causes:**
- `pyqtgraph` or `PySide6` issue — run `python main.py` from terminal and read the traceback.
- Corrupted settings file — delete `~/.rapidpy_adwin_comms.json` and relaunch.
- GPU driver issue with pyqtgraph — set environment variable `QT_OPENGL=software`.

---

*Document last updated: May 2026 — aligned with GUI compact-layout update (2-column control cards, self-test panel, PAR/FPAR hidden from UI).*
