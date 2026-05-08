# FW Bell 5100 / 5180 User Guide

This guide is for operators who want to use the RAPID Python gaussmeter tools with the FW Bell USB gaussmeter path that replaced the old assumptions about a COM-only Hirst workflow.

The current RAPID code supports two gaussmeter backends:

- Legacy `gm0.dll` for the older Hirst-style path.
- FW Bell USB support through `usb5100.dll` plus an x86 helper executable at `tools/usb5100_probe.exe`.

For the attached FW Bell instrument, use the FW Bell path.

## What You Need

You need all of the following before the GUI can talk to the instrument:

- The physical gaussmeter powered on and connected over USB.
- A working Windows USB driver binding for the device.
- The vendor DLL pair:
  - `usb5100.dll`
  - `libusb0.dll`
- The RAPID helper executable:
  - `tools/usb5100_probe.exe`
- A Python environment that can run the RAPID GUI.

Files already in this repo that are relevant:

- `RapidPy/gaussmeter_control/main.py`
- `RapidPy/com_port_mapper/main.py`
- `RapidPy/rapidpy_common/gaussmeter.py`
- `tools/usb5100_probe.c`
- `tools/usb5100_probe.exe`
- `tools/zadig.exe`
- `tools/fw_bell_5100.inf`
- `tools/fw_bell_5100_fixed.inf`
- `tools/installer_x86.exe`
- `tools/installer_x64.exe`

Important note about the DLLs:

- This repo does not currently include a checked-in `usb5100.dll` copy.
- The GUI now intentionally refuses to auto-use transient temp-extraction locations as if they were a normal install.
- You must either install the vendor DLLs to a stable location or browse directly to a stable `usb5100.dll` file in the GUI.

## Known Good Architecture

The working RAPID design for FW Bell is:

1. Python GUI runs as normal 64-bit Python.
2. Python does not load `usb5100.dll` directly.
3. Python shells out to `tools/usb5100_probe.exe`.
4. `tools/usb5100_probe.exe` is built as x86 and loads `usb5100.dll`.
5. `usb5100.dll` talks to the device through the libusb driver stack.

This is intentional. Do not try to replace it with direct 64-bit ctypes loading unless you are prepared to revalidate the whole hardware stack.

## Step 1: Identify the Device in Windows

With the gaussmeter connected and powered on, Windows should show a USB device with hardware ID similar to:

- `USB\VID_16A2&PID_5100`

The exact product name may show as `FW Bell 5100` even when the instrument answers `F.W.BELL MODEL 5180,R3.01` over SCPI.

If the device is missing entirely:

- check the cable
- check power to the gaussmeter
- try a different USB port
- verify Windows plays the connect sound or Device Manager refreshes

## Step 2: Install the USB Driver Binding

The path that worked in practice on this machine was Zadig plus `libusb-win32`, not the custom unsigned INF route.

Recommended procedure:

1. Close any RAPID gaussmeter GUI, helper, or vendor utility that might already have the device open.
2. Run `tools/zadig.exe` as Administrator.
3. In Zadig, enable `Options -> List All Devices`.
4. Select the FW Bell device matching `VID_16A2` / `PID_5100`.
5. Choose `libusb-win32` as the target driver.
6. Install or replace the driver.

What not to rely on as the primary path:

- `tools/fw_bell_5100.inf`
- `tools/fw_bell_5100_fixed.inf`

Those INF-based attempts ran into Windows signature enforcement during the earlier recovery/debugging work.

After Zadig finishes:

1. Unplug the gaussmeter.
2. Wait a few seconds.
3. Plug it back in.
4. Open Device Manager and confirm it lands under a libusb-related device class.

If Windows shows the device with status `Unknown`, `Code 10`, or does not start the device, the RAPID GUI will not be able to open it even if the DLLs are present.

## Step 3: Put the Vendor DLLs in a Stable Location

RAPID needs a stable `usb5100.dll` location. Do not depend on a temp extraction directory.

You need these two files together in the same directory:

- `usb5100.dll`
- `libusb0.dll`

Acceptable stable locations include:

- a vendor install directory such as `C:\Program Files (x86)\FW Bell\PC5180\`
- a manually created stable folder outside temp
- a lab tools folder that is not deleted on reboot

Avoid using:

- `%TEMP%\rapid_fwbell_5100\...`

That location is useful for recovery/debugging, but not as a reliable default installation target.

## Step 4: Tell RAPID Where `usb5100.dll` Lives

You have three supported ways to do this.

### Option A: Use the GUI Browse Button

This is the simplest operator path.

1. Launch the gaussmeter GUI.
2. Click `Browse DLL`.
3. Select the stable `usb5100.dll` file.
4. Leave the mode on `USB / driver auto`.
5. Read the driver status banner.

The GUI will pass that DLL path into the FW Bell backend.

### Option B: Set an Environment Variable

Set one of these before starting the GUI:

- `RAPID_USB5100_DLL`
- `RAPID_FW_BELL_DLL`

Example PowerShell for the current session:

```powershell
$env:RAPID_USB5100_DLL = 'C:\stable\fwbell\usb5100.dll'
python e:\Github\RAPID\RapidPy\gaussmeter_control\main.py
```

### Option C: Put the DLL on PATH

This is supported, but less explicit than Option A or B. Prefer A or B for day-to-day lab use.

## Step 5: Launch the RAPID Gaussmeter GUI

From the repo root, run:

```powershell
python e:\Github\RAPID\RapidPy\gaussmeter_control\main.py
```

If you want to use the COM port mapper too:

```powershell
python e:\Github\RAPID\RapidPy\com_port_mapper\main.py
```

## Step 6: Read the Driver Status Banner

The GUI now reports status according to the selected connection mode.

When `USB / driver auto` is selected, the banner is reporting FW Bell status only.

Examples:

- `Driver ready: C:\stable\fwbell\usb5100.dll [F.W.BELL MODEL 5180,R3.01]`
- `Driver unavailable: usb5100.dll not found. Set RAPID_USB5100_DLL, install the FW Bell runtime, or browse to the vendor DLL.`
- `Driver unavailable: C:\stable\fwbell\usb5100.dll: openUSB5100 failed`

Interpretation:

- `usb5100.dll not found` means RAPID cannot find the vendor DLL path.
- `openUSB5100 failed` means the DLL loaded, but the USB device/driver stack is not actually opening.
- A banner with `F.W.BELL MODEL 5180,R3.01` means the helper opened the device and SCPI is working.

## Step 7: Connect and Use the Instrument

Once the driver banner says `Driver ready`:

1. Leave `USB / driver auto` selected.
2. Click `Connect`.
3. Watch the live reading pane update.
4. Use the controls for:
   - measurement mode
   - units
   - range
   - auto range
   - auto zero

Backend notes:

- `set_mode`, `set_units`, and `set_range` are supported on the FW Bell path.
- `Auto zero` is supported.
- instrument time features are not supported on the FW Bell path
- peak-reset and some legacy-only time/null behaviors are not available on the FW Bell path

If you try an unsupported operation, RAPID should show a clear backend-specific error instead of silently failing.

## Troubleshooting

### Problem: The GUI says `usb5100.dll not found`

Fix:

1. Verify you actually have `usb5100.dll`.
2. Verify `libusb0.dll` is in the same folder.
3. Browse directly to the DLL in the GUI, or set `RAPID_USB5100_DLL`.

### Problem: The GUI says `openUSB5100 failed`

This means the DLL was found and loaded, but the device did not open.

Fix sequence:

1. Unplug and replug the gaussmeter.
2. Confirm the device is present in Device Manager.
3. Recheck the driver binding in Zadig.
4. Make sure no other program already has the device open.
5. Retry with the helper directly as described below.

### Problem: The helper works in one session and fails after reboot

Most likely causes:

- the DLL path was temp-based and disappeared or moved
- the USB binding changed
- Windows attached the device to a different driver instance

Use a stable DLL location and recheck Zadig binding.

### Problem: Manual COM mode does not help

That is expected for the FW Bell USB path. The FW Bell backend does not use manual COM port selection.

## Quick Validation Procedure

After setup, a fast validation loop is:

1. Launch the GUI.
2. Browse to the stable `usb5100.dll`.
3. Confirm the banner changes to `Driver ready`.
4. Click `Connect`.
5. Confirm the live value updates.
6. Change units and range and confirm the display responds.

If step 3 fails, use the developer guide to test the helper directly.

## Current Behavior on This Machine

At the time this guide was written, the RAPID GUI had been updated so that:

- it no longer auto-picks a transient `%TEMP%` sample DLL as the default FW Bell runtime
- in USB mode it reports FW Bell status directly
- when no stable DLL is configured, it reports that clearly

That means the GUI now distinguishes these cases cleanly:

- missing DLL configuration
- present DLL but failed USB open
- successful device response

The remaining operational requirement is still the Windows driver/device state. Software alone cannot recover a device that Windows has not started correctly.