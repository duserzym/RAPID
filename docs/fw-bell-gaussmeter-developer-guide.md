# FW Bell Driver And SCPI Developer Guide

This guide is for developers who want to work with the FW Bell gaussmeter driver path outside the RAPID GUI.

It documents the driver stack, the helper executable, the DLL expectations, and the SCPI commands currently wired into RAPID.

## Preparation Toolkit In This Repo

The FW Bell workflow now assumes the preparation assets in `tools/` are part of the repository contract.

Primary files:

- [`../tools/usb5100_probe.c`](../tools/usb5100_probe.c)
- [`../tools/usb5100_probe.exe`](../tools/usb5100_probe.exe)
- [`../tools/zadig.exe`](../tools/zadig.exe)
- [`../tools/zadig.ini`](../tools/zadig.ini)
- [`../tools/zadig_install_fwbell.py`](../tools/zadig_install_fwbell.py)

Bundled libusb-win32 support payload:

- [`../tools/installer_x86.exe`](../tools/installer_x86.exe)
- [`../tools/installer_x64.exe`](../tools/installer_x64.exe)
- [`../tools/x86/libusb0_x86.dll`](../tools/x86/libusb0_x86.dll)
- [`../tools/x86/libusb0.sys`](../tools/x86/libusb0.sys)
- [`../tools/amd64/libusb0.dll`](../tools/amd64/libusb0.dll)
- [`../tools/amd64/libusb0.sys`](../tools/amd64/libusb0.sys)
- [`../tools/ia64/libusb0.dll`](../tools/ia64/libusb0.dll)
- [`../tools/ia64/libusb0.sys`](../tools/ia64/libusb0.sys)
- [`../tools/license/libusb0/installer_license.txt`](../tools/license/libusb0/installer_license.txt)

Reference-only INF investigation files:

- [`../tools/fw_bell_5100.inf`](../tools/fw_bell_5100.inf)
- [`../tools/fw_bell_5100_fixed.inf`](../tools/fw_bell_5100_fixed.inf)

If you need to prepare another machine, start from the committed `tools/` tree in this repo before hunting for copies in temp directories or older downloads.

## Architecture Summary

The RAPID FW Bell path is split into two layers:

1. Python layer in `RapidPy/rapidpy_common/gaussmeter.py`
2. Native x86 helper in `tools/usb5100_probe.c` and `tools/usb5100_probe.exe`

Python does not directly load `usb5100.dll`.

Instead:

1. Python locates a stable `usb5100.dll`.
2. Python runs `tools/usb5100_probe.exe`.
3. The helper loads `usb5100.dll` and calls vendor exports.
4. The helper prints parseable key-value output.
5. Python parses the helper output and presents a normal `GaussmeterClient` API.

This exists because the proven vendor path is the x86 sidecar path, not direct x64 ctypes loading.

## Files You Will Touch

- `RapidPy/rapidpy_common/gaussmeter.py`
- `RapidPy/gaussmeter_control/gaussmeter_control/app.py`
- `RapidPy/com_port_mapper/com_port_mapper/probe.py`
- `tools/usb5100_probe.c`
- `tools/usb5100_probe.exe`

Useful support files in the repo:

- `tools/zadig.exe`
- `tools/zadig.ini`
- `tools/zadig_install_fwbell.py`
- `tools/installer_x86.exe`
- `tools/installer_x64.exe`
- `tools/fw_bell_5100.inf`
- `tools/fw_bell_5100_fixed.inf`
- `tools/x86/libusb0_x86.dll`
- `tools/amd64/libusb0.dll`
- `tools/x86/libusb0.sys`
- `tools/amd64/libusb0.sys`
- `tools/ia64/libusb0.dll`
- `tools/ia64/libusb0.sys`
- `tools/license/libusb0/installer_license.txt`

## Windows Driver Requirements

The device must be visible as the FW Bell USB device, typically:

- `USB\VID_16A2&PID_5100`

The path that was actually validated earlier in development was a Zadig-installed `libusb-win32` binding.

Practical guidance:

1. Use [`../tools/zadig.exe`](../tools/zadig.exe).
2. Enable `List All Devices`.
3. Choose the FW Bell device.
4. Install `libusb-win32`.

Preparation workflow notes:

- [`../tools/zadig.ini`](../tools/zadig.ini) captures the defaults used during the validated setup.
- [`../tools/zadig_install_fwbell.py`](../tools/zadig_install_fwbell.py) is the optional automation entry point if you want to script Zadig rather than click through the UI.
- The bundled `installer_*.exe` and `libusb0*` payload files are committed so the driver setup path can be reproduced from the repo itself.

Do not assume the unsigned INF files in the repo are the primary supported path. They were part of the investigation, but signature enforcement blocked that route on this machine.

## Vendor DLL Requirements

You need both files in the same stable directory:

- `usb5100.dll`
- `libusb0.dll`

Do not build new code against a temp extraction path and assume that is the installation layout.

RAPID intentionally stopped treating `%TEMP%\rapid_fwbell_5100\...` as a normal default runtime location.

Supported ways for the Python layer to find the DLL now are:

- explicit GUI browse path
- `RAPID_USB5100_DLL`
- `RAPID_FW_BELL_DLL`
- stable installed path such as `C:\Program Files (x86)\FW Bell\PC5180\usb5100.dll`
- PATH lookup

## Native Exports Used

The helper currently uses these vendor exports:

- `openUSB5100`
- `closeUSB5100`
- `scpiCommand`

Those are resolved in `tools/usb5100_probe.c` with `GetProcAddress` after loading `usb5100.dll`.

## Helper CLI Contract

The helper supports these modes:

- `status`
- `read`
- `command <SCPI>`

Optional explicit DLL argument:

```text
usb5100_probe.exe --dll C:\stable\fwbell\usb5100.dll status
usb5100_probe.exe --dll C:\stable\fwbell\usb5100.dll read
usb5100_probe.exe --dll C:\stable\fwbell\usb5100.dll command *IDN?
```

If `--dll` is omitted, the helper now expects `usb5100.dll` to be discoverable through a stable normal path such as PATH. It no longer hardcodes a user temp extraction directory.

## Expected Helper Output

Successful `status` output looks like:

```text
status=ok
dll=C:\stable\fwbell\usb5100.dll
command=*IDN?
response=F.W.BELL MODEL 5180,R3.01
```

Successful `read` output looks like:

```text
status=ok
dll=C:\stable\fwbell\usb5100.dll
command=:MEASURE:FLUX?
response=-0.018T
```

The Python backend currently parses these keys:

- `status`
- `dll`
- `command`
- `response`
- `dll_dir`
- `loaded`

## Current SCPI Commands In Use

Validated and/or integrated commands:

- `*IDN?`
- `:MEASURE:FLUX?`
- `:SYSTEM:AZERO`
- `:UNIT:FLUX:DC:TESLA`
- `:UNIT:FLUX:DC:GAUSS`
- `:UNIT:FLUX:DC:AM`
- `:UNIT:FLUX:AC:TESLA`
- `:UNIT:FLUX:AC:GAUSS`
- `:UNIT:FLUX:AC:AM`
- `:SENSE:FLUX:RANGE 0`
- `:SENSE:FLUX:RANGE 1`
- `:SENSE:FLUX:RANGE 2`
- `:SENSE:FLUX:RANGE:AUTO`

Behavior currently mapped in Python:

- RAPID mode index `0` and `1` normalize to FW Bell DC behavior
- RAPID mode index `2`, `3`, and `4` normalize to FW Bell AC behavior only where supported
- unsupported legacy-only behaviors raise explicit errors

Unsupported on the FW Bell backend today:

- instrument time read/write
- legacy null reset semantics
- legacy peak tracking reset semantics

## Python Entry Points

The main high-level API is in `RapidPy/rapidpy_common/gaussmeter.py`.

Important functions and classes:

- `gaussmeter_driver_status()`
- `fwbell_driver_status()`
- `find_usb5100_dll()`
- `find_usb5100_helper()`
- `GaussmeterClient`
- `_FwBellGaussmeterBackend`

USB auto mode uses the FW Bell backend when the FW Bell path is available.

## Reproducing The Driver Check Without The GUI

Use the helper first.

Example:

```powershell
e:\Github\RAPID\tools\usb5100_probe.exe --dll C:\stable\fwbell\usb5100.dll status
```

If that works, test raw reading:

```powershell
e:\Github\RAPID\tools\usb5100_probe.exe --dll C:\stable\fwbell\usb5100.dll read
```

Then test arbitrary SCPI:

```powershell
e:\Github\RAPID\tools\usb5100_probe.exe --dll C:\stable\fwbell\usb5100.dll command *IDN?
e:\Github\RAPID\tools\usb5100_probe.exe --dll C:\stable\fwbell\usb5100.dll command :MEASURE:FLUX?
```

After that, validate the Python layer.

Example snippet shape:

```python
from rapidpy_common.gaussmeter import gaussmeter_driver_status, GaussmeterClient

print(gaussmeter_driver_status([r"C:\stable\fwbell\usb5100.dll"]))

client = GaussmeterClient(dll_search_paths=[r"C:\stable\fwbell\usb5100.dll"])
client.connect()
print(client.read())
client.disconnect()
```

## Compiling The Helper

`tools/usb5100_probe.exe` should be built as x86.

Typical Windows/MSVC flow:

1. Open an x86 Visual Studio developer command prompt.
2. Build `tools/usb5100_probe.c` with MSVC.

Example command once the x86 toolchain environment is active:

```cmd
cl /nologo /W4 /EHsc /Fe:e:\Github\RAPID\tools\usb5100_probe.exe e:\Github\RAPID\tools\usb5100_probe.c
```

Why x86 matters:

- the validated vendor runtime path is 32-bit
- the helper isolates that from the main Python process

## Common Failure Modes

### `usb5100.dll was not found`

Meaning:

- the helper cannot find the vendor DLL

Fix:

- pass `--dll`
- set `RAPID_USB5100_DLL`
- add the DLL directory to PATH

### `LoadLibrary usb5100.dll error=126`

Meaning:

- Windows found the target DLL path but could not load its dependencies

Fix:

- make sure `libusb0.dll` is next to `usb5100.dll`
- verify you are not mixing incompatible copies

### `openUSB5100 failed`

Meaning:

- the DLL loaded, but the device could not be opened

Typical causes:

- driver binding is wrong
- device is not started by Windows
- another process already owns the device
- the device is unplugged or not powered correctly

### The device appears twice in PnP and both are unhealthy

That is a Windows/device-state problem, not a Python parsing problem. Fix the driver/device state first, then retry the helper.

## Recommended Development Order

When working on FW Bell support, use this order:

1. Prepare the machine from the committed `tools/` payload first.
2. Make the device appear correctly in Windows.
3. Make `usb5100_probe.exe --dll ... status` work.
4. Make `read` work.
5. Validate raw SCPI commands.
6. Only then change Python backend behavior.
7. Only then change the GUI.

That order avoids confusing GUI bugs with driver-layer failures.

For machine preparation, the expected sequence is:

1. Use [`../tools/zadig.exe`](../tools/zadig.exe) or [`../tools/zadig_install_fwbell.py`](../tools/zadig_install_fwbell.py) to bind the device.
2. Keep the repo's libusb-win32 payload available if the install tool needs to reference its bundled files.
3. Put the vendor `usb5100.dll` and `libusb0.dll` in a stable directory.
4. Validate the repo helper [`../tools/usb5100_probe.exe`](../tools/usb5100_probe.exe) before touching Python code.

## Current RAPID Behavior Worth Knowing

As of the latest changes in this repo:

- the GUI status is mode-aware
- USB mode reports FW Bell status directly
- the backend no longer auto-selects a transient temp-extracted sample DLL as the default runtime
- the helper no longer hardcodes a user temp DLL path when `--dll` is omitted

Those changes were made specifically to stop the GUI from reporting misleading FW Bell availability based on a stale recovery path.