# Opening and building Paleomag on a Windows 10/11 machine

Four scripts live beside this file. Run them from the repository root.

| Script | What it does |
|---|---|
| `Test-LaunchReadiness.ps1` | One-shot preflight: dependencies, registry access, compiler registration |
| `Find-VB6Projects.ps1` | Finds every `.vbp` on the computer and caches the list |
| `Start-VB6.ps1` | Opens a project in the IDE **elevated**, which is what avoids the registry error |
| `Install-VB6Dependency.ps1` | Reads a component's embedded type library; verifies and registers it |
| `Build-VB6Project.ps1` | Compiles to an EXE, with machine-local reference fix-ups |

## "Error accessing the system registry"

### Why it happens

While a project loads, the VB6 IDE registers the project's own type library and
the licence keys for its licensed controls (MSCOMCTL, MSCOMM32, MSFLXGRD and
friends) into the **machine** COM hive:

```
HKLM\SOFTWARE\Classes
HKLM\SOFTWARE\Classes\Licenses
HKLM\SOFTWARE\WOW6432Node\Classes
```

On Windows Vista and later a standard user token cannot write there, even for
an account that belongs to Administrators, because UAC hands out a filtered
token. VB6 predates UAC, so instead of asking for elevation it reports
`Error accessing the system registry`.

Confirm it on any machine with:

```bash
powershell -ExecutionPolicy Bypass -File .\VB6\Start-VB6.ps1 -Diagnose
```

### The fix

Run the IDE elevated:

```bash
powershell -ExecutionPolicy Bypass -File .\VB6\Start-VB6.ps1
```

The script self-elevates through the normal UAC prompt and opens the project.
For a permanent one-click route, create a desktop shortcut that carries the
elevation flag:

```bash
powershell -ExecutionPolicy Bypass -File .\VB6\Start-VB6.ps1 -Shortcut
```

### What we deliberately did **not** do

The other common "fix" is to grant your user account Full Control over
`HKLM\SOFTWARE\Classes`. Do not do that. It gives every process running as you
permanent write access to the machine-wide COM registration hive, which means
any of them can redirect a CLSID for **every** account on the computer. That is
a lasting weakening of the machine to save a UAC click per IDE session.

## Finding projects

```bash
powershell -ExecutionPolicy Bypass -File .\VB6\Find-VB6Projects.ps1
```

Walks the repository, your profile, the usual lab folders, and every non-system
fixed drive, skipping Windows, package caches, and version-control internals.
The RAPID Paleomag project is ranked first. Results are cached at
`%LOCALAPPDATA%\RAPID\vb6-projects.json`, so `Start-VB6.ps1 -Project "<wildcard>"`
resolves a name without rescanning.

Use `-Deep` to include the system drive from its root, and `-Refresh` after
moving projects around.

## Installing a legacy component

Legacy dependencies that are not on the machine come from the licensed archive,
never from a DLL download site. The installer verifies before it touches
anything: the image must be 32-bit, and when you pass `-ExpectedTypeLibGuid`
that GUID must actually be inside the binary, so a same-named but different
component cannot be registered by mistake.

```bash
powershell -ExecutionPolicy Bypass -File .\VB6\Install-VB6Dependency.ps1 -Source "F:\Paleomag2013\vbSendMail\vbSendMail.dll" -TargetName "vbSendMail_v3.0.dll" -ExpectedTypeLibGuid "{332B82D3-3ED6-11D4-B1B5-00105AA5CCFF}"
```

Add `-WhatIfOnly` to verify without changing anything. An existing target file
is backed up before it is replaced.

## Building the EXE

```bash
powershell -ExecutionPolicy Bypass -File .\VB6\Build-VB6Project.ps1
```

Add `-PlanOnly` to see the reference resolution without compiling or elevating.

### The machine-local project copy

Component version numbers differ between machines. This computer, for example,
has `MSCOMCTL.OCX` registering type library **2.0**, while the committed project
asks for **2.2**.

Rather than edit the committed `.vbp` — which would break every other machine
that does have 2.2 — the build writes `<name>.localbuild.vbp` beside it with the
substitutions applied, compiles that, deletes it afterwards, and records every
substitution in `build\vb6\vb6-build-receipt.json`. The file is gitignored.

Use `-KeepLocalProject` to inspect what was generated, or `-NoFixups` to compile
the project exactly as committed.

## "MSCOMCTL.OCX could not be loaded"

This is a **design-time** error: it happens when the IDE opens the project, not
when the compiled EXE runs.

The committed project asks for `MSComctlLib` type library **2.2**:

```
Object={831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0; MSCOMCTL.OCX
```

If the control installed on the machine embeds type library **2.0**, the IDE
cannot bind it and reports that the control could not be loaded. A `2.2` key
may still exist under the type-library GUID as an empty stub with no `win32`
payload; that resolves to nothing and does not help.

### File version is not type-library version

These are independent numbers, and only the second one matters to a `.vbp`:

| Component | File version | Type library |
|---|---|---|
| `MSCOMCTL.OCX` (2004) | 6.01.9782 | 2.0 |
| `MSCOMCTL.OCX` (2005) | 6.01.9786 | 2.0 |
| `MSCOMCTL.OCX` post-MS12-027 | 6.1.98.x | **2.2** |
| `vbSendMail_v3.0.dll` | 3.06.0005 | 5.7 |

Swapping one 6.01.97xx build for another does not change anything: both embed
2.0. Only the 2012-and-later 6.1.98.x builds embed 2.2.

### Check a candidate before installing it

Read the type library straight out of any component, registering nothing:

```bash
powershell -ExecutionPolicy Bypass -File .\VB6\Install-VB6Dependency.ps1 -Source "<path to MSCOMCTL.OCX>" -WhatIfOnly
```

It prints the embedded GUID, name, and version. If it does not say
`MSComctlLib version 2.2`, that file will not satisfy the project reference,
whatever its file version says.

### Immediate: open a remapped copy

```bash
powershell -ExecutionPolicy Bypass -File .\VB6\Start-VB6.ps1 -UseLocalCopy
```

This generates `<name>.localbuild.vbp` with the reference remapped to the
version registered here and opens that instead. Edits you make land in the
temporary copy, not the repository, so treat it as read-only browsing or copy
your changes back deliberately.

`-Diagnose` lists every component whose requested version is unavailable,
without launching anything.

### Durable: install a 2.2 control

Install the current signed Microsoft `MSCOMCTL.OCX` (6.1.98.x). It registers
type library 2.2, so the committed project opens with no workaround, and it
also replaces a build that predates the MS12-027 fix.

```bash
powershell -ExecutionPolicy Bypass -File .\VB6\Install-VB6Dependency.ps1 -Source "<path to the newer MSCOMCTL.OCX>" -ExpectedTypeLibGuid "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}"
```

The installer backs up the existing file first and prints the type-library
versions that appear afterwards, so you can confirm 2.2 registered.

### Why the committed project still says 2.2

Changing it to 2.0 would fix this machine and break every machine that has the
patched control. The reference stays as committed, and the version difference
is absorbed per machine by the build copy.

### The compiled EXE is not affected

A compiled VB6 EXE binds controls by CLSID, not by type-library version.
`PALEOMAG2013.exe` embeds the MSCOMCTL ListView CLSID, which is registered here
by the installed control, so the EXE builds and runs regardless of the 2.0/2.2
mismatch. If you see this error from the EXE rather than the IDE, it is a
different problem - check that the control is registered at all.

## "No make available in the Working Model Edition"

This means the VB6 compiler is disabled because the **edition was never
registered**, not that anything is wrong with the project.

VB6 registers its edition when its setup program runs. If the `VB98` folder was
copied onto the machine instead, every binary runs — the IDE opens, the
toolchain (`LINK.EXE`, `C2.EXE`) is present — but VB6 falls back to its most
restricted mode and refuses to compile. That was the state of this computer
until 2026-09-03, when the English Visual Studio 6.0 Enterprise setup was run
and `ProductDir` appeared; the project compiled immediately afterwards.

What actually predicts `/make`:

| Marker | Reliable? |
|---|---|
| `HKLM\...\VisualStudio\6.0\Setup\Microsoft Visual Basic\ProductDir` | **Yes** |
| `LINK.EXE` and `C2.EXE` beside `VB6.EXE` | **Yes** |
| `HKCU\SOFTWARE\Microsoft\VisualStudio\6.0` | No — absent on a working install here |
| A "Visual Studio 6.0" uninstall entry | No — absent on a working install here |

`Test-LaunchReadiness.ps1` checks only the two reliable markers. If they are
missing, run the setup program from your licensed VB6 media, with two cautions:

1. **Match the language of the media to the installation you want.** Installing
   from media of a different language re-introduces that language's components
   and satellite DLLs across the shared `VB98`, `Common`, and `Designer` trees.
2. **Do not bypass the licence.** Fabricating edition or licence registry data
   is not a supported route and is not something this repository will script.

## Running the EXE for the first time

`Sub Main` in `modProg.bas` reads the settings-file path from
`GetSetting(App.EXEName, "Settings", "INIFile", ...)`. On first run that value
does not exist, so the program opens a file dialog asking for `Paleomag.ini`,
then remembers your choice in the registry under `PALEOMAG2013`.

Point it at the lab's real `Paleomag.ini`. For a no-communication smoke test,
copy `VB6\Defaults.ini` somewhere writable, rename it `Paleomag.ini`, and select
that — then confirm no-communication mode before anything is connected.
