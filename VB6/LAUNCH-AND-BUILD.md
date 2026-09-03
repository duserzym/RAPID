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

Component version numbers differ between machines. A machine whose
`MSCOMCTL.OCX` registers type library 2.0, for instance, cannot bind the 2.2
reference the committed project carries.

Rather than edit the committed `.vbp` — which would break every other machine
that does have 2.2 — the build writes `<name>.localbuild.vbp` beside it with the
substitutions applied, compiles that, deletes it afterwards, and records every
substitution in `build\vb6\vb6-build-receipt.json`. The file is gitignored.

Use `-KeepLocalProject` to inspect what was generated, or `-NoFixups` to compile
the project exactly as committed.

On this computer no substitution is needed any more: since MSCOMCTL 6.01.9846
was installed the plan is clean and `-NoFixups` builds succeed. The machinery
stays for machines that still carry an older control.

## "MSCOMCTL.OCX could not be loaded"

**Resolved on this computer, 2026-09-03.** `MSCOMCTL.OCX` 6.01.9846 is
installed and registers type library 2.2, so the committed project opens and
builds with no substitution. What follows is the reasoning, for the next
machine.

This is a **design-time** error: it happens when the IDE opens the project, not
when the compiled EXE runs. The committed project asks for `MSComctlLib` type
library **2.2**:

```
Object={831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0; MSCOMCTL.OCX
```

If the installed control embeds an older type library the IDE cannot bind it. A
`2.2` key may exist under the type-library GUID as an empty stub with no
`win32` payload; that resolves to nothing and does not help.

### File version is not type-library version

These are independent numbers, and only the second one matters to a `.vbp`.
The type-library version moved three times across otherwise similar-looking
builds:

| Component | File version | Type library | Source |
|---|---|---|---|
| `MSCOMCTL.OCX` | 6.00.8177 | 2.0 | Original VB6 media |
| `MSCOMCTL.OCX` | 6.01.9782 | 2.0 | 2004 |
| `MSCOMCTL.OCX` | 6.01.9786 | 2.0 | 2005 |
| `MSCOMCTL.OCX` | 6.01.9834 | **2.1** | KB2708437 (MS12-027) |
| `MSCOMCTL.OCX` | **6.01.9846** | **2.2** | KB3096896 (MS16-004) |
| `vbSendMail_v3.0.dll` | 3.06.0005 | 5.7 | legacy archive |

Swapping one 6.01.97xx build for another changes nothing. Even the MS12-027
build only reaches 2.1. **Only KB3096896 provides 2.2.**

### Check any candidate before installing it

Read the type library straight out of a component, registering nothing:

```bash
powershell -ExecutionPolicy Bypass -File .\VB6\Install-VB6Dependency.ps1 -Source "<path to MSCOMCTL.OCX>" -WhatIfOnly
```

If it does not say `MSComctlLib version 2.2`, that file will not satisfy the
reference, whatever its file version says.

### Getting a 2.2 control

`VB60SP6-KB3096896-x86-ENU.msi` from the Microsoft Download Center
([details page](https://www.microsoft.com/en-us/download/details.aspx?id=50722)).
Extract it without installing, so you can inspect and choose what to apply:

```bash
msiexec /a "VB60SP6-KB3096896-x86-ENU.msi" /qn TARGETDIR="C:\temp\kb3096896"
```

Then install **only** `SYSTEM\mscomctl.OCX`:

```bash
powershell -ExecutionPolicy Bypass -File .\VB6\Install-VB6Dependency.ps1 -Source "C:\temp\kb3096896\SYSTEM\mscomctl.OCX" -TargetName "MSCOMCTL.OCX" -ExpectedTypeLibGuid "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}"
```

**Do not install the `comctl32.ocx` from the same package.** The project asks
for `ComctlLib` **1.3**, and the rollups ship 1.4 (KB2708437) and 1.5
(KB3096896). Installing it trades one mismatch for another. Only apply the one
control the project actually needs.

The installer backs up the file it replaces, deletes the stale `.oca`
type-information cache, and prints the type-library versions that appear
afterwards.

Note on verifying the download: the SHA-256 published in the 2016 KB article
does not match the file Microsoft serves today, because the package was
re-signed and re-released in November 2020. Check the Authenticode signature
instead - it is timestamped, valid, and chains to Microsoft Root Certificate
Authority 2011.

### If you cannot get a 2.2 control

```bash
powershell -ExecutionPolicy Bypass -File .\VB6\Start-VB6.ps1 -UseLocalCopy
```

That opens a copy with the reference remapped to whatever is registered.
Builds already handle this automatically. Edits in the IDE land in the
temporary copy, not the repository.

### Why the committed project says 2.2

2.2 is the version the project was authored against, and the version this
computer now has. Downgrading it would break every machine with a patched
control, and VB6 will not bind a lower minor version than the one requested -
which is exactly why 2.0 and 2.1 both failed here.

### The compiled EXE was never affected

A compiled VB6 EXE binds controls by CLSID, not type-library version.
`PALEOMAG2013.exe` built and ran through all of this.

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
