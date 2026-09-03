# Opening and building Paleomag on a Windows 10/11 machine

Four scripts live beside this file. Run them from the repository root.

| Script | What it does |
|---|---|
| `Test-LaunchReadiness.ps1` | One-shot preflight: dependencies, registry access, compiler registration |
| `Find-VB6Projects.ps1` | Finds every `.vbp` on the computer and caches the list |
| `Start-VB6.ps1` | Opens a project in the IDE **elevated**, which is what avoids the registry error |
| `Install-VB6Dependency.ps1` | Verifies and registers a legacy 32-bit component from an archive |
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

## "No make available in the Working Model Edition"

This means the VB6 compiler is disabled because the **edition was never
registered**, not that anything is wrong with the project.

VB6 registers its edition when its setup program runs. If the `VB98` folder was
copied onto the machine instead, every binary runs — the IDE opens, the
toolchain (`LINK.EXE`, `C2.EXE`) is present — but VB6 falls back to its most
restricted mode and refuses to compile.

`Test-LaunchReadiness.ps1` reports the three observable symptoms:

- the setup key `HKLM\...\VisualStudio\6.0\Setup\Microsoft Visual Basic` has no
  `ProductDir`;
- `HKCU\SOFTWARE\Microsoft\VisualStudio\6.0` does not exist;
- there is no "Visual Basic 6" uninstall entry.

The fix is to run the setup program from your licensed VB6 media. Two cautions
before you do:

1. **Match the language of the media to the installation you want.** Installing
   from media of a different language re-introduces that language's components
   and satellite DLLs across the shared `VB98`, `Common`, and `Designer` trees.
2. **Do not bypass the licence.** Fabricating edition or licence registry data
   is not a supported route and is not something this repository will script.
