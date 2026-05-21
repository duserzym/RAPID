# RAPID Compiled Apps Manual Index

This index tracks the compiled RapidPy operator apps, the current manual source, and whether a dedicated VB6 transition sheet section exists in that manual yet.

## Current Compiled-App Manual Status

| App | Build output | Manual source | Website app page | VB6 transition sheet |
| --- | --- | --- | --- | --- |
| Gaussmeter Control | `dist/RapidPy_Gaussmeter.exe` | `docs/gaussmeter-user-manual.md` | `docs/site/apps/index.html#gaussmeter-control` | Not yet |
| VRM Decay Logger | `dist/RapidPyVRM.exe` | `docs/vrm-logger-user-manual.md` | `docs/site/apps/index.html#vrm-decay-logger` | Not yet |
| ADwin Communications Tester | `dist/RapidPyADWin.exe` | `docs/adwin-comms-user-manual.md` | `docs/site/apps/index.html#adwin-comms` | Not yet |
| COM Port Mapper | `dist/RapidPyCOMMapper.exe` | README only | `docs/site/apps/index.html#com-port-mapper` | Not applicable yet |
| AF Tuner | `dist/RapidPyAFTuner.exe` | README today; long-form manual pending | `docs/site/apps/index.html#af-tuner` | Planned |
| XY Sample Changer | `dist/RapidPyChangerXY.exe` | `docs/changer-xy-control-user-manual.md` | `docs/site/apps/index.html#changer-xy-control` | Not yet |
| Up/Down Control | `dist/RapidPyUpDown.exe` | `docs/updown-control-user-manual.md` | `docs/site/apps/index.html#updown-control` | Yes |
| RAPID v4 Main Shell | `dist/RapidPySystemShell.exe` | README today; long-form manual pending | `docs/site/apps/index.html#rapid-v4-main-shell` | Planned |

## Transition-Sheet Scope

A VB6 transition sheet should help an experienced operator answer three questions quickly:

1. Which VB6 form owned this workflow?
2. What is the Python app or panel name now?
3. Which actions, settings files, and safety assumptions changed in the migration?

The Up/Down manual is the current reference implementation for that format.

## Build Layout Reminder

Current compiled apps are being normalized to a single repo layout:

- PyInstaller specs live in `installer/`
- temporary build products live in repo-root `build/`
- compiled one-file executables live in repo-root `dist/`
- screenshot and documentation helper scripts live in repo-root `tools/`
