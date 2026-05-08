# RapidPy COM Port Mapper

Lightweight operator utility for finding RAPID serial devices on Windows before launching a subsystem app.

## What It Does

- Enumerates visible COM ports with Windows adapter metadata
- Highlights ports that look like enhanced or PCI serial adapters
- Safely probes known RAPID protocols:
  - SQUID magnetometer at `1200,N,8,1`
  - Quicksilver motor controllers at `57600,N,8,2`
  - 908A gaussmeter through the legacy `gm0.dll` driver when that DLL is installed
- Shows legacy VB6 default role hints:
  - `COM3` Vacuum
  - `COM4` Up/Down motor
  - `COM5` Turning motor
  - `COM6` X/Changer motor
  - `COM7` Y motor
  - `COM8` Susceptibility
  - `COM9` AF
  - `COM10` SQUID

Ports without a positive protocol match are still listed with their Windows-friendly names and hardware IDs so the operator can narrow down the remaining candidates.

## Run

```bash
cd RapidPy/com_port_mapper
python -m pip install -r requirements.txt
python main.py
```

## Probe Scope

The scanner currently makes high-confidence identifications for:

- SQUID magnetometer serial lines
- Quicksilver motor controller serial lines
- 908A gaussmeter ports, but only when `gm0.dll` is present on the machine or exposed through `RAPID_GM0_DLL`

Vacuum, susceptibility, and AF ports are shown with adapter metadata and VB6 legacy hints, but they are not force-probed with control commands because the current converted codebase does not expose a read-only identity command for those devices.