# RapidPy main application: portable pilot release

The glass workspace includes Dashboard, Sample Queue, Sequence, Live Measure,
Settings, and Calibration Center. Its Windows bundle also includes AF Tuner,
Data Viewer, Gaussmeter, Up/Down, VRM, and Webcam helpers. No separate Python
installation or source checkout is required. Vendor drivers, station settings,
and approved calibration remain necessary for hardware operation.

## Launch and verify

Copy the **entire** `dist/RapidPyMain` folder, including `_internal`, to the
operator computer. From a trusted copy of the release, run the bundled
verifier before launching `RapidPyMain.exe`.

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\Verify-RapidPy.ps1 -RequireFullInventory
```

The verifier checks every packaged file against `build_manifest.json`. It
checks startup and the No-Comm UI with a temporary user profile, temporary
safety journal, and Windows-only PATH. It prints the location of the JSON
report, which records the manifest hash and the startup and UI results.
Use `-EvidenceDirectory` to save the report under a specific directory
outside the package.

The individual diagnostic entry points are also available.

```powershell
.\RapidPyMainConsole.exe --check-startup
.\RapidPyMainConsole.exe --smoke-test
.\RapidPyMain.exe --tool data_viewer
```

The smoke test uses temporary configuration and INI preferences, forces
No-Communication, constructs all six main panels, imports all six helper entry
points, decodes a generated video clip in Webcam Monitor, verifies a simulated
VRM recording, and prints JSON.
It uses an isolated configuration and leaves instrument connections untouched.
Startup diagnostics are read-only. Configuration normally
lives in `~/.rapid/config.json`; `RAPID_CONFIG` selects a different file.
Windows preferences and queue recovery state belong to the operator account.
Main-window helper launches use separate bundled processes; VRM retains its
ownership leases and handoff environment.

In No-Communication mode, **Diagnostics → VRM Data Collection** opens the
simulated SQUID. Select **Take Baseline**, **Start Logging**, and **Stop** to
practice the recording workflow. The logger writes each session's CSV and
manifest in its own `SIMULATED` subfolder next to the selected output path.

**View → Webcam Monitor** lists USB cameras by their operating-system device
identity and opens supported RTSP or HTTP(S) streams through Qt Multimedia.
The preview and snapshot button become active on receipt of a usable frame.
**Snapshot** saves a full-resolution PNG to the selected folder. Use
**Disconnect** to release the camera when you finish monitoring.

## Operator workflow

1. Open Settings. Select No-Communication for practice and set the data folder.
   For supervised hardware testing, import the station INI and review ports,
   geometry, motion positions, enable flags, and calibration.
2. Load a sample index or enter specimens in Sample Queue. Review specimen
   names, changer positions, and orientations.
3. Build/load a sequence, inspect its preview, save it, and select a specimen
   or queue. Live treatment plans are validated before ordinary preflight.
4. Measure the holder before specimens. Rejected magnetic blocks or failed
   bridge reads cannot replace the accepted holder correction.
5. Watch Live Measure/Step Monitor. Use Pause/Halt and inspect the safe-return
   result before touching the instrument.
6. Review results in Plots/Data Review. Simulation output stays in `SIMULATED`.
   Use **Open File** in **Data Review** to open a CIT specimen directly or to
   browse the specimens listed in a `.sam` index.
   In **Zijderveld (2D)**, choose **N–S projection** for east to the
   right or **E–W projection** for north to the right. Positive inclination
   is plotted below zero in both views.
   Use **View → Step labels** to show endpoints, label every measured
   treatment step, or hide labels on the directional plots.
   On narrow displays the controls wrap and the plots stack vertically. Scroll
   down to reach the measurement table and specimen list.
   Use **View → Recent steps** to review the latest measurements in the
   directional plots, intensity decay plot, measurement table, and fit range.
   Enter `0` to show all steps, the default. With Auto-watch on, both edges
   of the fit range follow the displayed window unless you select an interior
   step. Susceptibility and Arai plots show their full recorded data.
   Fit angles are independent of the moment scale. A free fit reports the
   direction of the remanence removed between the selected endpoints. An
   origin-anchored fit points from the origin toward the selected readings.
   On the stereonet, enable **CSD contours** to show the angular scatter
   around each reading. Solid arcs are in the lower hemisphere. Dashed arcs
   are in the upper hemisphere. New MagIC measurements carry recorded CSD
   in `dir_csd`.
   Data Review fills blank CSD values from matched RMG readings in verified
   RAPID bundles and archives.
   Run bundles retain scientific outputs, provenance, workflow summaries,
   communication evidence, and a hashed artifact index.
   For specimens with orientation headers, use **Core**, **Geographic**, or
   **Bedding** to review the corresponding direction. RAPID uses the `.sam`
   magnetic declination when supplied. New CIT rows use moment per volume
   (emu/cm³), while the live moment and MagIC use the measured moment.
   Data Review displays the magnitude stored in a CIT file as recorded.
   MagIC magnitudes are displayed in A·m².
   RMG files include treatment levels in gauss or °C, susceptibility, moment
   components, available uncertainty, and core direction. Sample height is
   included when a queue measured it and the station has an imported `UpDownMotor1cm`.
   In **Data Review**, choose **Susceptibility** under **Plot** to review RMG
   readings in measurement order, by recorded AF level, or by temperature.
   This view is available when the selected specimen has susceptibility readings.
   If publication is interrupted, reopening the original specimen for a run
   checks the journal before acquisition and restores the prior accepted files.
   Unexpected changes leave the journal available for review. A completed
   scientific publication can show diagnostics as pending if RAPID exits while
   completing the workflow summary or transcript.

7. After a native interruption, open **Diagnostics > Recover Interrupted Queue**.
   RAPID can retry an interrupted field cutoff, stop and check axes in place for
   pending motion, measurement, treatment or vacuum stages, or publish a shutdown
   whose transport closures were already verified. Field cutoff and motor stop
   leave the original queue latched until specimen support and vacuum are settled.
8. To abort the interrupted queue and shut down the station, sign in and open
   **Diagnostics > Shut Down Interrupted Queue…**. This is available after the
   original worker has exited or after restarting RAPID. Confirm only after all
   specimens are secured independently of vacuum, the control rod is empty,
   and the work area is clear. RAPID requires fresh field cutoff, stopped axes,
   lift clearance, positive valve and pump OFF replies, and transport closure
   before recording the queue as aborted. With the original app session still
   open, RAPID also closes the original SQUID and susceptibility connections.
   A failed check leaves the queue pending. Use **Retry Queue Shutdown** when
   RAPID had already recorded the original terminal connection closures as the
   remaining operation.
   Review interrupted rows after settling the station. Re-run starts a new
   acquisition, while Skip and Abort exclude the row.
9. To finish a single-step VB6 Up/Down pair from a saved Up acquisition, set the
   file's queue orientation to **Down** and enable **Both sides**. When starting the
   queue, select the original VB6 `.UP` for that file and step. Confirm that it
   covers the selected specimens and uses the same SQUID axis convention and gains.
   Enter the historical range factor, which must match the accepted native setting.
   For `SUSC`, transcribe each specimen's original standardized susceptibility from
   the Up pass record. RAPID measures Down without reapplying the treatment and
   links the original bytes and operator confirmation to the combined result.

### Continue a saved manual measurement

In Live Measure, choose **Save and Stop After Step** to complete the current
step and return the instrument to its safe state. When RAPID reports that the
run is saved, it has retained the original measurements and evidence. Any
previously accepted publication stays available in Data Review.
In live mode RAPID records fresh stopped-axis readbacks before confirming the save.

After relaunching, open the same specimen and sequence with the original
station settings and holder correction. Choose **Continue Saved Run** and
confirm the specimen identity, orientation, and treatment history. RAPID
verifies the saved evidence and starts with the next untouched step. It
publishes the full sequence after the remaining steps finish and the safe
return succeeds. If a later treatment may already have started, inspect the
specimen and station before deciding whether to start a fresh run.

### Review an unexpectedly interrupted manual run

After an unexpected exit, select the original specimen in Live Measure and
open **Review Interrupted Run**. The review shows which steps have captured
readings and which had an intent recorded without a complete acquisition.
Establish the state of the instrument and the treated specimen through the
station recovery procedure. Then record the disposition under your operator
name. This preserves the original treatment evidence and any unpublished
output, clears the interrupted run, and permits an intentional new run.
The review does not issue instrument commands or accept partial measurements.

### Chain-drive queues

For a chain-drive queue, import the original profile with
`XYTable.UseXYTableAPS=False`, the chain motor's port and address, slot range,
empty-hole interval, counts per slot, and the accepted lift and turning
calibration. Start the queue in Sample Queue. After confirming the empty rod,
enter the number physically marked on the chain slot under it. RAPID records
this observation before connecting, then checks lift clearance and relabels the
stopped chain controller. The chain has no independent absolute-slot sensor
in this workflow. The readback checks use the origin you entered.

RAPID uses the nearest calibrated empty slot for each blank or specimen and
the original specimen slot for drop-off. The chain parks at its origin for
initial loading and tray flips. The optional return commands park at the
first specimen slot in the selected order. After an interruption, use the
same recovery and shutdown commands described above.

In **Diagnostics > DC Motors**, imported stations show their participating
axes and fixed wiring. To use calibrated slots on an XY station, confirm the
empty rod and use **Home XY to Center**. On a chain station, select the marked
slot currently under the empty rod and use **Set Chain Reference**. Repeat
the reference after reconnecting. Changer movement verifies lift clearance.
Pickup and drop-off use the current verified specimen slot, with drop-off
height calculated for the imported station mode. This panel moves motors.
The native queue manages vacuum grip and specimen release.

**Run Queue** checks specimen files and output paths in the background.
The Preflight status stays on while those checks run. **Halt** cancels the
pending start. If you edit the queue, sequence, or settings during preflight,
start the updated queue again. Native station confirmation follows a
successful preflight.

Calibration Center records proposals, approval, expiry/invalidation, and
rollback. Review/approve calibration before using its coefficients. Help →
“Where did this VB6 control go?” supplies the searchable transition map.

**Diagnostics → Calibrate Rod** opens the Bartington susceptibility procedure.
Enter the certified standard identity, height, moment susceptibility in CGS,
scan spacing, and fit threshold. The lift advances at half the entered spacing.
Enable the Bartington bridge in Settings for live mode. The procedure uses the
accepted motor and bridge configuration and asks the signed-in
operator to confirm the mounted standard and clear opening. The scan saves its
position and bridge readings, independent center check, and proposed coil position
and moment factor. Use **Halt** or **Stop Scan** to cancel. If motor settlement
requires recovery, open DC Motors and use **Recover Verify Stopped**. A No-Comm run
creates simulated evidence for practice. Registry approval stores a provenance
record. Review the proposal and apply any accepted station settings separately.

For the operator-mounted magnetic standard, select **SQUID Rod Scan** in the
same Procedure menu. Use the accepted lift and SQUID configuration with the
1× range. Enter the standard height, fine spacing, fit threshold, and maximum
baseline drift in emu. A coarse pass locates the strongest axis response, and
a fine pass measures its center inside the accepted travel. The scan repeats
the center and baseline readings and records a measurement and zero position
proposal. Review the evidence and independently verify fit and positioning
before changing station settings.

For either rod calibration, **Standard tray slot** selects automatic loading.
Choose the slot containing the standard and confirm the empty control rod and
the calibrated empty scan slot. A chain station also needs the current physical
chain reference. RAPID uses the native queue station to load the standard,
check its measured height against the certified height, and return it to its
original slot. A disagreement in height returns the standard before sensor
acquisition. Approval requires evidence of the original-slot return and
verified station shutdown. The default **Mounted by operator** uses the manual
mount workflow. No-Comm records a selected tray slot in the simulated scan.

If a tray scan is interrupted with the station still owned, use **Shutdown
Interrupted Queue** in Diagnostics. Use **Retry Queue Shutdown** when only
original connection closure remains. The original journal retains the standard
identity, slot, and last verified transfer state.

**Diagnostics → U-channel Scan** handles mounted long cores and EPR tubes.
Import the station's `MinUpDownPos`, `FloorPos`, and `UpDownMotor1cm`. Use the
accepted SQUID calibration at 1× and measure the extension from the carriage
reference to the lowest point of the full mounted assembly. Enter the
bench-verified start, end, and zero counts for a long core. For an EPR tube,
enter the tube length, sample bounds measured from its lower end, and the two
buffers. RAPID calculates the scan span and places the zero 5 cm above the
first scan position. The plan must fit both the carriage and floor envelopes.

Choose one pass at 0° or four passes at 0°, 90°, 180°, and 270°. Live turns
require a verified top lift and the accepted 0° reference. RAPID takes two
bracketing zero readings for each group of positions and checks their drift.
Each output bundle contains `scan.json` with the plan and readbacks and
`readings.csv` with VB6-style metadata and SQUID units. A completed scan also
contains `profile.csv` with baseline-corrected instrument-axis moments in emu.
The specimen orientation is retained in the manifest. No-Comm writes the same
bundle with the simulation flag. **Halt** or **Stop Scan** cancels acquisition.
The scan shows a count of saved readings. If the process exits before the scan
finishes, its `.uchannel-pending-*` folder keeps the plan and `readings.jsonl`.
Only a completed scan publishes a corrected profile to its final bundle.

**Diagnostics → Transverse Probe Scan** uses a manually mounted probe at the
configured AF lift height. Verify the turning encoder reference and select a
clear start/end interval containing the current encoder angle. Verify degausser
cooling and confirm the probe setup before starting. Set the angular
pitch, a calibrated transverse field in mT, and the 908A COM port. Close the
separate 908A Gaussmeter window before scanning. The scan checks stationary
lift and turning readbacks, then measures a signed DC peak at each angle with
a calibrated transverse ADwin ramp. The table shows each verified reading as
it arrives, with the requested angle, actual encoder angle and signed peak in
mT. **Pause** waits for the next boundary between readings and verifies AF
outputs off and both axes stopped. **Resume** rechecks the probe position
before continuing. The scan keeps the station reserved while paused.

The scan leaves the probe at the last measured angle. Use **Move to Best
Angle** to turn separately to the measured angle with the largest absolute
peak. The app verifies that the encoder and original scan evidence still
match before moving. **Stop** checks both motor axes and the AF outputs.
If either verification fails, use **Verify Safe State** in the same tool with
the original station wiring. Recovery stops and verifies the equipment in
place. A live scan flushes each verified reading to a JSONL checkpoint in the
hardware diagnostics folder before it starts the next angle. Stop and
recovery records reference the checkpoint. Use complete JSONL lines to
inspect interrupted readings. Start a new scan to obtain a best angle after
an interruption. No-Comm exercises pause, resume, and move with simulated
readings.

## Externally treated specimens

Record and load a **Thermal Routine Planning** plan in Calibration Center.
In live mode, select the named specimen and use **Start / Resume** in Live
Measure. Select the completed thermal step and enter the furnace-log
reference, recorded peak, dwell, handling temperature, and completion time.
Confirm the cooled specimen is mounted for SQUID measurement according to the
station procedure.

Measure after each furnace treatment in plan order. RAPID appends one
SQUID measurement to the specimen's verified bundle. The requested label
identifies the step, and RMG and MagIC retain the observed temperature.
The bundle provenance links each operator attestation to that step's
acquisition record. Duplicate, out-of-order, and unverifiable existing
bundles are rejected before SQUID preflight. The normal **No-Comm**
thermal plan run remains a simulation.

## Review an earlier publication

When a successful run replaces or extends a specimen's saved results, RAPID
first saves the earlier publication under `.rapidpy-history` in that specimen's
output folder. The current `provenance.json` gives the location of the prior
`archive_manifest.json` in `previous_publication.relative_path`. The neighboring
`bundle.zip` contains the earlier legacy and MagIC files, run diagnostics, and
saved acquisition evidence listed in its artifact index.

In **Data Review**, use **Open MagIC Directory** and select the folder that
contains `archive_manifest.json`. Data Review checks the ZIP members against the
recorded SHA-256 hashes and opens the earlier MagIC results without extraction.
Archives from No-Comm runs stay inside the `SIMULATED` output.

## Back up completed measurements

Set **Settings → General → Backup folder** to an existing local or network
directory outside the specimen output. After an accepted measurement and its
quicklook are saved, RAPID verifies the published files against the artifact
index and copies them to `RapidPy backups/Measurements` in that folder.
Practice runs use `RapidPy backups/SIMULATED`. The copy runs in the background.
**Live Measure** reports its status and the archive path. Use **Open MagIC
Directory** in Data Review to open a backup archive.

The primary measurement remains available if the copy fails. Pending copies
survive an app restart. Reconnect the original backup folder, select that folder
in Settings, and click **Retry Pending Backups** in Live Measure. If a newer run
has replaced the publication, RAPID uses the verified specimen history to copy
the original run.

To back up the current publication manually, use the portable console and point
the first argument to the specimen output folder containing `artifact_index.json`.

```powershell
.\RapidPyMainConsole.exe --backup-measurement "D:\RAPID_data\Specimen-A" "E:\Lab backups"
```

## Register sample indexes

In **Sample Queue → Load Index**, use **Load File** to open a SAM or CSV index.
Choose its treatment, measurement direction, and acquisition blocks per
treatment, then select **Save to Registry**. **Registered indexes** reopens the
source file with those defaults. Save again to update them or select **Remove**
to remove the registry entry. A blank treatment uses the current Sequence.

Select a specimen and enter its changer position to add it to the queue. Paired
acquisition requires one treatment. If the queue already contains specimens
from the same index with different settings, align those rows before adding
another specimen.

## Rebuild and test

Use an environment with the declared RapidPy dependencies, setuptools, wheel,
and PyInstaller. The default build interpreter is the repository `.venv`:

```powershell
.\installer\build_rapid_main.ps1
.\.venv\Scripts\python.exe tools\test_rapid_main.py
.\.venv\Scripts\python.exe tools\test_rapid_main.py --suite data_viewer
```

The build packages a newly staged installation wheel, including metadata,
icons, common assets, helpers, and GPL licensing. Tests use isolated
configuration and Qt preferences. The manually triggered **Qualify the
portable Windows bundle** workflow runs the RAPID Main tests in separate
processes, with Sequence layout checks on a 1920 × 1080 virtual display. It also
runs the Data Review suite, then verifies the transferred package on a separate
Windows runner. The qualification artifact contains the report for that commit.
Physical station, vendor driver, and display checks use the lab's acceptance procedure.

### Measure workstation performance

Run the verifier with `-MeasurePerformance` on the oldest supported control
computer and retain its `verification.json` with the build's station evidence.
The separate Windows verification job runs the same workload on the hosted
runner.

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\Verify-RapidPy.ps1 `
    -RequireFullInventory -MeasurePerformance -PerformanceSpecimens 40 -PerformanceIdleSeconds 5
```

This opens an isolated offscreen No-Communication session with the simulated
vacuum pump running, then uses the Sample Queue button with synthetic specimens.
The report records process launch
to the first sampled window, idle CPU time, working set and private bytes,
queue duration, published specimen and step counts, and the 95th percentile
and longest intervals between 20 ms Qt timer ticks. CPU percentage is relative
to one logical processor. The verifier samples Windows memory every 100 ms.
The report identifies the candidate revision and records the workload at each
run so the lab can compare equivalent results on its control computers.

The portable verifier opens Sample Index and Data Review at 150% and 200%
scaling on offscreen work areas. It checks that the registry and Data Review
controls fit, that the plot, measurement table, and specimen list are reachable,
and that a maximized window keeps its state. Lab sign-off covers monitor moves
and dock or undock events on the actual workstation displays.

`StationAcceptance.md` ships beside this guide. Its evidence audit checks the
recorded bench cases and attachment hashes for the candidate revision.

## Remaining full-replacement gates

This is a **pilot candidate**. Calibrated axial and transverse AF, independent
MCC ARM bias, capacitor pulse IRM with a dedicated backfield polarity relay,
and synchronized RRM have native execution and recovery paths. Each requires
accepted station configuration, physical interlock tests, and reference results.

- Transverse probe auto-positioning has native turning motor, calibrated ADwin
  AF and 908A DC peak integration. The station needs probe fixture clearance,
  peak fidelity and interruption acceptance with the actual meter.
- External thermal steps need a physical reconciliation against the furnace
  log and SQUID output. Automated furnace operation needs an identified
  controller, protocol, readback, and interlocks.
- Bartington and SQUID rod scans need bench acceptance for manual mounts and
  automatic tray loading. U-channel cores and EPR tubes need acceptance with the
  station's fixtures.

Vendor driver installation, fresh-Windows deployment, mixed-DPI display review,
physical fault and recovery tests, scientific reference comparisons, and lab
approval remain pending. The portable app relies on station driver installation.
Follow
`docs/rapid_hardware_acceptance_procedure_2026-08-29.md` for bench qualification.

## Verification checkpoint — 3 October 2026

The continuation suite passes all 599 tests in 82.771 seconds. The copied
portable bundle passes dependency/icon discovery and the offline six-panel,
six-helper UI smoke check on this Windows 10 computer. Packaging excludes
incompatible ICU/CRT shims discovered on the build host's tooling PATH, so Qt
uses the Windows ICU API. Reproduce the checks with:

```powershell
.\installer\verify_rapid_main.ps1 -Bundle .\dist\RapidPyMain
```

This validates offline launch on this computer; it does not replace a fresh
Windows deployment or instrument/scientific acceptance.
