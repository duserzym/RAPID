# SQUID read rate, live stream and motion capture

Status (2026-10-07): **software implemented and simulator/fake-tested; not yet bench-validated.**
None of this changes how a bracketed measurement is taken, reduced or published.

## What it is for

The bracketed measurement reads the 2G SQUID only with the specimen at rest:
zero, four orientations (0/90/180/270°), zero. The specimen also spends several
seconds travelling down the borehole into the pickup coils, and turning 90°
between orientations. During those motions the signal is changing in a way the
physics predicts exactly:

* **Descent / ascent.** Each axis traces the coil sensitivity function scaled by
  the specimen moment. Fitting a peak (or a measured coil-response template)
  gives the response amplitude and the lift height of maximum coupling, which
  also checks the configured measurement position.
* **90° turn.** The horizontal moment sweeps the X and Y coils. Each follows a
  first harmonic of the turn angle, and a rigid rotation forces X and Y to share
  one amplitude 90° apart. Fitting every sample along the path, rather than four
  stops, averages uncorrelated noise across the turn and reports how well the
  data obey the rigid-rotation model.

Recording those intervals continuously, then filtering and fitting them, may
add resolution and sensitivity on weak specimens. This is the hypothesis to
test on the bench (see *Validation plan*).

## Operator controls

**SQUID → Communication Settings → Continuous read**

| Control | Meaning |
|---|---|
| Live Stream… | Opens the live stream window (below). |
| Capture while moving | Opt-in. Records the SQUID during the descent, the three 90° turns and the ascent of every measurement block. |
| Descent / Turns / Ascent | Choose which motions are recorded. |
| Output folder | Default `<data folder>/squid_motion_capture` (or `~/.rapidpy/squid_motion_capture` if no data folder is set). |

Settings are applied only when **Save Settings** is pressed.

**Live Stream & Spectrum window**

| Control | Meaning |
|---|---|
| Sample every | Requested interval between sample starts. `0` = as fast as the serial link allows. |
| Axes | X+Y+Z, Z only (fastest) or X+Y. |
| Counter read | Re-read the flux counter every N samples and reuse it in between. `1` (default) is exact. Larger values roughly halve the per-axis cost, but a flux jump is not seen until the next counter read. Samples that reuse a counter are flagged `counts_fresh=0` in the CSV. |
| Latch count / data hold | Pauses after `LatchCount` / `LatchData`. The VB6 defaults are 100 / 120 ms. Shorter holds raise the rate. Whether the 2G electronics tolerate them must be confirmed on the bench. |
| Display filter | Zero-phase low-pass and notch, drawn as a dashed overlay. Statistics and the spectrum always describe the raw signal. |
| Save CSV / Open capture | Save a live stream; open a live or motion-capture CSV to see its spectrum and the descent/rotation fit. |
| Use these rate settings by default | Copies the rate settings into the Communication Settings dialog (saved with *Save Settings*). They are also used by motion capture. |

The window shows the **link-limited rate** before you start, and the **achieved
rate** while streaming. Each sample costs one latch (two commands plus the
holds) and one or two query round-trips per axis. By the built-in estimate, a
full X/Y/Z sample with the VB6 holds takes about 1.2 s at 1200 baud and about
0.36 s at 9600 baud. Z-only with no holds and a counter read every 10 samples
takes about 35 ms at 9600 baud. Real controller latency adds to these figures.

Bracketed static reads always keep the VB6 latch timing, whatever is set here.

## Files

Each block with capture enabled writes one folder:

```
squid_motion_capture/20261007T183012Z_block-000042/
  01-descent-measurement.csv
  02-turn-90deg.csv
  03-turn-180deg.csv
  04-turn-270deg.csv
  05-ascent-zero.csv
  analysis.json
```

* **CSV.** The first line is `# {json summary}`: label, timing, requested and
  achieved rate, motion target/actual, errors and stop reason. Then one row per
  sample: `index, t_s, duration_s, raw_x..z, counts_x..z, dvm_x..z, counts_fresh`.
  `raw = -dvm - counts * range` (VB6 `getVal`). Unread axes are `nan`. After a
  `# positions` line come the encoder samples `t_s, position`: lift counts for
  descent/ascent, degrees for turns, on the same `perf_counter` time base.
* **analysis.json.** Per-axis noise summary (drift, white-noise density,
  spectral lines) plus the rotation fit or per-axis pass-through fit for each
  segment.

These files are auxiliary. They are written outside the scientific run bundle,
and are not referenced by the provenance index or the published outputs.

## Safety and evidence boundaries

* Capture runs only while a lift or turn is in progress. The motors use their
  own serial port.
* Each capture is stopped and its worker thread joined **before** the next SQUID
  command. If the worker cannot release the link within its bounded timeout, the
  block fails with `MotionCaptureError` and is never retried on a shared link.
* A stream transport error is recorded on the trace. It never fails the block.
* Traces are attached to `BracketedAcquisition.motion_traces`. They never enter
  the six bracketed observations, the zero-pair validation, the reduction, the
  holder correction or published results.
* The live stream runs inside the SQUID dialog's exclusive `squid` ownership
  lease, so it cannot run during a measurement or queue.
* In NO-COMM mode the live stream uses a clearly labelled simulated signal.
  Hardware mode never falls back to simulation.

## Implementation map

| Piece | Location |
|---|---|
| Read plan, rate estimate, paced sampler, background capture, CSV codec | `RapidPy/rapidpy_common/squid_stream.py` |
| Resampling, Welch PSD, noise floor/lines, low-pass, notch, rotation and pass-through fits | `RapidPy/rapidpy_common/signal_analysis.py` (NumPy only) |
| Per-call latch holds, DVM-only read, reply timeout | `RapidPy/updown_control/updown_control/app.py` (`RawSquidClient`) |
| In-motion encoder position observer | `RapidPy/rapidpy_common/hardware.py` (`set_position_observer`) and `rapid_main/squid_transport.py` (`observe_positions`) |
| Block capture, analysis and sidecar files | `RapidPy/rapid_main/rapid_main/motion_capture.py` |
| Acquisition hook | `rapid_main/acquisition.py` (`motion_capture=` on `BracketedAcquisitionService`) |
| Settings | `SquidConfig.stream_*`, `SquidConfig.motion_capture_*` |
| UI | `rapid_main/dialogs/squid_comm.py`, `rapid_main/dialogs/squid_stream.py` |
| Tests | `rapid_main/tests/test_squid_stream.py`, `test_signal_analysis.py`, `test_motion_capture.py`, `test_squid_stream_dialog.py` |

## Validation plan (bench)

1. **Achievable rate.** At the station baud, stream with the VB6 holds, then
   progressively shorter holds. Record the achieved rate and confirm that replies
   stay well formed (no timeouts, no stale lines).
2. **Noise floor.** With an empty holder at the measurement position, stream for
   10 minutes per setting. Compare white-noise density and spectral lines, and
   look for motor or mains aliases worth a notch.
3. **Counter refresh.** On a strong specimen, compare `counts_every = 1` with
   larger values across descent/ascent, where flux jumps are most likely.
4. **Coil response template.** Pass a small point-like standard through the coils
   at slow lift speed, then save and average the descent traces. Fit the
   template, and check that the peak height agrees with the configured
   measurement position.
5. **Rotation fit versus bracketed reduction.** On standards and weak natural
   specimens, compare the horizontal amplitude/phase from the turn fits with the
   bracketed four-position result, using repeated blocks. Quantify whether the
   continuous estimate has lower scatter.
6. Only after steps 1–5 should any continuous estimate be considered for
   publication. That would need documented scientific acceptance, as
   `ROADMAP.md` requires.
