"""Spectral analysis, filtering and model fits for continuous SQUID traces.

Everything here is NumPy-only (no SciPy dependency in the packaged apps) and
works on short, irregularly sampled records like the ones produced by
:mod:`rapidpy_common.squid_stream` while a specimen moves through the
pickup coils or turns between orientations.

Two physical models give the continuous record its extra value:

* **Rotation** -- while the specimen turns about the vertical axis, its
  horizontal moment sweeps the X/Y coils.  Each horizontal axis follows a
  first harmonic of the turn angle, ``a0 + a1 cos(theta) + b1 sin(theta)``,
  and a rigid rotation forces X and Y to share one amplitude 90 degrees
  apart.  Fitting every sample along the path, rather than four stops,
  averages uncorrelated noise across the whole turn and reports how well the
  data obey the rigid-rotation model.
* **Pass-through** -- while the lift lowers the specimen, each axis traces the
  coil sensitivity function scaled by the moment.  A peak model (Gaussian, or
  a measured coil-response template) on a linear baseline gives the response
  amplitude and the height of maximum coupling, which is also a direct check
  of the configured measurement position.

These fits are diagnostics and research tools.  They have not been validated
against the bracketed four-position reduction and are never used to publish
a measurement.
"""

from __future__ import annotations

import math
from dataclasses import dataclass, field

import numpy as np


class SignalAnalysisError(ValueError):
    """Raised when a record is too short or malformed for the requested analysis."""


# ── resampling and spectra ──────────────────────────────────────────────────


def _clean(t, y) -> tuple[np.ndarray, np.ndarray]:
    t = np.asarray(t, dtype=float)
    y = np.asarray(y, dtype=float)
    if t.shape != y.shape:
        raise SignalAnalysisError("time and value arrays must have the same length")
    keep = np.isfinite(t) & np.isfinite(y)
    t, y = t[keep], y[keep]
    order = np.argsort(t, kind="stable")
    t, y = t[order], y[order]
    if t.size:
        unique = np.concatenate(([True], np.diff(t) > 0))
        t, y = t[unique], y[unique]
    return t, y


def resample_uniform(t, y, dt: float | None = None) -> tuple[np.ndarray, np.ndarray, float]:
    """Linearly resample an irregular record onto a uniform grid.

    ``dt`` defaults to the median sample spacing, so the grid matches the
    achieved stream rate.  Returns ``(t_uniform, y_uniform, dt)``.
    """

    t, y = _clean(t, y)
    if t.size < 4:
        raise SignalAnalysisError("at least four finite samples are required")
    if dt is None:
        dt = float(np.median(np.diff(t)))
    if not math.isfinite(dt) or dt <= 0:
        raise SignalAnalysisError("sample spacing must be positive")
    count = int(math.floor((t[-1] - t[0]) / dt)) + 1
    grid = t[0] + dt * np.arange(count)
    return grid, np.interp(grid, t, y), float(dt)


def detrend(y, order: int = 1) -> np.ndarray:
    """Remove a least-squares polynomial (``order`` 0 = mean, 1 = linear drift)."""

    y = np.asarray(y, dtype=float)
    if order < 0:
        return y.copy()
    x = np.linspace(-1.0, 1.0, y.size)
    coeffs = np.polyfit(x, y, min(order, max(0, y.size - 1)))
    return y - np.polyval(coeffs, x)


def welch_psd(y, fs: float, *, segment: int | None = None, overlap: float = 0.5, detrend_order: int = 1):
    """One-sided Welch power spectral density (units^2 / Hz) with a Hann window."""

    y = np.asarray(y, dtype=float)
    if y.size < 8:
        raise SignalAnalysisError("at least eight samples are required for a spectrum")
    if fs <= 0:
        raise SignalAnalysisError("sample rate must be positive")
    if segment is None:
        segment = int(2 ** math.floor(math.log2(max(8, min(y.size, 256)))))
        if y.size >= 64:
            segment = int(2 ** math.floor(math.log2(y.size / 2)))
    segment = int(min(max(8, segment), y.size))
    step = max(1, int(segment * (1.0 - overlap)))
    window = np.hanning(segment)
    scale = 1.0 / (fs * np.sum(window**2))
    spectra = []
    for start in range(0, y.size - segment + 1, step):
        chunk = detrend(y[start:start + segment], detrend_order) * window
        spectrum = np.abs(np.fft.rfft(chunk)) ** 2 * scale
        if segment % 2 == 0:
            spectrum[1:-1] *= 2.0
        else:
            spectrum[1:] *= 2.0
        spectra.append(spectrum)
    freqs = np.fft.rfftfreq(segment, d=1.0 / fs)
    return freqs, np.mean(spectra, axis=0)


@dataclass(frozen=True)
class SpectralLine:
    frequency_hz: float
    power: float
    prominence: float  # ratio to the median noise floor


@dataclass(frozen=True)
class NoiseSummary:
    sample_rate_hz: float
    samples: int
    mean: float
    std: float
    detrended_std: float
    drift_per_s: float
    noise_density: float  # sqrt(median PSD), units / sqrt(Hz)
    white_noise_rms: float  # noise_density * sqrt(Nyquist bandwidth)
    lines: tuple[SpectralLine, ...] = ()


def find_spectral_lines(freqs, psd, *, prominence: float = 8.0, max_lines: int = 5, skip_dc_bins: int = 1) -> tuple[SpectralLine, ...]:
    """Local PSD maxima standing ``prominence`` times above the median floor."""

    freqs = np.asarray(freqs, dtype=float)
    psd = np.asarray(psd, dtype=float)
    if psd.size < 3:
        return ()
    floor = float(np.median(psd[skip_dc_bins:])) or float(np.finfo(float).tiny)
    lines = []
    for i in range(max(1, skip_dc_bins), psd.size - 1):
        if psd[i] >= psd[i - 1] and psd[i] >= psd[i + 1] and psd[i] >= prominence * floor:
            lines.append(SpectralLine(float(freqs[i]), float(psd[i]), float(psd[i] / floor)))
    lines.sort(key=lambda line: line.power, reverse=True)
    return tuple(lines[:max_lines])


def noise_summary(t, y, *, prominence: float = 8.0) -> NoiseSummary:
    """Characterise a (nominally static) record: drift, white noise floor and lines."""

    grid, uniform, dt = resample_uniform(t, y)
    fs = 1.0 / dt
    slope = float(np.polyfit(grid - grid[0], uniform, 1)[0]) if uniform.size > 1 else 0.0
    freqs, psd = welch_psd(uniform, fs)
    floor = float(np.median(psd[1:])) if psd.size > 1 else float(psd[0])
    density = math.sqrt(max(floor, 0.0))
    return NoiseSummary(
        sample_rate_hz=fs,
        samples=int(uniform.size),
        mean=float(np.mean(uniform)),
        std=float(np.std(uniform)),
        detrended_std=float(np.std(detrend(uniform, 1))),
        drift_per_s=slope,
        noise_density=density,
        white_noise_rms=density * math.sqrt(fs / 2.0),
        lines=find_spectral_lines(freqs, psd, prominence=prominence),
    )


# ── zero-phase filtering ────────────────────────────────────────────────────


def lowpass_taps(cutoff_hz: float, fs: float, numtaps: int = 31) -> np.ndarray:
    """Windowed-sinc (Hamming) low-pass FIR with unity DC gain."""

    if not 0 < cutoff_hz < fs / 2:
        raise SignalAnalysisError("cutoff must be between 0 and the Nyquist frequency")
    numtaps = int(numtaps) | 1  # odd length keeps the filter symmetric
    n = np.arange(numtaps) - (numtaps - 1) / 2
    taps = np.sinc(2 * cutoff_hz / fs * n) * np.hamming(numtaps)
    return taps / np.sum(taps)


def filtfilt_fir(y, taps) -> np.ndarray:
    """Zero-phase FIR filtering (forward then backward) with reflected edges."""

    y = np.asarray(y, dtype=float)
    taps = np.asarray(taps, dtype=float)
    pad = min(y.size - 1, taps.size * 3)
    if pad < 1:
        return y.copy()
    padded = np.concatenate((2 * y[0] - y[pad:0:-1], y, 2 * y[-1] - y[-2:-pad - 2:-1]))
    forward = np.convolve(padded, taps, mode="same")
    backward = np.convolve(forward[::-1], taps, mode="same")[::-1]
    return backward[pad:pad + y.size]


def lowpass(y, fs: float, cutoff_hz: float, numtaps: int = 31) -> np.ndarray:
    """Zero-phase low-pass; preserves edges and the timing of features."""

    y = np.asarray(y, dtype=float)
    numtaps = min(int(numtaps) | 1, max(3, (y.size // 3) | 1))
    return filtfilt_fir(y, lowpass_taps(cutoff_hz, fs, numtaps))


def notch(y, fs: float, freq_hz: float, width_hz: float | None = None) -> np.ndarray:
    """Zero-phase frequency-domain notch with a raised-cosine edge.

    Removes a narrow interference line (for example a motor-step tone or an
    aliased mains harmonic) without shifting the rest of the record.
    """

    y = np.asarray(y, dtype=float)
    if not 0 < freq_hz < fs / 2:
        raise SignalAnalysisError("notch frequency must be between 0 and Nyquist")
    width = float(width_hz) if width_hz else max(fs / max(y.size, 1) * 2.0, freq_hz * 0.05)
    mean = float(np.mean(y))
    spectrum = np.fft.rfft(y - mean)
    freqs = np.fft.rfftfreq(y.size, d=1.0 / fs)
    distance = np.abs(freqs - freq_hz) / width
    gain = np.where(distance >= 1.0, 1.0, 0.5 - 0.5 * np.cos(np.pi * np.clip(distance, 0.0, 1.0)))
    return np.fft.irfft(spectrum * gain, n=y.size) + mean


# ── rotation (90-degree turn) model ─────────────────────────────────────────


@dataclass(frozen=True)
class HarmonicFit:
    offset: float
    cos_coeff: float
    sin_coeff: float
    drift_per_s: float
    amplitude: float
    phase_deg: float
    residual_rms: float
    amplitude_sigma: float


@dataclass(frozen=True)
class RotationFit:
    """First-harmonic fits of each axis against turn angle, plus a joint X/Y fit."""

    samples: int
    angle_span_deg: float
    axes: dict[str, HarmonicFit] = field(default_factory=dict)
    horizontal_amplitude: float = math.nan
    horizontal_phase_deg: float = math.nan
    horizontal_sigma: float = math.nan
    rotation_sense: int = 0  # +1 or -1: which handedness fits X/Y best
    xy_consistency: float = math.nan  # |X| / |Y| harmonic amplitude ratio (1 for a rigid rotation)


def _lstsq_with_sigma(design: np.ndarray, values: np.ndarray):
    coeffs, *_ = np.linalg.lstsq(design, values, rcond=None)
    residual = values - design @ coeffs
    dof = max(1, values.size - design.shape[1])
    sigma2 = float(residual @ residual) / dof
    try:
        covariance = sigma2 * np.linalg.inv(design.T @ design)
    except np.linalg.LinAlgError:
        covariance = np.full((design.shape[1], design.shape[1]), np.nan)
    return coeffs, residual, covariance


def fit_harmonic(angle_deg, values, times=None) -> HarmonicFit:
    """Fit ``offset + a cos(theta) + b sin(theta) [+ drift * t]``."""

    theta = np.radians(np.asarray(angle_deg, dtype=float))
    values = np.asarray(values, dtype=float)
    keep = np.isfinite(theta) & np.isfinite(values)
    columns = [np.ones(keep.sum()), np.cos(theta[keep]), np.sin(theta[keep])]
    with_drift = times is not None
    if with_drift:
        t = np.asarray(times, dtype=float)[keep]
        columns.append(t - t.mean())
    design = np.column_stack(columns)
    if keep.sum() <= design.shape[1]:
        raise SignalAnalysisError("not enough samples for a harmonic fit")
    coeffs, residual, cov = _lstsq_with_sigma(design, values[keep])
    a, b = float(coeffs[1]), float(coeffs[2])
    amplitude = math.hypot(a, b)
    if amplitude > 0:
        grad = np.array([a / amplitude, b / amplitude])
        amp_sigma = float(math.sqrt(max(0.0, grad @ cov[1:3, 1:3] @ grad)))
    else:
        amp_sigma = float(math.sqrt(max(0.0, cov[1, 1])))
    return HarmonicFit(
        offset=float(coeffs[0]),
        cos_coeff=a,
        sin_coeff=b,
        drift_per_s=float(coeffs[3]) if with_drift else 0.0,
        amplitude=amplitude,
        phase_deg=math.degrees(math.atan2(b, a)),
        residual_rms=float(np.sqrt(np.mean(residual**2))),
        amplitude_sigma=amp_sigma,
    )


def fit_rotation(angle_deg, x, y, z=None, times=None) -> RotationFit:
    """Fit a turn: per-axis harmonics and a joint rigid-rotation X/Y solution.

    The joint model is ``X = cx + A cos(theta + phi)``,
    ``Y = cy + s * A sin(theta + phi)`` with handedness ``s`` chosen by the
    smaller residual, solved linearly as ``X = cx + p cos - q sin``,
    ``Y = cy + s (p sin + q cos)``.  A short turn (a single 90-degree step)
    constrains the harmonic less than a full revolution; ``angle_span_deg``
    is reported so callers can judge that.
    """

    theta_deg = np.asarray(angle_deg, dtype=float)
    x = np.asarray(x, dtype=float)
    y = np.asarray(y, dtype=float)
    keep = np.isfinite(theta_deg) & np.isfinite(x) & np.isfinite(y)
    if keep.sum() < 6:
        raise SignalAnalysisError("at least six finite samples are required for a rotation fit")
    theta = np.radians(theta_deg[keep])
    span = float(np.nanmax(theta_deg[keep]) - np.nanmin(theta_deg[keep]))
    t = None if times is None else np.asarray(times, dtype=float)[keep]
    axes = {
        "X": fit_harmonic(theta_deg[keep], x[keep], t),
        "Y": fit_harmonic(theta_deg[keep], y[keep], t),
    }
    if z is not None:
        z = np.asarray(z, dtype=float)[keep]
        if np.isfinite(z).sum() > 4:
            axes["Z"] = fit_harmonic(theta_deg[keep], z, t)

    n = theta.size
    best = None
    for sense in (1, -1):
        design = np.zeros((2 * n, 4))
        design[:n, 0] = 1.0
        design[n:, 1] = 1.0
        design[:n, 2] = np.cos(theta)
        design[:n, 3] = -np.sin(theta)
        design[n:, 2] = sense * np.sin(theta)
        design[n:, 3] = sense * np.cos(theta)
        values = np.concatenate((x[keep], y[keep]))
        coeffs, residual, cov = _lstsq_with_sigma(design, values)
        rss = float(residual @ residual)
        if best is None or rss < best[0]:
            best = (rss, sense, coeffs, cov)
    _, sense, coeffs, cov = best
    p, q = float(coeffs[2]), float(coeffs[3])
    amplitude = math.hypot(p, q)
    if amplitude > 0:
        grad = np.array([p / amplitude, q / amplitude])
        sigma = float(math.sqrt(max(0.0, grad @ cov[2:4, 2:4] @ grad)))
    else:
        sigma = math.nan
    ratio = axes["X"].amplitude / axes["Y"].amplitude if axes["Y"].amplitude > 0 else math.nan
    return RotationFit(
        samples=int(n),
        angle_span_deg=span,
        axes=axes,
        horizontal_amplitude=amplitude,
        horizontal_phase_deg=math.degrees(math.atan2(q, p)),
        horizontal_sigma=sigma,
        rotation_sense=int(sense),
        xy_consistency=float(ratio),
    )


# ── pass-through (borehole descent) model ───────────────────────────────────


@dataclass(frozen=True)
class PassThroughFit:
    """Peak response of one axis as the specimen travels through the coils."""

    amplitude: float
    center: float
    width: float
    baseline_offset: float
    baseline_slope: float
    residual_rms: float
    snr: float
    samples: int
    model: str


def _peak_basis(z: np.ndarray, center: float, width: float, template_z=None, template=None) -> np.ndarray:
    if template is not None:
        return np.interp(z - center, np.asarray(template_z, dtype=float), np.asarray(template, dtype=float), left=0.0, right=0.0)
    return np.exp(-0.5 * ((z - center) / width) ** 2)


def fit_pass_through(position, values, *, template_position=None, template=None, centers: int = 121, widths: int = 25) -> PassThroughFit:
    """Fit ``baseline(z) + A * peak(z - z0)`` by variable projection.

    The peak is a Gaussian of free width, or -- when ``template`` is given -- a
    measured coil-response curve sampled at ``template_position`` offsets.
    The non-linear centre (and width) are grid-searched; amplitude and the
    linear baseline are solved exactly at every grid point.
    """

    z, v = _clean(position, values)
    if z.size < 8:
        raise SignalAnalysisError("at least eight finite samples are required for a pass-through fit")
    span = float(z[-1] - z[0])
    if span <= 0:
        raise SignalAnalysisError("positions must vary during a pass-through")
    zc = z - z.mean()
    center_grid = np.linspace(z[0], z[-1], int(centers))
    if template is not None:
        width_grid = np.array([math.nan])
    else:
        step = span / max(z.size - 1, 1)
        width_grid = np.geomspace(max(step, span / 200.0), span / 2.0, int(widths))

    best = None
    for width in width_grid:
        for center in center_grid:
            basis = _peak_basis(z, center, width, template_position, template)
            if not np.any(basis):
                continue
            design = np.column_stack((np.ones_like(z), zc, basis))
            coeffs, *_ = np.linalg.lstsq(design, v, rcond=None)
            residual = v - design @ coeffs
            rss = float(residual @ residual)
            if best is None or rss < best[0]:
                best = (rss, center, width, coeffs, residual)
    if best is None:
        raise SignalAnalysisError("pass-through model could not be evaluated over these positions")
    rss, center, width, coeffs, residual = best
    rms = math.sqrt(rss / z.size)
    amplitude = float(coeffs[2])
    return PassThroughFit(
        amplitude=amplitude,
        center=float(center),
        width=float(width),
        baseline_offset=float(coeffs[0]),
        baseline_slope=float(coeffs[1]),
        residual_rms=rms,
        snr=abs(amplitude) / rms if rms > 0 else math.inf,
        samples=int(z.size),
        model="template" if template is not None else "gaussian",
    )


# ── position / angle helpers ───────────────────────────────────────────────


def interpolate_motion(sample_times, position_times=None, positions=None, *, start=None, end=None, t_start=None, t_end=None) -> np.ndarray:
    """Place each SQUID sample on the motion path.

    Uses observed encoder positions when available; otherwise assumes uniform
    motion from ``start`` at ``t_start`` to ``end`` at ``t_end``.  Samples
    outside the observed span are clamped to the nearest observed position.
    """

    sample_times = np.asarray(sample_times, dtype=float)
    if position_times is not None and positions is not None and len(position_times) >= 2:
        pt, pv = _clean(position_times, positions)
        if pt.size >= 2:
            return np.interp(sample_times, pt, pv)
    if None in (start, end, t_start, t_end) or t_end <= t_start:
        raise SignalAnalysisError("motion path needs observed positions or start/end values and times")
    fraction = np.clip((sample_times - t_start) / (t_end - t_start), 0.0, 1.0)
    return float(start) + fraction * (float(end) - float(start))
