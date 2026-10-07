"""Spectral analysis, zero-phase filters, and motion model fits for SQUID traces."""
from __future__ import annotations

import math
import unittest

import numpy as np

from rapidpy_common.signal_analysis import (
    SignalAnalysisError,
    fit_harmonic,
    fit_pass_through,
    fit_rotation,
    interpolate_motion,
    lowpass,
    noise_summary,
    notch,
    resample_uniform,
    welch_psd,
)


class SpectrumTests(unittest.TestCase):
    def setUp(self):
        self.rng = np.random.default_rng(7)

    def test_resample_handles_irregular_unsorted_records(self):
        t = np.array([0.0, 0.2, 0.1, 0.35, 0.4, 0.6, np.nan])
        y = 2.0 * np.nan_to_num(t)
        grid, values, dt = resample_uniform(t, y)
        self.assertAlmostEqual(dt, 0.1, places=6)  # median spacing
        np.testing.assert_allclose(values, 2.0 * grid)

    def test_psd_integrates_to_variance(self):
        fs = 20.0
        y = self.rng.normal(0.0, 0.3, 4096)
        freqs, psd = welch_psd(y, fs)
        variance = float(np.sum(psd) * (freqs[1] - freqs[0]))
        self.assertAlmostEqual(variance, 0.09, delta=0.01)

    def test_noise_summary_reports_drift_floor_and_line(self):
        fs = 10.0
        t = np.arange(0, 120, 1 / fs)
        y = 0.4 * np.sin(2 * np.pi * 1.7 * t) + self.rng.normal(0, 0.05, t.size) + 0.02 * t
        summary = noise_summary(t, y)
        self.assertAlmostEqual(summary.sample_rate_hz, fs, places=6)
        self.assertAlmostEqual(summary.drift_per_s, 0.02, delta=0.002)
        self.assertAlmostEqual(summary.lines[0].frequency_hz, 1.7, delta=0.05)
        # White floor of N(0, 0.05) sampled at 10 Hz: sigma / sqrt(fs / 2).
        self.assertAlmostEqual(summary.noise_density, 0.05 / math.sqrt(fs / 2), delta=0.01)

    def test_notch_removes_line_without_shifting_mean(self):
        fs = 10.0
        t = np.arange(0, 60, 1 / fs)
        clean = 1.5 + self.rng.normal(0, 0.02, t.size)
        y = clean + 0.5 * np.sin(2 * np.pi * 2.0 * t)
        filtered = notch(y, fs, 2.0, width_hz=0.3)
        self.assertLess(np.std(filtered - clean), 0.03)
        self.assertAlmostEqual(np.mean(filtered), np.mean(y), places=9)

    def test_lowpass_is_zero_phase(self):
        fs = 20.0
        t = np.arange(0, 20, 1 / fs)
        slow = np.sin(2 * np.pi * 0.2 * t)
        y = slow + 0.3 * np.sin(2 * np.pi * 6.0 * t)
        filtered = lowpass(y, fs, 1.0, numtaps=41)
        core = slice(60, -60)
        self.assertLess(np.max(np.abs(filtered[core] - slow[core])), 0.05)
        lag = np.argmax(np.correlate(filtered[core], slow[core], "full")) - (slow[core].size - 1)
        self.assertEqual(lag, 0)

    def test_short_records_are_rejected(self):
        with self.assertRaises(SignalAnalysisError):
            resample_uniform([0, 1, 2], [1, 2, 3])
        with self.assertRaises(SignalAnalysisError):
            lowpass(np.ones(50), 10.0, 6.0)


class RotationFitTests(unittest.TestCase):
    def setUp(self):
        self.rng = np.random.default_rng(11)

    def _turn(self, start, stop, n, sense, amplitude=2.0, phase_deg=35.0, noise=0.03):
        theta = np.linspace(start, stop, n)
        rad = np.radians(theta + phase_deg)
        x = 0.7 + amplitude * np.cos(rad) + self.rng.normal(0, noise, n)
        y = -0.4 + sense * amplitude * np.sin(rad) + self.rng.normal(0, noise, n)
        z = 1.1 + self.rng.normal(0, noise, n)
        return theta, x, y, z

    def test_full_turn_recovers_moment_phase_and_handedness(self):
        for sense in (1, -1):
            theta, x, y, z = self._turn(0, 360, 180, sense)
            fit = fit_rotation(theta, x, y, z)
            self.assertEqual(fit.rotation_sense, sense)
            self.assertAlmostEqual(fit.horizontal_amplitude, 2.0, delta=0.02)
            self.assertAlmostEqual(fit.horizontal_phase_deg, 35.0, delta=1.0)
            self.assertAlmostEqual(fit.xy_consistency, 1.0, delta=0.03)
            self.assertLess(fit.axes["Z"].amplitude, 0.03)

    def test_single_quarter_turn_still_constrains_amplitude(self):
        theta, x, y, _ = self._turn(90, 180, 30, 1)
        fit = fit_rotation(theta, x, y)
        self.assertAlmostEqual(fit.angle_span_deg, 90.0)
        self.assertAlmostEqual(fit.horizontal_amplitude, 2.0, delta=0.1)
        self.assertGreater(fit.horizontal_sigma, 0.0)

    def test_continuous_fit_beats_four_point_noise(self):
        # Same noise per reading: many samples along the path shrink the uncertainty.
        theta, x, y, _ = self._turn(0, 360, 400, 1, noise=0.2)
        dense = fit_rotation(theta, x, y)
        sparse = fit_rotation(*[arr[::40] for arr in (theta, x, y)])
        self.assertLess(dense.horizontal_sigma, sparse.horizontal_sigma / 2)

    def test_harmonic_fit_separates_drift(self):
        theta = np.linspace(0, 360, 100)
        t = np.linspace(0, 10, 100)
        values = 1.0 + 0.5 * np.cos(np.radians(theta)) + 0.03 * t
        fit = fit_harmonic(theta, values, t)
        self.assertAlmostEqual(fit.drift_per_s, 0.03, places=6)
        self.assertAlmostEqual(fit.amplitude, 0.5, places=6)


class PassThroughFitTests(unittest.TestCase):
    def test_gaussian_peak_finds_height_of_maximum_coupling(self):
        rng = np.random.default_rng(3)
        z = np.linspace(-26000, -31000, 160)
        values = 0.1 + 2e-6 * (z - z.mean()) + 4.0 * np.exp(-0.5 * ((z + 30100) / 600) ** 2) + rng.normal(0, 0.04, z.size)
        fit = fit_pass_through(z, values)
        self.assertAlmostEqual(fit.center, -30100, delta=60)
        self.assertAlmostEqual(fit.amplitude, 4.0, delta=0.15)
        self.assertGreater(fit.snr, 50)

    def test_measured_template_is_scaled_and_shifted(self):
        offsets = np.linspace(-2000, 2000, 201)
        template = np.exp(-0.5 * (offsets / 500) ** 2) - 0.3 * np.exp(-0.5 * ((offsets - 900) / 300) ** 2)
        z = np.linspace(-1500, 3500, 120)
        values = 0.2 - 2.5 * np.interp(z - 1200, offsets, template, left=0, right=0)
        fit = fit_pass_through(z, values, template_position=offsets, template=template)
        self.assertEqual(fit.model, "template")
        self.assertAlmostEqual(fit.amplitude, -2.5, delta=0.05)
        self.assertAlmostEqual(fit.center, 1200, delta=50)

    def test_motion_interpolation_prefers_observed_positions(self):
        path = interpolate_motion([0.0, 0.5, 1.0, 2.0], [0.0, 1.0], [10.0, 30.0])
        np.testing.assert_allclose(path, [10.0, 20.0, 30.0, 30.0])
        uniform = interpolate_motion([0.0, 1.0, 2.0], start=0.0, end=90.0, t_start=0.0, t_end=2.0)
        np.testing.assert_allclose(uniform, [0.0, 45.0, 90.0])
        with self.assertRaises(SignalAnalysisError):
            interpolate_motion([0.0], start=None, end=1, t_start=0, t_end=1)


if __name__ == "__main__":
    unittest.main()
