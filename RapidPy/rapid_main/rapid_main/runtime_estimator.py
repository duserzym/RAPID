"""
runtime_estimator.py — Estimate sequence run time for RAPID v4.

Provides per-step-type time estimates (configurable) and computes
total/remaining sequence durations for display in the status bar.

Default estimates are based on typical RAPID magnetometer operation:
  NRM measurement:   ~120 s
  AF demagnetisation step: ~180 s (includes coil ramp + settle + measure)
  Thermal step:      ~600 s (placeholder; actual time depends on temp ramp)
  IRM acquisition:   ~300 s
  ARM acquisition:   ~240 s
  Unknown:           ~120 s
"""
from __future__ import annotations

import re
from datetime import datetime, timedelta
from typing import Optional


# ---------------------------------------------------------------------------
# Default per-step time estimates (seconds)
# ---------------------------------------------------------------------------

DEFAULT_STEP_TIMES: dict[str, int] = {
    "NRM":   120,
    "AF":    180,
    "AFMAX": 180,
    "AFZ":   180,
    "TT":    600,
    "TH":    600,
    "TEMP":  600,
    "IRM":   300,
    "ARM":   240,
    "PTRM":  600,
    "ZF":    120,
    "IF":    120,
    "_default": 120,
}


def _step_type(label: str) -> str:
    """Extract the step-type key (uppercase prefix) from a demag label."""
    upper = label.strip().upper()
    for key in ("AFMAX", "AFZ", "PTRM", "ARM", "IRM", "AF", "TT", "TH", "TEMP", "NRM", "ZF", "IF"):
        if upper.startswith(key):
            return key
    return "_default"


class RuntimeEstimator:
    """
    Estimates sequence run time from a list of step labels.

    Parameters
    ----------
    step_times:
        Dict mapping step-type key → seconds.  Defaults to DEFAULT_STEP_TIMES.
    """

    def __init__(self, step_times: Optional[dict[str, int]] = None) -> None:
        self._times = dict(DEFAULT_STEP_TIMES)
        if step_times:
            self._times.update(step_times)

    @property
    def step_times(self) -> dict[str, int]:
        """Current per-step-type time mapping (seconds)."""
        return self._times

    @step_times.setter
    def step_times(self, new_times: dict[str, int]) -> None:
        """Replace the time mapping (used when settings change at runtime)."""
        self._times = dict(DEFAULT_STEP_TIMES)
        self._times.update(new_times)

    def step_seconds(self, label: str) -> int:
        """Return estimated seconds for one step with the given label."""
        key = _step_type(label)
        return self._times.get(key, self._times["_default"])

    def estimate_sequence(self, labels: list[str]) -> timedelta:
        """Total estimated duration for all steps."""
        total = sum(self.step_seconds(lbl) for lbl in labels)
        return timedelta(seconds=total)

    def estimate_remaining(self, labels: list[str], current_idx: int) -> timedelta:
        """
        Estimated remaining time from *current_idx* (inclusive) to end.

        Parameters
        ----------
        labels:
            Full list of step labels in execution order.
        current_idx:
            0-based index of the step currently executing (or about to execute).
        """
        remaining = sum(
            self.step_seconds(lbl)
            for lbl in labels[current_idx:]
        )
        return timedelta(seconds=remaining)

    def format_duration(self, td: timedelta) -> str:
        """Format a timedelta as ``"Xh Ym"`` or ``"Ym Zs"``."""
        total_secs = int(td.total_seconds())
        if total_secs < 0:
            total_secs = 0
        hours, rem = divmod(total_secs, 3600)
        minutes, secs = divmod(rem, 60)
        if hours > 0:
            return f"{hours}h {minutes:02d}m"
        if minutes > 0:
            return f"{minutes}m {secs:02d}s"
        return f"{secs}s"

    def status_bar_text(
        self,
        labels: list[str],
        current_idx: int,
        *,
        start_time: Optional[datetime] = None,
    ) -> str:
        """
        Build the runtime status bar segment string.

        Returns something like:
          ``"Est. remaining: 2h 14m  |  Step 12/47  |  ETA: 16:32"``
        """
        total = len(labels)
        remaining = self.estimate_remaining(labels, current_idx)
        rem_str = self.format_duration(remaining)

        parts = [
            f"Est. remaining: {rem_str}",
            f"Step {current_idx + 1}/{total}",
        ]

        if start_time is not None:
            eta = datetime.now() + remaining
            parts.append(f"ETA: {eta.strftime('%H:%M')}")

        return "  |  ".join(parts)

    def total_bar_text(self, labels: list[str]) -> str:
        """
        Summary text shown when a sequence is loaded but not yet running.

        Returns something like: ``"Sequence: 47 steps  ~  2h 34m total"``
        """
        td = self.estimate_sequence(labels)
        return f"Sequence: {len(labels)} steps  ~  {self.format_duration(td)} total"
