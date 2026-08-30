"""Device ownership guard tests for rapid_main hardware-service safety."""
from __future__ import annotations

import sys
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[2]))

from rapid_main.device_ownership import DeviceOwnershipError, DeviceOwnershipManager


class TestDeviceOwnershipManager(unittest.TestCase):
    def test_first_owner_acquires_successfully(self) -> None:
        mgr = DeviceOwnershipManager()
        lease = mgr.acquire("measurement", "panel_a")
        self.assertEqual(mgr.owner_of("measurement"), "panel_a")
        lease.release()
        self.assertIsNone(mgr.owner_of("measurement"))

    def test_conflicting_owner_is_blocked(self) -> None:
        mgr = DeviceOwnershipManager()
        _ = mgr.acquire("measurement", "panel_a")
        with self.assertRaises(DeviceOwnershipError) as exc:
            mgr.acquire("measurement", "panel_b")
        self.assertIn("in use", str(exc.exception))

    def test_reentrant_acquire_allowed_by_default(self) -> None:
        mgr = DeviceOwnershipManager()
        first = mgr.acquire("measurement", "panel_a")
        second = mgr.acquire("measurement", "panel_a")
        self.assertEqual(mgr.owner_of("measurement"), "panel_a")
        first.release()
        second.release()
        self.assertIsNone(mgr.owner_of("measurement"))

    def test_non_reentrant_blocked_for_second_call(self) -> None:
        mgr = DeviceOwnershipManager()
        _ = mgr.acquire("measurement", "panel_a")
        with self.assertRaises(DeviceOwnershipError):
            mgr.acquire("measurement", "panel_a", allow_reentrant=False)


if __name__ == "__main__":
    unittest.main(verbosity=2)
