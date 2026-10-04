"""Device ownership guard tests for rapid_main hardware-service safety."""
from __future__ import annotations

import sys
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[2]))

from rapid_main.device_ownership import DeviceOwnershipError, DeviceOwnershipManager


class TestDeviceOwnershipManager(unittest.TestCase):
    def test_group_conflict_leaves_every_resource_and_count_unchanged(self):
        mgr = DeviceOwnershipManager()
        original = mgr.acquire('squid', 'diagnostic')
        with self.assertRaises(DeviceOwnershipError):
            mgr.acquire_many(('changer', 'squid', 'vacuum'), 'queue')
        self.assertIsNone(mgr.owner_of('changer'))
        self.assertIsNone(mgr.owner_of('vacuum'))
        self.assertEqual(mgr.owner_of('squid'), 'diagnostic')
        original.release()
        self.assertIsNone(mgr.owner_of('squid'))

    def test_group_keeps_reentrant_worker_counts_and_releases_once(self):
        mgr = DeviceOwnershipManager()
        group = mgr.acquire_many(('changer', 'measurement'), 'queue')
        worker = mgr.acquire('measurement', 'queue')
        group.release()
        self.assertIsNone(mgr.owner_of('changer'))
        self.assertEqual(mgr.owner_of('measurement'), 'queue')
        group.release()
        self.assertEqual(mgr.owner_of('measurement'), 'queue')
        worker.release()
        later = mgr.acquire_many(('changer', 'measurement'), 'queue')
        group.release()
        self.assertEqual(mgr.owner_of('changer'), 'queue')
        later.release()
        self.assertIsNone(mgr.owner_of('measurement'))

    def test_duplicate_group_resources_are_rejected_without_reserving(self):
        mgr = DeviceOwnershipManager()
        with self.assertRaises(DeviceOwnershipError): mgr.acquire_many(('changer', 'changer'), 'queue')
        self.assertIsNone(mgr.owner_of('changer'))

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
