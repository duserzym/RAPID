"""Compose empty-hole pose, owned blank measurement and verified rod return."""
import unittest
from unittest.mock import patch

from rapid_main.holder_state import HolderStateStore
from tests import test_queue_backend_coordinator as backend_fixture
from tests import test_queue_holder_geometry as holder_fixture


class QueueHolderCoordinatorTests(backend_fixture.QueueBackendCoordinatorFixture, unittest.TestCase):
    def setUp(self):
        super().setUp()
        self.backend._holder_store = HolderStateStore(self.path.parent / 'holder.json')
        self.backend._direction_up = True
        self.backend._susceptibility_records = []

    def test_holder_command_never_picks_specimen_and_clears_after_owned_measurement(self):
        with holder_fixture.QueueHolderGeometryTests.injected_squid(self), patch.object(self.motor, 'sample_pickup') as pickup:
            worker = self.command(self.backend.measure_queue_holder)
        self.assertTrue(worker.ok, worker.error)
        pickup.assert_not_called()
        self.assertEqual(self.backend._holder_store.current.holder_id, 'holder-046')
        self.assertEqual(self.store.latest_holder_context(self.session.token)['phase'], 'clear')
        self.assertIsNone(self.backend._queue_specimen_geometry)
        self.assertEqual(self.lift_serial().position, 0)
        self.assertFalse(self.vacuum.is_valve_connected())
        self.assertTrue(self.vacuum.is_pump_on())
        self.assertIsNotNone(self.store.pending())

    def test_acquisition_failure_preserves_bound_blank_and_pending_stage(self):
        with patch.object(self.backend, 'read_squid', side_effect=RuntimeError('instrument failed')):
            worker = self.command(self.backend.measure_queue_holder)
        self.assertFalse(worker.ok)
        self.assertTrue(self.backend._queue_specimen_geometry.is_holder)
        self.assertFalse(self.vacuum.is_valve_connected())
        self.assertIsNone(self.backend._holder_store.current)
        # A failed acquisition function cannot invent an accepted correction.
        self.assertIsNotNone(self.store.pending())

    def test_nonempty_marker_and_active_specimen_fail_before_outputs(self):
        before = len(self.commands)
        worker = self.command(lambda: self.backend.measure_queue_holder(1))
        self.assertFalse(worker.ok)
        self.assertEqual(len(self.commands), before)
        self.load_backend()
        before = len(self.commands)
        worker = self.command(self.backend.measure_queue_holder)
        self.assertFalse(worker.ok)
        self.assertEqual(len(self.commands), before)
        self.assertTrue(self.vacuum.is_valve_connected())

    def test_holder_uses_vb6_up_orientation_then_restores_the_selected_file_direction(self):
        self.backend._direction_up = False
        original = self.backend.measure_bound_holder
        def measure():
            self.assertIs(self.backend._direction_up, True)
            return original()
        with holder_fixture.QueueHolderGeometryTests.injected_squid(self), patch.object(self.backend, 'measure_bound_holder', measure):
            worker = self.command(self.backend.measure_queue_holder)
        self.assertTrue(worker.ok, worker.error)
        self.assertFalse(self.backend._direction_up)
