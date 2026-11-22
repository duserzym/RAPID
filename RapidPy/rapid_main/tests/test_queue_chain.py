"""Chain queue integration with injected motors, vacuum and field circuits."""
import unittest
from unittest.mock import Mock, patch

from rapid_main.queue_restart import QueueRestartRecovery, RestartShutdownAttestation
from rapidpy_common.hardware_safety import HardwareSafetyError
from tests.test_queue_startup import QueueStartupFixture


class QueueChainTests(QueueStartupFixture, unittest.TestCase):
    def setUp(self):
        super().setUp()
        self.use_chain_station()

    def run_queue_command(self, action):
        worker = self.command(action)
        self.assertTrue(worker.ok, worker.error)

    def start_chain(self, slot=99):
        startup = self.prepare(chain_slot=slot, plan={'file_directions': {'F1': True}})
        self.run_queue_command(self.backend.start_queue_lifetime)
        startup.table.sleep = self.clock.sleep
        self.backend._queue_coordinator.lift.monotonic = self.clock.monotonic
        return startup

    def chain_serial(self):
        return self.motor._connections['COM3']._serial

    def test_full_transfer_routes_three_axes_wraps_and_returns_to_original_slot(self):
        startup = self.start_chain()
        self.assertEqual(set(self.motor._connections), {'COM3', 'COM5', 'COM6'})
        self.assertEqual(self.store.read()['plan']['startup']['chain_slot'], 99)
        self.assertEqual(self.store.read()['stage']['plan']['action'], 'reference_confirmed_chain_slot')
        self.assertFalse(any(b' 134 ' in row or b' 135 ' in row for row in self.chain_serial().writes))

        self.run_queue_command(self.backend.park_queue_station)
        self.assertEqual(self.store.read()['stage']['plan']['slot'], 100)
        self.assertLess(abs(self.store.read()['stage']['plan']['target_counts']
                            - self.store.read()['stage']['plan']['starting_counts']), 1050)
        self.run_queue_command(lambda: self.backend.record_queue_tray_confirmation('load', operator='Operator'))
        self.run_queue_command(lambda: self.backend.park_queue_station(1))
        self.assertEqual(self.store.read()['stage']['plan']['slot'], 1)
        self.run_queue_command(lambda: self.backend.park_queue_station(0))
        self.run_queue_command(lambda: self.backend.record_queue_tray_confirmation('flip', 'F1', operator='Operator'))
        with startup.session.claim():
            self.assertFalse(self.backend.queue_file_direction('F1'))

        with patch.object(self.backend, 'measure_bound_holder', return_value='blank') as acquisition:
            self.run_queue_command(self.backend.measure_queue_holder)
        acquisition.assert_called_once_with(averaging_cycles=None)
        self.assertEqual(self.store.read()['stage']['record']['holder_context']['hole'], 100)

        self.run_queue_command(lambda: self.backend.load_queue_specimen(1, 'S1', file_id='F1'))
        self.assertEqual(self.store.read()['stage']['plan']['slot'], 100)
        self.assertTrue(self.vacuum.is_valve_connected())
        with startup.session.claim():
            self.assertEqual(self.backend._sample_height(), 2800)
        lift = self.lift_serial()
        expected_dropoff = b'134 -1920 '
        def guard_release(payload):
            if payload == '10V00':
                self.assertEqual(lift.position, -1920)
                self.assertTrue(any(expected_dropoff in value for value in lift.writes))
        original_write = self.vacuum_serial.write
        def write(payload):
            guard_release(payload.decode('ascii').strip())
            return original_write(payload)
        with patch.object(self.vacuum_serial, 'write', side_effect=write):
            self.run_queue_command(self.backend.return_queue_specimen)
        self.assertFalse(self.vacuum.is_valve_connected())
        self.assertEqual(lift.position, 0)
        self.assertEqual(self.store.read()['stage']['record']['transfer_context']['original_slot'], 1)
        self.run_queue_command(self.backend.finish_queue_lifetime)
        self.assertEqual(self.store.read()['status'], 'verified')
        self.assertEqual(self.motor._connections, {})

    def test_missing_or_invalid_physical_slot_creates_no_queue_or_commands(self):
        original = self.store.read()
        before = list(self.commands), list(self.vacuum_serial.writes)
        for slot in (None, True, 0, 101, 2.5):
            with self.subTest(slot=slot), self.assertRaises((HardwareSafetyError, ValueError)):
                self.prepare(chain_slot=slot)
        self.assertEqual(self.store.read(), original)
        self.assertEqual((self.commands, self.vacuum_serial.writes), before)
        self.assertEqual(self.motor._connections, {})

    def test_chain_reference_is_journaled_before_relabel_and_stays_pending_on_failure(self):
        startup = self.prepare(chain_slot=99)
        connect = self.motor.connect
        def arm(*args, **kwargs):
            connect(*args, **kwargs)
            def observe(command):
                if command.startswith('11 10 '):
                    stage = self.store.read()['stage']
                    self.assertEqual(stage['status'], 'pending')
                    self.assertEqual(stage['plan']['attestation']['chain_slot'], 99)
                    raise OSError('injected relabel failure')
            self.chain_serial().on_write = observe
        with patch.object(self.motor, 'connect', side_effect=arm):
            worker = self.command(self.backend.start_queue_lifetime)
        self.assertFalse(worker.ok)
        self.assertIn('injected relabel failure', worker.error)
        self.assertEqual(self.store.read()['stage']['status'], 'pending')
        self.assertFalse(any(b' 134 ' in row or b' 135 ' in row for row in self.chain_serial().writes))
        self.assertFalse(startup.completed)

    def test_wrong_chain_slot_aborts_motion_without_replaying_during_shutdown(self):
        startup = self.start_chain()
        chain = self.chain_serial()
        chain.slop = round(abs(self.motor.config.one_step))  # Stops at a different registered slot.
        worker = self.command(lambda: self.backend.park_queue_station(10))
        self.assertFalse(worker.ok)
        state = self.store.read()
        self.assertEqual(state['stage']['status'], 'pending')
        self.assertEqual(state['stage']['plan']['slot'], 10)
        previous = len(chain.writes)
        operator = 'Recovery Operator'
        confirmation = RestartShutdownAttestation.confirm(state, operator,
            specimen_secured=True, empty_rod=True, work_area_clear=True)
        recovery = QueueRestartRecovery(self.backend, self.vacuum, state,
            operator=operator, shutdown=confirmation, live_startup=startup)
        recovery.run()
        self.assertFalse(any(b' 134 ' in row or b' 135 ' in row or b' 145' in row for row in chain.writes[previous:]))
        self.assertEqual(self.store.read()['status'], 'verified')
        self.assertEqual(self.store.read()['record']['settlement']['outcome'], 'aborted')

    def test_typed_chain_reference_does_not_survive_changed_connection(self):
        startup = self.start_chain()
        with startup.session.claim():
            reference = self.backend._queue_coordinator.reference.require(startup.session)
            before = len(self.chain_serial().writes)
            with patch.object(self.motor, '_connection_id', 'different'):
                with self.assertRaisesRegex(HardwareSafetyError, 'original chain reference|Original chain reference'):
                    self.backend._queue_coordinator.reference.require(startup.session)
            self.assertEqual(len(self.chain_serial().writes), before)
            self.assertEqual(reference.slot, 99)

    def misaligned_loaded_specimen(self):
        self.start_chain()
        self.run_queue_command(lambda: self.backend.load_queue_specimen(1, 'S1', file_id='F1'))
        geometry = self.backend._queue_coordinator.table.geometry
        self.chain_serial().position = geometry.chain_reference_counts(10, -.02)

    def test_wrong_empty_slot_blocks_chain_acquisition_before_instrument_io(self):
        from rapid_main.queue_acquisition import run_queue_acquisition
        self.misaligned_loaded_specimen()
        acquire = Mock()
        worker = self.command(lambda: run_queue_acquisition(self.backend, 'squid', acquire))
        self.assertFalse(worker.ok)
        self.assertIn('original calibrated empty hole', worker.error)
        acquire.assert_not_called()
        self.assertTrue(self.vacuum.is_valve_connected())
        self.assertEqual(self.store.read()['stage']['family'], 'acquisition')
        self.assertEqual(self.store.read()['stage']['status'], 'pending')

    def test_wrong_empty_slot_blocks_chain_treatment_before_field_io(self):
        self.misaligned_loaded_specimen()
        before = list(self.commands)
        with patch('rapid_main.af_treatment.AfTreatmentService') as service:
            worker = self.command(lambda: self.backend._execute_af_treatment('AF50'))
        self.assertFalse(worker.ok)
        self.assertIn('original calibrated empty hole', worker.error)
        service.return_value.execute.assert_not_called()
        self.assertEqual(self.commands, before)
        self.assertTrue(self.vacuum.is_valve_connected())
        self.assertEqual(self.store.read()['stage']['family'], 'af')
        self.assertEqual(self.store.read()['stage']['status'], 'pending')
