"""Original native scientific adapters with injected serial handles, no readings."""
from dataclasses import replace
import unittest
from unittest.mock import patch

from rapid_main.diagnostic_services import SusceptibilityBackendAdapter, UnavailableBackend
from rapid_main.queue_instruments import QueueInstrumentLifetime
from rapid_main.queue_table_motion import QueueTableMoveRecord
from rapidpy_common.hardware_safety import HardwareSafetyError
from tests import test_queue_backend_coordinator as fixture


class ScientificSerial:
    def __init__(self, owner, name):
        self.owner, self.name = owner, name
        config = owner.backend._config
        cfg = config.squid if name == 'squid' else config.susceptibility
        self.port, self.baudrate = cfg.port, cfg.baud
        self.parity, self.bytesize, self.stopbits = ('N', 8, 1) if name == 'squid' else (cfg.parity, cfg.bytesize, cfg.stopbits)
        self.is_open, self.fail_close, self.ignore_close = True, False, False
        self.closes = 0

    def close(self):
        stage = self.owner.session.store.pending()['stage']
        if stage['status'] != 'pending' or stage['plan']['action'] != 'queue_terminal_close':
            raise AssertionError('Instrument close was not journaled before I/O')
        self.closes += 1
        if self.fail_close:
            raise OSError('injected ' + self.name + ' close failure')
        if not self.ignore_close:
            self.is_open = False


class QueueInstrumentTests(fixture.QueueBackendCoordinatorFixture, unittest.TestCase):
    def test_explicitly_disabled_bridge_placeholder_has_no_physical_handle(self):
        self.backend._config.susceptibility.enabled = False
        self.backend._susceptibility = UnavailableBackend('Bridge', 'disabled')
        prepared = QueueInstrumentLifetime(self.backend)
        self.assertFalse(prepared.profile['bridge_present'])
        self.assertIsNone(prepared.bridge_client)

    def test_enabled_bridge_placeholder_cannot_prepare_native_owner(self):
        self.backend._susceptibility = UnavailableBackend('Bridge', 'missing')
        with self.assertRaisesRegex(HardwareSafetyError, 'native susceptibility'):
            QueueInstrumentLifetime(self.backend)

    def test_replaced_raw_client_is_rejected_before_native_io(self):
        original = self.instruments.reader._client
        self.instruments.reader._client = object()
        try:
            before = len(self.commands), len(self.vacuum_serial.writes)
            worker = self.command(lambda: self.backend.load_queue_specimen(1, 'S1', file_id='F1'))
            self.assertFalse(worker.ok)
            self.assertEqual((len(self.commands), len(self.vacuum_serial.writes)), before)
        finally:
            self.instruments.reader._client = original

    def test_scientific_port_collision_is_rejected_before_open(self):
        self.backend._config.susceptibility.port = self.backend._config.squid.port.lower()
        # Rebuild only the idle adapter to match the deliberately bad settings.
        self.backend._susceptibility = SusceptibilityBackendAdapter(self.backend._config.susceptibility)
        with self.assertRaisesRegex(HardwareSafetyError, 'distinct configured serial'):
            QueueInstrumentLifetime(self.backend)

    def additional_stage_profiles(self):
        profiles = super().additional_stage_profiles()
        cfg = self.backend._config.susceptibility
        cfg.enabled, cfg.port = True, 'COM7'
        self.backend._susceptibility = SusceptibilityBackendAdapter(cfg)
        self.backend._queue_instruments = QueueInstrumentLifetime(self.backend)
        profiles['acquisition'] = self.backend._acquisition_safety_profile()
        return profiles

    def setUp(self):
        super().setUp()
        self.instruments = self.backend._queue_instruments

    def attach(self, name):
        # Connection-only fixture stage, not scientific specimen acceptance.
        with self.session.claim():
            child = self.session.child_store
            operation = dict(action='squid' if name == 'squid' else 'susceptibility')
            token = child.begin('acquisition', operation, self.backend._acquisition_safety_profile())
            self.backend._geometry_stage_token = token
            handle = ScientificSerial(self.instruments, name)
            self.instruments._client(name)._serial = handle
            self.instruments.validate()
            record = QueueTableMoveRecord(operation, self.table.profile, (), (), '', '', True,
                schema='rapidpy.test_connection_only.v1')
            child.finish(token, self.backend._acquisition_safety_profile(), record)
            self.backend._geometry_stage_token = None
            return handle

    def finish(self):
        return self.command(self.backend.finish_queue_lifetime)

    def test_both_original_instruments_close_before_root_and_owner_detach(self):
        squid, bridge = self.attach('squid'), self.attach('susceptibility')
        worker = self.finish()
        self.assertTrue(worker.ok, worker.error)
        self.assertEqual((squid.closes, bridge.closes), (1, 1))
        self.assertFalse(squid.is_open or bridge.is_open)
        self.assertTrue(self.instruments.is_settled())
        root = self.store.read()
        self.assertEqual(root['status'], 'verified')
        self.assertEqual(root['record']['settlement']['instruments']['profile'], self.instruments.profile)
        self.assertEqual(root['record']['settlement']['instruments']['connected'], {'squid': True, 'susceptibility': True})
        self.assertIsNotNone(self.session._lease)  # Actual worker exit still precedes caller release.

    def test_failed_squid_close_still_settles_bridge_and_retries_only_original_handle(self):
        squid, bridge = self.attach('squid'), self.attach('susceptibility')
        squid.fail_close = True
        worker = self.finish()
        self.assertFalse(worker.ok)
        self.assertIs(self.instruments.raw._serial, squid)
        self.assertEqual((squid.closes, bridge.closes), (1, 1))
        self.assertIsNone(self.instruments.bridge_client._serial)
        self.assertIsNotNone(self.session._lease)
        token = self.store.pending()['stage']['token']
        before = len(self.commands), len(self.vacuum_serial.writes)
        squid.fail_close = False
        worker = self.finish()
        self.assertTrue(worker.ok, worker.error)
        self.assertEqual((squid.closes, bridge.closes), (2, 1))
        self.assertEqual((len(self.commands), len(self.vacuum_serial.writes)), before)
        self.assertEqual(self.store.read()['stage']['token'], token)

    def test_failed_bridge_close_does_not_reopen_settled_squid(self):
        squid, bridge = self.attach('squid'), self.attach('susceptibility')
        bridge.fail_close = True
        self.assertFalse(self.finish().ok)
        self.assertIsNone(self.instruments.squid._reader)
        self.assertIs(self.instruments.bridge_client._serial, bridge)
        bridge.fail_close = False
        worker = self.finish()
        self.assertTrue(worker.ok, worker.error)
        self.assertEqual((squid.closes, bridge.closes), (1, 2))

    def test_close_returning_with_open_port_retains_pending_queue(self):
        squid = self.attach('squid')
        squid.ignore_close = True
        worker = self.finish()
        self.assertFalse(worker.ok)
        self.assertIs(self.instruments.raw._serial, squid)
        self.assertEqual(self.store.pending()['stage']['status'], 'pending')
        self.assertFalse(self.instruments.is_settled())

    def test_replaced_serial_on_close_retry_is_rejected_without_any_close(self):
        squid = self.attach('squid')
        squid.fail_close = True
        self.assertFalse(self.finish().ok)
        replacement = ScientificSerial(self.instruments, 'squid')
        self.instruments.raw._serial = replacement
        before = squid.closes, replacement.closes, len(self.commands)
        worker = self.finish()
        self.assertFalse(worker.ok)
        self.assertIn('lost or replaced', worker.error)
        self.assertEqual((squid.closes, replacement.closes, len(self.commands)), before)

    def test_changed_client_config_blocks_retry_before_io(self):
        bridge = self.attach('susceptibility')
        bridge.fail_close = True
        self.assertFalse(self.finish().ok)
        cfg = self.instruments.bridge_client.config
        self.instruments.bridge_client.config = replace(cfg, baud=cfg.baud * 2)
        before = bridge.closes, len(self.commands)
        worker = self.finish()
        self.assertFalse(worker.ok)
        self.assertEqual((bridge.closes, len(self.commands)), before)

    def test_replaced_adapter_blocks_transfer_before_cutoff_or_motion(self):
        self.backend._measurement = object()
        before = len(self.commands), len(self.vacuum_serial.writes)
        worker = self.command(lambda: self.backend.load_queue_specimen(1, 'S1', file_id='F1'))
        self.assertFalse(worker.ok)
        self.assertEqual((len(self.commands), len(self.vacuum_serial.writes)), before)

    def test_unstaged_connection_cannot_be_adopted_by_original_queue(self):
        self.instruments.raw._serial = ScientificSerial(self.instruments, 'squid')
        with self.session.claim(), self.assertRaisesRegex(HardwareSafetyError, 'pending acquisition'):
            self.instruments.validate()

    def test_wrong_serial_settings_are_retained_but_never_accepted(self):
        with self.session.claim():
            token = self.session.child_store.begin('acquisition', {'action': 'squid'}, self.backend._acquisition_safety_profile())
            self.backend._geometry_stage_token = token
            squid = ScientificSerial(self.instruments, 'squid')
            squid.baudrate *= 2
            self.instruments.raw._serial = squid
            with self.assertRaisesRegex(HardwareSafetyError, 'serial settings'):
                self.instruments.validate()
            self.assertIs(self.instruments.handles['squid'], squid)
            self.assertIs(self.instruments.raw._serial, squid)
            self.assertEqual(squid.closes, 0)

    def test_root_publication_retry_does_not_repeat_scientific_closes(self):
        squid, bridge = self.attach('squid'), self.attach('susceptibility')
        with patch.object(self.store, 'finish_queue', side_effect=OSError('publication failed')):
            self.assertFalse(self.finish().ok)
        self.assertEqual(self.store.pending()['stage']['status'], 'verified')
        worker = self.finish()
        self.assertTrue(worker.ok, worker.error)
        self.assertEqual((squid.closes, bridge.closes), (1, 1))
