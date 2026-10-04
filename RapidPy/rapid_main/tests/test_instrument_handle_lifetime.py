"""Original SQUID/bridge handles survive failed close without replacement."""
import unittest
from types import SimpleNamespace
from unittest.mock import patch

from updown_control.app import RawSquidClient, SquidMomentReader, SquidCommunicationError
from rapid_main.diagnostic_services import SquidBackendAdapter
from rapid_main.susceptibility_transport import SusceptibilitySerialClient, SusceptibilityTransportConfig, SusceptibilityTransportError


class SerialHandle:
    def __init__(self):
        self.is_open = True
        self.fail = True
        self.ignore_close = False
        self.closes = 0

    def close(self):
        self.closes += 1
        if self.fail:
            raise OSError('injected close failure')
        if not self.ignore_close:
            self.is_open = False


class InstrumentHandleTests(unittest.TestCase):
    def test_bridge_failed_open_and_failed_cleanup_retains_handle_without_replacement(self):
        handle = SerialHandle()
        handle.is_open = False
        factory_calls = []
        def factory(**kwargs):
            factory_calls.append(kwargs)
            return handle
        client = SusceptibilitySerialClient(SusceptibilityTransportConfig(port='COM7'), serial_factory=factory)
        with self.assertRaisesRegex(SusceptibilityTransportError, 'original handle retained'):
            client.connect()
        self.assertIs(client._serial, handle)
        self.assertFalse(client.is_connected)
        with self.assertRaises(OSError):
            client.connect()
        self.assertEqual(len(factory_calls), 1)
        self.assertIs(client._serial, handle)
        handle.fail = False
        client.close()
        self.assertIsNone(client._serial)

    def test_bridge_failed_open_successful_cleanup_has_no_remaining_handle(self):
        handle = SerialHandle()
        handle.is_open, handle.fail = False, False
        client = SusceptibilitySerialClient(SusceptibilityTransportConfig(port='COM7'), serial_factory=lambda **kw: handle)
        with self.assertRaises(SusceptibilityTransportError):
            client.connect()
        self.assertIsNone(client._serial)
        self.assertEqual(handle.closes, 1)

    def test_squid_buffer_failure_and_failed_cleanup_blocks_reads_and_replacement(self):
        for phase in ('input', 'output'):
            with self.subTest(phase=phase):
                handle = SerialHandle()
                def fail(): raise OSError('buffer initialization failed')
                handle.reset_input_buffer = fail if phase == 'input' else lambda: None
                handle.reset_output_buffer = fail if phase == 'output' else lambda: None
                client = RawSquidClient()
                with patch('updown_control.app.serial.Serial', return_value=handle) as factory:
                    with self.assertRaisesRegex(SquidCommunicationError, 'original handle retained'):
                        client.connect('COM1')
                    self.assertIs(client._serial, handle)
                    self.assertFalse(client.is_connected)
                    with self.assertRaisesRegex(SquidCommunicationError, 'not connected'):
                        client.read_xyz_raw()
                    with self.assertRaises(OSError):
                        client.connect('COM1')
                    factory.assert_called_once()
                handle.fail = False
                client.disconnect()
                self.assertIsNone(client._serial)

    def test_squid_failed_open_closes_returned_handle_before_reporting_failure(self):
        handle = SerialHandle()
        handle.is_open, handle.fail = False, False
        client = RawSquidClient()
        with patch('updown_control.app.serial.Serial', return_value=handle):
            with self.assertRaises(SquidCommunicationError):
                client.connect('COM1')
        self.assertIsNone(client._serial)
        self.assertEqual(handle.closes, 1)

    def test_failed_diagnostic_baseline_cannot_reuse_previous_baseline(self):
        adapter = SquidBackendAdapter()
        def fail_read(): raise OSError('baseline failed')
        def fail_close(): raise OSError('close failed')
        reader = SimpleNamespace(is_connected=True, raw_client=object(), take_baseline=fail_read, disconnect=fail_close)
        adapter._reader, adapter._baseline_raw = reader, (1., 2., 3.)
        with self.assertRaises(OSError):
            adapter.test_connection()
        self.assertIs(adapter._reader, reader)
        self.assertIsNone(adapter._baseline_raw)
        with self.assertRaisesRegex(RuntimeError, 'baseline'):
            adapter.read_squid()

    def clients(self):
        return (RawSquidClient(), SusceptibilitySerialClient(SusceptibilityTransportConfig(port='COM7')))

    def test_failed_close_retains_exact_handle_and_retry_closes_only_that_handle(self):
        for client in self.clients():
            with self.subTest(client=type(client).__name__):
                handle = SerialHandle()
                client._serial = handle
                close = getattr(client, 'disconnect', None) or client.close
                with self.assertRaisesRegex(OSError, 'close failure'):
                    close()
                self.assertIs(client._serial, handle)
                handle.fail = False
                close()
                self.assertIsNone(client._serial)
                self.assertEqual(handle.closes, 2)

    def test_returning_without_closing_does_not_forget_handle(self):
        for client in self.clients():
            with self.subTest(client=type(client).__name__):
                handle = SerialHandle()
                handle.fail, handle.ignore_close = False, True
                client._serial = handle
                close = getattr(client, 'disconnect', None) or client.close
                with self.assertRaisesRegex(RuntimeError, 'remains open'):
                    close()
                self.assertIs(client._serial, handle)

    def test_adapter_retains_original_reader_and_baseline_on_failed_close(self):
        adapter = SquidBackendAdapter()
        reader = SquidMomentReader()
        handle = SerialHandle()
        reader.raw_client._serial = handle
        adapter._reader, adapter._baseline_raw = reader, (1., 2., 3.)
        with self.assertRaises(OSError):
            adapter.disconnect()
        self.assertIs(adapter._reader, reader)
        self.assertIs(adapter.raw_client._serial, handle)
        self.assertEqual(adapter._baseline_raw, (1., 2., 3.))
        handle.fail = False
        adapter.disconnect()
        self.assertIsNone(adapter._reader)

    def test_preparation_has_no_io_and_acquisition_connection_has_no_baseline(self):
        reader = SimpleNamespace(is_connected=False, raw_client=object())
        calls = []
        def connect(*args, **kwargs):
            calls.append((args, kwargs))
            reader.is_connected = True
        reader.connect = connect
        reader.take_baseline = lambda: self.fail('baseline belongs to diagnostic connection only')
        with patch('rapid_main.diagnostic_services.SquidMomentReader', return_value=reader) as factory:
            adapter = SquidBackendAdapter()
            original = adapter.prepare_raw_client()
            self.assertIs(adapter.prepare_raw_client(), original)
            self.assertEqual(calls, [])
            self.assertFalse(adapter.is_connected())
            adapter.connect_for_acquisition()
            adapter.connect_for_acquisition()
            self.assertIs(adapter.raw_client, original)
            self.assertEqual(len(calls), 1)
            factory.assert_called_once_with()
