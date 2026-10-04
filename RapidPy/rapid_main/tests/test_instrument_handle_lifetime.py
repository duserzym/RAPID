"""Original SQUID/bridge handles survive failed close without replacement."""
import unittest
from types import SimpleNamespace
from unittest.mock import patch

from updown_control.app import RawSquidClient, SquidMomentReader
from rapid_main.diagnostic_services import SquidBackendAdapter
from rapid_main.susceptibility_transport import SusceptibilitySerialClient, SusceptibilityTransportConfig


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
