"""Original scientific adapters and serial handles owned by one native queue."""
import copy
from dataclasses import asdict
import uuid

from rapidpy_common.hardware_safety import HardwareSafetyError


class QueueInstrumentLifetime:
    @staticmethod
    def port_key(value):
        return str(value).strip().upper().lstrip(chr(92) + '.')

    def __init__(self, backend):
        from .diagnostic_services import SquidBackendAdapter, SusceptibilityBackendAdapter, UnavailableBackend
        from .susceptibility_transport import SusceptibilitySerialClient
        from updown_control.app import RawSquidClient, SquidMomentReader
        self.backend, self.session = backend, None
        self.squid = backend._measurement
        self.bridge = backend._susceptibility
        config = backend._config
        if (not isinstance(self.squid, SquidBackendAdapter) or self.squid.simulated is not False
                or asdict(self.squid._cfg) != asdict(config.squid)):
            raise HardwareSafetyError('The original native SQUID adapter and accepted settings are required.')
        self.squid.prepare_raw_client()  # Object creation only; no port/reset/read.
        self.reader, self.raw = self.squid._reader, self.squid.raw_client
        if (not isinstance(self.reader, SquidMomentReader) or not isinstance(self.raw, RawSquidClient)
                or self.raw._serial is not None):
            raise HardwareSafetyError('Settle retained SQUID handles before preparing a new native queue.')
        self.bridge_client = None
        absent_bridge = self.bridge is None or (isinstance(self.bridge, UnavailableBackend)
            and config.susceptibility.enabled is False)
        if not absent_bridge:
            if (not isinstance(self.bridge, SusceptibilityBackendAdapter) or self.bridge.simulated is not False
                    or asdict(self.bridge._cfg) != asdict(config.susceptibility)
                    or not isinstance(self.bridge._client, SusceptibilitySerialClient)):
                raise HardwareSafetyError('The original native susceptibility adapter and accepted settings are required.')
            self.bridge_client = self.bridge._client
            expected = dict(port=config.susceptibility.port, baud=config.susceptibility.baud,
                parity=config.susceptibility.parity, bytesize=config.susceptibility.bytesize,
                stopbits=config.susceptibility.stopbits, response_timeout_s=config.susceptibility.response_timeout,
                scale_factor=config.susceptibility.scale_factor)
            if any(getattr(self.bridge_client.config, key) != value for key, value in expected.items()):
                raise HardwareSafetyError('Actual susceptibility client settings differ from the accepted adapter.')
            if self.bridge_client._serial is not None:
                raise HardwareSafetyError('Settle retained susceptibility handles before preparing a new native queue.')
        elif config.susceptibility.enabled is not False:
            raise HardwareSafetyError('The enabled susceptibility instrument must have its original native adapter.')
        self.profile = dict(schema='rapidpy.queue_instruments.v1', instance_id=uuid.uuid4().hex,
            squid=asdict(config.squid), susceptibility=asdict(config.susceptibility),
            bridge_present=self.bridge_client is not None)
        self.handles = dict(squid=None, susceptibility=None)
        self.closed = dict(squid=False, susceptibility=self.bridge_client is None)
        self.closing = False
        ports = [config.squid.port] + ([config.susceptibility.port] if self.bridge_client is not None else [])
        keys = [self.port_key(port) for port in ports]
        if any(not key for key in keys) or len(set(keys)) != len(keys):
            raise HardwareSafetyError('Scientific instruments require distinct configured serial ports.')

    def bind(self, session):
        session.child_store._owned()
        if self.session is not None and self.session is not session:
            raise HardwareSafetyError('Another queue already owns the original scientific instruments.')
        self.session = session
        self.validate()

    def _objects(self, *, require_owner=True):
        config = self.backend._config
        owner = getattr(self.backend, '_queue_instruments', None)
        if ((owner is not self if require_owner else owner is not None and owner is not self)
                or self.backend._measurement is not self.squid or self.backend._susceptibility is not self.bridge
                or self.reader.raw_client is not self.raw
                or asdict(config.squid) != self.profile['squid']
                or asdict(config.susceptibility) != self.profile['susceptibility']
                or asdict(self.squid._cfg) != self.profile['squid']
                or (not self.closed['squid'] and self.squid._reader is not self.reader)
                or (self.closed['squid'] and self.squid._reader is not None)):
            raise HardwareSafetyError('Restore the original queue scientific SQUID/bridge objects and accepted settings.')
        if self.bridge_client is not None:
            cfg = self.profile['susceptibility']
            expected = dict(port=cfg['port'], baud=cfg['baud'], parity=cfg['parity'], bytesize=cfg['bytesize'],
                stopbits=cfg['stopbits'], response_timeout_s=cfg['response_timeout'], scale_factor=cfg['scale_factor'])
            if (self.bridge._client is not self.bridge_client or asdict(self.bridge._cfg) != cfg
                    or any(getattr(self.bridge_client.config, key) != value for key, value in expected.items())):
                raise HardwareSafetyError('Restore the original susceptibility client and serial configuration.')

    def _client(self, name):
        return self.raw if name == 'squid' else self.bridge_client

    def validate(self, *, allow_verified=False):
        if self.session is None:
            raise HardwareSafetyError('Bind original instruments to their durable queue before I/O.')
        self.session.child_store._owned()
        state = self.session.store.read()
        if allow_verified and state and state['status'] == 'verified':
            if (state['family'] != 'queue' or state['token'] != self.session.token
                    or state['record'].get('settlement', {}).get('instruments', {}).get('profile') != self.profile):
                raise HardwareSafetyError('Exact original instrument settlement record is required.')
            root = state
        else:
            root = self.session.store._queue(self.session.token)
        self.session.store.verify_history(root)
        if root['profile']['resources'].get('instruments') != self.profile or self.session.is_recovery:
            raise HardwareSafetyError('Original live instrument ownership and durable profile are required.')
        self._objects()
        stage = root['stage']
        for name in self.handles:
            client = self._client(name)
            handle = client._serial if client is not None else None
            if handle is not None and self.handles[name] is None:
                action = 'squid' if name == 'squid' else 'susceptibility'
                if (self.closing or self.closed[name] or not stage or stage['family'] != 'acquisition'
                        or stage['status'] != 'pending' or stage['token'] != self.backend._geometry_stage_token
                        or stage['plan']['action'] not in ({action, 'flux_recovery'} if name == 'squid' else {action})):
                    raise HardwareSafetyError('A new scientific handle requires its original pending acquisition stage.')
                self.handles[name] = handle  # Retain even if its settings/readback fail.
            if handle is not (None if self.closed[name] else self.handles[name]):
                raise HardwareSafetyError('Original queue scientific serial handle was lost or replaced.')
            if handle is not None:
                cfg = self.profile[name if name == 'squid' else 'susceptibility']
                expected = dict(port=cfg['port'], baudrate=cfg['baud'], bytesize=8 if name == 'squid' else cfg['bytesize'],
                    parity='N' if name == 'squid' else cfg['parity'].upper(), stopbits=1 if name == 'squid' else cfg['stopbits'])
                if any(getattr(handle, key, None) != value for key, value in expected.items()):
                    raise HardwareSafetyError('Actual scientific serial settings differ from the original queue profile.')
                if not self.closing and getattr(handle, 'is_open', None) is not True:
                    raise HardwareSafetyError('Original scientific serial connection is unverified.')
        return root

    def settlement_plan(self):
        self.validate()
        self.closing = True
        return dict(profile=copy.deepcopy(self.profile),
            connected={name: handle is not None for name, handle in self.handles.items()})

    def close(self, name):
        root = self.validate()
        stage = root['stage']
        if (not self.closing or not stage or stage['status'] != 'pending'
                or stage['plan'].get('action') != 'queue_terminal_close'
                or stage['plan'].get('instruments', {}).get('profile') != self.profile):
            raise HardwareSafetyError('Original pending terminal close stage is required to settle instruments.')
        if self.closed[name]:
            return
        adapter = self.squid if name == 'squid' else self.bridge
        adapter.disconnect()
        client = self._client(name)
        handle = self.handles[name]
        if client._serial is not None or (handle is not None and getattr(handle, 'is_open', None) is not False):
            raise HardwareSafetyError('Original scientific serial handle has not settled.')
        self.closed[name] = True

    def is_settled(self):
        """Read-only live object checks for GUI release after actual worker exit."""
        try:
            self._objects(require_owner=False)
            return (self.closing and all(self.closed.values())
                and all(self._client(name) is None or self._client(name)._serial is None for name in self.closed)
                and all(handle is None or getattr(handle, 'is_open', None) is False for handle in self.handles.values()))
        except Exception:
            return False

    def is_original(self):
        """Read-only object/handle check before GUI grants original queue reentry."""
        try:
            self._objects()
            return (not self.closing and not self.closed['squid']
                and all(self._client(name) is None or self._client(name)._serial is self.handles[name]
                        for name in self.handles)
                and all(handle is None or getattr(handle, 'is_open', None) is True for handle in self.handles.values()))
        except Exception:
            return False

    def require_settled(self):
        self.validate(allow_verified=True)
        if not self.is_settled():
            raise HardwareSafetyError('Both original scientific instruments must settle before queue publication/release.')
