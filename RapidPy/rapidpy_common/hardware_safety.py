"""Durable treatment latch; an interrupted operation must survive restart.

The file is independent of measurement/configuration destinations. A short
OS file lock serializes transactions across processes; hardware waits never
hold that lock. Only the operation token and matching station profile can
publish verified completion. Corrupt state is an error, never an empty latch.
"""
from contextlib import contextmanager
from datetime import datetime, timezone
import hashlib
import json
import os
from pathlib import Path
import uuid


class HardwareSafetyError(RuntimeError):
    pass


def _canonical(value):
    return json.dumps(value, sort_keys=True, separators=(",", ":"), allow_nan=False).encode("utf-8")


def default_safety_path():
    # Explicit override is for isolated test/deployment environments, not an
    # operator setting. Changing RAPID_CONFIG or data_dir cannot reset a latch.
    override = os.environ.get("RAPID_SAFETY_STATE")
    return Path(override) if override else Path.home() / ".rapid" / "hardware_safety.json"


class HardwareSafetyStore:
    schema = "rapidpy.hardware_safety.v1"

    def __init__(self, path):
        self.path = Path(path)

    @contextmanager
    def _locked(self, suffix=".lock"):
        try:
            self.path.parent.mkdir(parents=True, exist_ok=True)
            handle = self.path.with_suffix(self.path.suffix + suffix).open("a+b")
        except OSError as exc:
            raise HardwareSafetyError(f"Cannot access hardware safety transaction lock: {exc}") from exc
        with handle:
            handle.seek(0, os.SEEK_END)
            if handle.tell() == 0:
                handle.write(b"\0")
                handle.flush()
            handle.seek(0)
            try:
                if os.name == "nt":
                    import msvcrt
                    msvcrt.locking(handle.fileno(), msvcrt.LK_NBLCK, 1)
                else:
                    import fcntl
                    fcntl.flock(handle.fileno(), fcntl.LOCK_EX | fcntl.LOCK_NB)
            except OSError as exc:
                raise HardwareSafetyError("Another process is updating the hardware safety latch.") from exc
            try:
                yield
            finally:
                handle.seek(0)
                if os.name == "nt":
                    msvcrt.locking(handle.fileno(), msvcrt.LK_UNLCK, 1)
                else:
                    fcntl.flock(handle.fileno(), fcntl.LOCK_UN)

    @contextmanager
    def operation_lease(self):
        """OS-owned lifetime lease, automatically released after a crash.

        A second process may inspect the durable latch, but cannot recover
        while its owner is still treating/recovering the instrument. This is
        separate from the short transaction lock used by begin/finish.
        """
        with self._locked(".operation.lock"):
            yield

    def read(self):
        try:
            with self.path.open("r", encoding="utf-8") as handle:
                envelope = json.load(handle, parse_constant=lambda value: (_ for _ in ()).throw(ValueError(value)))
            state = envelope["state"]
            if envelope["sha256"] != hashlib.sha256(_canonical(state)).hexdigest():
                raise ValueError("checksum mismatch")
            if (state["schema"] != self.schema or state["status"] not in {"pending", "verified"}
                    or state["family"] not in {"pulse", "rrm", "af", "arm", "af_diagnostic", "motion_diagnostic", "station_diagnostic"}
                    or not isinstance(state["token"], str) or len(state["token"]) != 32
                    or not isinstance(state["profile"], dict) or not isinstance(state["plan"], dict)):
                raise ValueError("unsupported or incomplete state")
            int(state["token"], 16)
            if not isinstance(state["sample_id"], str) or not isinstance(state["run_id"], str):
                raise ValueError("invalid specimen/run identity")
            datetime.fromisoformat(state["started_at"])
            if state["status"] == "verified" and (not isinstance(state["record"], dict)
                    or state["record"].get("safe_state_confirmed") is not True
                    or state["record"].get("simulated") is not False):
                raise ValueError("verified state lacks physical safe-state evidence")
            return state
        except FileNotFoundError:
            return None
        except (OSError, ValueError, TypeError, KeyError) as exc:
            raise HardwareSafetyError(f"Hardware safety state cannot be verified ({self.path}): {exc}") from exc

    def _write(self, state):
        temporary = self.path.with_name(self.path.name + ".tmp-" + uuid.uuid4().hex)
        payload = {"state": state, "sha256": hashlib.sha256(_canonical(state)).hexdigest()}
        try:
            with temporary.open("x", encoding="utf-8") as handle:
                json.dump(payload, handle, sort_keys=True, indent=2, allow_nan=False)
                handle.write("\n")
                handle.flush()
                os.fsync(handle.fileno())
            os.replace(temporary, self.path)
            if os.name != "nt":
                descriptor = os.open(self.path.parent, os.O_RDONLY)
                try:
                    os.fsync(descriptor)
                finally:
                    os.close(descriptor)
        except (OSError, ValueError, TypeError) as exc:
            raise HardwareSafetyError(f"Cannot persist hardware safety state: {exc}") from exc
        finally:
            if temporary.exists():
                temporary.unlink()

    def begin(self, family, plan, profile, *, sample_id="", run_id=""):
        if family not in {"pulse", "rrm", "af", "arm", "af_diagnostic", "motion_diagnostic", "station_diagnostic"}:
            raise HardwareSafetyError("Unsupported treatment safety family.")
        # Round-trip makes a detached, strict JSON snapshot before file I/O.
        plan, profile = json.loads(_canonical(plan)), json.loads(_canonical(profile))
        with self._locked():
            prior = self.read()
            if prior and prior["status"] == "pending":
                raise HardwareSafetyError("An unfinished hardware treatment requires verified recovery.")
            token = uuid.uuid4().hex
            self._write(dict(schema=self.schema, status="pending", family=family, token=token,
                             plan=plan, profile=profile, sample_id=sample_id, run_id=run_id,
                             started_at=datetime.now(timezone.utc).isoformat(), record=None))
            return token

    def finish(self, token, profile, record):
        snapshot = record.to_dict()
        with self._locked():
            state = self.read()
            if not state or state["status"] != "pending" or state["token"] != token:
                raise HardwareSafetyError("Stale hardware safety operation token.")
            if _canonical(state["profile"]) != _canonical(profile):
                raise HardwareSafetyError("Station wiring/calibration changed; the original station must be recovered.")
            state["record"] = snapshot
            # Strict booleans prevent truthy strings/mocks from clearing it.
            if record.safe_state_confirmed is True and record.simulated is False:
                state["status"] = "verified"
            state["updated_at"] = datetime.now(timezone.utc).isoformat()
            self._write(state)

    def join_diagnostic_resource(self, token, name, binding):
        """Persist a newly participating station resource before its first I/O.

        Existing bindings can never be replaced, including after a crash. This
        is limited to the held-vacuum/lift station lifecycle, not treatments.
        """
        if name not in {'lift', 'vacuum'}:
            raise HardwareSafetyError('Unsupported station diagnostic resource.')
        binding = json.loads(_canonical(binding))
        with self._locked():
            state = self.read()
            if not state or state['status'] != 'pending' or state['token'] != token or state['family'] != 'station_diagnostic':
                raise HardwareSafetyError('Stale station diagnostic resource token.')
            resources = state['profile'].get('resources')
            if not isinstance(resources, dict):
                raise HardwareSafetyError('Station diagnostic resource profile is invalid.')
            if name in resources and _canonical(resources[name]) != _canonical(binding):
                raise HardwareSafetyError('The original station resource binding cannot be changed while outputs are held.')
            resources[name] = binding
            self._write(state)
            return json.loads(_canonical(state['profile']))

    def pending(self, profile=None):
        state = self.read()
        if not state or state["status"] == "verified":
            return None
        if profile is not None and _canonical(state["profile"]) != _canonical(profile):
            raise HardwareSafetyError("Station wiring/calibration changed; restore the latched station profile before recovery.")
        return state
