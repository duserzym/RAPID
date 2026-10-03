"""Ownership and crash recovery for standalone ADwin field diagnostics."""
from contextlib import contextmanager
import copy
from dataclasses import asdict, dataclass
from datetime import datetime, timezone
import hashlib
import json
import os
from pathlib import Path
import re
import uuid

from .hardware_safety import HardwareSafetyError, HardwareSafetyStore, default_safety_path


def station_profile(controller):
    return {"adwin_board": asdict(controller.board), "limits": asdict(controller.limits)}


@dataclass(frozen=True)
class AfDiagnosticRecord:
    treatment_id: str
    sample_id: str
    run_id: str
    operation: dict
    timestamp_iso: str
    error: str
    cleanup_error: str
    safe_state_confirmed: bool
    observations: tuple = ()
    simulated: bool = False
    schema: str = "rapidpy.af.diagnostic.v1"

    def to_dict(self):
        return asdict(self)


def _record(operation, error, cleanup_error, *, recovery=False, observations=()):
    return AfDiagnosticRecord("af-" + uuid.uuid4().hex, "diagnostic", "", operation,
                              datetime.now(timezone.utc).isoformat(), error, cleanup_error,
                              not cleanup_error, observations=tuple(observations),
                              schema="rapidpy.af.diagnostic_recovery.v1" if recovery else "rapidpy.af.diagnostic.v1")


def publish_diagnostic_record(store, record, *, family="af"):
    if family not in {"af", "motion", "station"} or not re.fullmatch(family + r"-[0-9a-f]{32}", record.treatment_id):
        raise HardwareSafetyError("Invalid AF diagnostic evidence identity.")
    folder = store.path.parent / "hardware_diagnostics" / record.treatment_id
    try:
        folder.mkdir(parents=True, exist_ok=False)
        path = folder / "record.json"
        payload = (json.dumps(record.to_dict(), indent=2, sort_keys=True, allow_nan=False) + "\n").encode("utf-8")
        held = family == 'station' and record.outputs_held and not record.error and not record.cleanup_error
        index = {"schema": f"rapidpy.{family}.diagnostic_artifact_index.v1",
                 "state": "held" if held else ("verified" if record.safe_state_confirmed else "unsafe"),
                 "artifacts": [{"relative_path": path.name, "sha256": hashlib.sha256(payload).hexdigest(), "required": True}]}
        for name, content in (("record.json", payload), ("artifact_index.json", (json.dumps(index, indent=2) + "\n").encode("utf-8"))):
            with (folder / name).open("xb") as handle:
                handle.write(content)
                handle.flush()
                os.fsync(handle.fileno())
    except (OSError, TypeError, ValueError) as exc:
        raise HardwareSafetyError(f"Cannot publish immutable AF diagnostic evidence: {exc}") from exc
    return folder


@contextmanager
def af_diagnostic_operation(controller, operation, *, store=None):
    """Journal before field I/O; always stop/zero/read back before release."""
    store = store if store is not None else HardwareSafetyStore(default_safety_path())
    profile = station_profile(controller)
    with store.operation_lease():
        token = store.begin("af_diagnostic", operation, profile, sample_id="diagnostic")
        error = ""
        observations = []
        try:
            yield observations
        except BaseException as exc:
            error = f"{type(exc).__name__}: {exc}"
            raise
        finally:
            cleanup_error = ""
            try:
                controller.recover_safe_field()
            except Exception as exc:
                cleanup_error = str(exc)
            record = _record(operation, error, cleanup_error, observations=observations)
            publish_diagnostic_record(store, record)
            store.finish(token, profile, record)
            if cleanup_error:
                raise HardwareSafetyError(f"AF diagnostic output recovery failed: {cleanup_error}")


def recover_af_diagnostic(controller, *, store=None):
    """Recover only this helper's original board; never boot/start a process."""
    store = store if store is not None else HardwareSafetyStore(default_safety_path())
    profile = station_profile(controller)
    with store.operation_lease():
        pending = store.pending()
        if pending is not None and pending["family"] != "af_diagnostic":
            raise HardwareSafetyError("Recover the unfinished main-app treatment from the main app.")
        if pending is not None:
            pending = store.pending(profile)
        error = ""
        try:
            controller.recover_safe_field()
        except Exception as exc:
            error = str(exc)
        record = _record(pending["plan"] if pending else {"action": "verify_outputs_off"}, "", error, recovery=True)
        publish_diagnostic_record(store, record)
        if pending:
            store.finish(pending["token"], profile, record)
        if error:
            raise HardwareSafetyError(f"AF diagnostic recovery failed: {error}")
        return record


class ManualAdwinDiagnostic:
    """Keep manual output functionality under one durable ownership lifetime."""
    def __init__(self, controller, *, store=None):
        self.controller = controller
        self.store = store if store is not None else HardwareSafetyStore(default_safety_path())
        self._profile = station_profile(controller)
        self._board, self._limits = copy.deepcopy(controller.board), copy.deepcopy(controller.limits)
        self._manager = None
        self._observations = None

    @property
    def active(self):
        return self._manager is not None

    def write(self, request, action):
        # Strict JSON validation precedes lease acquisition and physical I/O.
        snapshot = json.loads(json.dumps(request, allow_nan=False))
        if station_profile(self.controller) != self._profile:
            raise HardwareSafetyError("Manual diagnostic board settings changed; recover the original board before further output.")
        starting = self._manager is None
        if starting:
            manager = af_diagnostic_operation(self.controller, {"action": "manual_output", "initial_request": snapshot}, store=self.store)
            observations = manager.__enter__()
            self._manager, self._observations = manager, observations
        try:
            if starting:
                self.controller.recover_safe_field()
            result = action()
            self._observations.append({"request": snapshot, "result": result})
            return result
        except BaseException as exc:
            self.close(type(exc), exc, exc.__traceback__)
            raise

    def observe(self, observation):
        if self._observations is not None:
            snapshot = json.loads(json.dumps(observation, allow_nan=False))
            self._observations.append(snapshot)

    def close(self, exc_type=None, exc=None, traceback=None):
        manager = self._manager
        if manager is None:
            return
        # Cleanup always addresses the original station even if an embedding
        # client has incorrectly mutated its controller in the meantime.
        self.controller.board, self.controller.limits = copy.deepcopy(self._board), copy.deepcopy(self._limits)
        try:
            manager.__exit__(exc_type, exc, traceback)
        finally:
            self._manager, self._observations = None, None
