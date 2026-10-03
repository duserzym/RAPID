"""Ownership and crash recovery for standalone ADwin field diagnostics."""
from contextlib import contextmanager
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


def publish_diagnostic_record(store, record):
    if not re.fullmatch(r"af-[0-9a-f]{32}", record.treatment_id):
        raise HardwareSafetyError("Invalid AF diagnostic evidence identity.")
    folder = store.path.parent / "hardware_diagnostics" / record.treatment_id
    try:
        folder.mkdir(parents=True, exist_ok=False)
        path = folder / "record.json"
        payload = (json.dumps(record.to_dict(), indent=2, sort_keys=True, allow_nan=False) + "\n").encode("utf-8")
        index = {"schema": "rapidpy.af.diagnostic_artifact_index.v1",
                 "state": "verified" if record.safe_state_confirmed else "unsafe",
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
