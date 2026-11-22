"""Per-index queue defaults in the operator's application settings."""
from dataclasses import asdict, dataclass
import hashlib
import json
import os
from pathlib import Path

from PySide6 import QtCore

from rapid_main.queue_compiler import parse_queue_sequence, validate_average_count


_GROUP = "sample-index-registry/v1"


@dataclass(frozen=True, slots=True)
class IndexQueueDefaults:
    treatment: str = ""
    do_up: bool = True
    do_both: bool = False
    avg_steps: int = 1

    def __post_init__(self) -> None:
        if not isinstance(self.treatment, str) or any(ord(char) < 32 for char in self.treatment):
            raise ValueError("Enter a treatment sequence on one line.")
        if type(self.do_up) is not bool or type(self.do_both) is not bool:
            raise ValueError("Choose a measurement direction.")
        validate_average_count(self.avg_steps)
        if self.treatment.strip():
            self.resolve_treatment("NRM")

    def resolve_treatment(self, fallback: str) -> str:
        labels = parse_queue_sequence(self.treatment.strip() or fallback)
        if self.do_both and len(labels) != 1:
            raise ValueError("Paired acquisition requires exactly one treatment. Choose one treatment or one side.")
        return " → ".join(labels)


def _entry_key(path: Path) -> str:
    identity = os.path.normcase(str(path))
    return hashlib.sha256(os.fsencode(identity)).hexdigest()


def _sync(settings: QtCore.QSettings) -> None:
    settings.sync()
    if settings.status() != QtCore.QSettings.Status.NoError:
        raise OSError("The sample index registry could not access the operator settings.")


def read_registered_indexes(settings: QtCore.QSettings) -> tuple[dict[Path, IndexQueueDefaults], int]:
    _sync(settings)
    entries = {}
    invalid = 0
    settings.beginGroup(_GROUP)
    try:
        for key in settings.childKeys():
            try:
                saved = json.loads(settings.value(key, "", type=str))
                if not isinstance(saved, dict):
                    raise ValueError("Invalid registered index")
                path = Path(saved.pop("path"))
                if not path.is_absolute() or _entry_key(path) != key:
                    raise ValueError("Invalid registered index path")
                entries[path] = IndexQueueDefaults(**saved)
            except (KeyError, TypeError, ValueError):
                invalid += 1
    finally:
        settings.endGroup()
    return entries, invalid


def save_registered_index(settings: QtCore.QSettings, path: Path, defaults: IndexQueueDefaults) -> None:
    path = path.resolve(strict=True)
    if not path.is_file():
        raise ValueError("Choose a readable sample index file.")
    settings.setValue(f"{_GROUP}/{_entry_key(path)}", json.dumps({"path": str(path), **asdict(defaults)}))
    _sync(settings)


def remove_registered_index(settings: QtCore.QSettings, path: Path) -> None:
    settings.remove(f"{_GROUP}/{_entry_key(path)}")
    _sync(settings)
