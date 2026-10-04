from __future__ import annotations

from dataclasses import dataclass, replace
from typing import Iterable


@dataclass(slots=True)
class QueueSample:
    sample_name: str
    file_id: str
    hole: int
    do_up: bool = True
    do_both: bool = False
    measurement_step_count: int = 1
    measurement_labels: tuple[str, ...] = ()
    source_file: str = ""
    row_id: str = ""


@dataclass(slots=True)
class QueueCommand:
    command_type: str
    hole: int = 0
    file_id: str = ""
    sample_name: str = ""
    measurement_labels: tuple[str, ...] = ()
    row_id: str = ""


@dataclass(slots=True)
class QueueOptions:
    ascending: bool = True
    load_return: bool = True
    do_return: bool = True
    repeat_holder: bool = True
    samples_between_holder: int = 8
    use_xy_table: bool = True


def _first_hole(samples: list[QueueSample], ascending: bool) -> int:
    if not samples:
        return 0
    holes = sorted(item.hole for item in samples)
    return holes[0] if ascending else holes[-1]


@dataclass(frozen=True)
class QueueValidationResult:
    """Container for queue validation results."""

    errors: tuple[str, ...]
    warnings: tuple[str, ...]

    def is_valid(self) -> bool:
        return not self.errors


def validate_queue_samples(samples: Iterable[QueueSample]) -> QueueValidationResult:
    """Validate queue metadata before command compilation.

    This helper is intentionally conservative and is intended to gate queue start
    once `rapid_main` wiring is extended into the control workflow.
    """
    errors: list[str] = []
    warnings: list[str] = []

    seen_holes: dict[int, int] = {}
    file_settings: dict[str, tuple] = {}
    row_ids = set()
    for idx, item in enumerate(samples):
        if not item.sample_name:
            errors.append(f"row {idx + 1}: sample_name is required")

        if type(item.hole) is not int or item.hole < 1:
            errors.append(f"row {idx + 1}: hole must be >= 1")

        if type(item.measurement_step_count) is not int or item.measurement_step_count < 1:
            errors.append(
                f"row {idx + 1}: measurement_step_count must be >= 1"
            )
        if not isinstance(item.file_id, str) or not item.file_id.strip():
            errors.append(f"row {idx + 1}: file_id is required")
        if type(item.do_up) is not bool or type(item.do_both) is not bool:
            errors.append(f"row {idx + 1}: orientation settings must be explicit booleans")
        labels = item.measurement_labels
        if (type(labels) is not tuple or any(not isinstance(label, str) or not label.strip() for label in labels)):
            errors.append(f"row {idx + 1}: measurement labels must be a tuple of nonempty strings")
        elif labels and len(labels) != item.measurement_step_count:
            errors.append(f"row {idx + 1}: measurement step count does not match its labels")
        if not isinstance(item.source_file, str):
            errors.append(f'row {idx + 1}: source_file must be a string')
        if not isinstance(item.row_id, str):
            errors.append(f'row {idx + 1}: row_id must be a string')
        elif item.row_id:
            try:
                if len(item.row_id) != 32: raise ValueError('length')
                int(item.row_id, 16)
                if item.row_id in row_ids: raise ValueError('duplicate')
                row_ids.add(item.row_id)
            except ValueError:
                errors.append(f'row {idx + 1}: row_id must be unique 32-digit hexadecimal identity')
        settings = (item.do_up, item.do_both, item.measurement_step_count, labels, item.source_file)
        if isinstance(item.file_id, str) and item.file_id in file_settings and file_settings[item.file_id] != settings:
            errors.append(f"row {idx + 1}: rows from one file must agree on measurement steps and orientation settings")
        elif isinstance(item.file_id, str):
            file_settings[item.file_id] = settings

        if type(item.hole) is not int:
            continue
        if item.hole in seen_holes:
            warnings.append(
                f"duplicate hole {item.hole}: rows {seen_holes[item.hole] + 1} and {idx + 1}"
            )
        else:
            seen_holes[item.hole] = idx

    if not seen_holes:
        warnings.append("queue is empty; no measurement commands will be generated")

    return QueueValidationResult(tuple(errors), tuple(warnings))


def resolve_queue_samples(samples: Iterable[QueueSample], default_labels: Iterable[str]) -> list[QueueSample]:
    """Snapshot each file's executable labels before queue compilation/startup."""
    defaults = tuple(default_labels)
    samples = list(samples)
    needs_defaults = any(type(item.measurement_labels) is tuple and not item.measurement_labels for item in samples)
    if (needs_defaults and not defaults) or any(not isinstance(label, str) or not label.strip() for label in defaults):
        raise ValueError('Default measurement sequence must contain nonempty labels.')
    return [replace(item, measurement_labels=defaults, measurement_step_count=len(defaults))
            if type(item.measurement_labels) is tuple and not item.measurement_labels else replace(item) for item in samples]


def compile_queue(
    samples: list[QueueSample],
    options: QueueOptions,
    *,
    strict: bool = False,
    default_labels: Iterable[str] | None = None,
) -> list[QueueCommand]:
    """Compile VB6-style queue commands from sample metadata.

    This follows the frmChanger/ProcessSamplesToQueue shape:
    1) optional InitUp markers per file
    2) Holder measurement first
    3) ordered Meas commands across holes
    4) periodic Holder inserts
    5) optional load-return Goto
    6) preprocess for doBoth/doUp semantics
    7) optional final return Goto
    """
    if default_labels is not None:
        samples = resolve_queue_samples(samples, default_labels)
    validation = validate_queue_samples(samples)
    if strict and validation.errors:
        raise ValueError(f"Invalid queue specification: {'; '.join(validation.errors)}")

    if not samples:
        return []

    cmds: list[QueueCommand] = []

    per_file: dict[str, QueueSample] = {}
    for item in samples:
        if item.file_id and item.file_id not in per_file:
            per_file[item.file_id] = item

    for file_item in per_file.values():
        do_both = file_item.do_both and file_item.measurement_step_count <= 1
        if do_both and file_item.do_up:
            cmds.append(QueueCommand("InitUp", 0, file_item.file_id, ""))

    cmds.append(QueueCommand("Holder"))

    ordered = sorted(samples, key=lambda s: s.hole, reverse=not options.ascending)
    measured_count = 0
    for item in ordered:
        cmds.append(QueueCommand("Meas", item.hole, item.file_id, item.sample_name, item.measurement_labels, item.row_id))
        measured_count += 1
        if (
            options.repeat_holder
            and options.samples_between_holder > 0
            and measured_count % options.samples_between_holder == 0
            and cmds[-1].command_type != "Holder"
        ):
            # Holder is a blank at a station-resolved empty hole, never the
            # specimen slot just measured. Zero requests that empty-hole path.
            cmds.append(QueueCommand("Holder", 0))

    goto_hole = -1 if options.use_xy_table else _first_hole(samples, options.ascending)
    if options.load_return:
        cmds.append(QueueCommand("Goto", goto_hole))

    cmds = preprocess_queue(cmds, per_file, options=options)

    if options.do_return:
        cmds.append(QueueCommand("Goto", goto_hole))

    return cmds


def preprocess_queue(commands: list[QueueCommand], per_file: dict[str, QueueSample], *, options: QueueOptions | None = None) -> list[QueueCommand]:
    """Apply VB6-style preprocessing for dual-side sample measurements.

    For files marked doBoth+doUp, insert a Flip marker and duplicate Meas
    entries later in the queue so both orientations are measured.
    """
    options = options or QueueOptions()
    appended: list[QueueCommand] = []
    needs_second_pass: dict[str, bool] = {}
    second_pass_count = 0

    for cmd in commands:
        if cmd.command_type == "InitUp" and cmd.file_id:
            item = per_file.get(cmd.file_id)
            if item and item.do_both and item.do_up and item.measurement_step_count <= 1:
                needs_second_pass[cmd.file_id] = True
                appended.append(QueueCommand("Flip", -1 if options.use_xy_table else 0, cmd.file_id, ""))
                appended.append(QueueCommand("Holder", 0))
        elif cmd.command_type == "Flip" and cmd.file_id:
            needs_second_pass[cmd.file_id] = False
        elif cmd.command_type == "Meas" and needs_second_pass.get(cmd.file_id, False):
            appended.append(replace(cmd))
            second_pass_count += 1
            if (options.repeat_holder and options.samples_between_holder > 0
                    and second_pass_count % options.samples_between_holder == 0):
                appended.append(QueueCommand("Holder", 0))

    # VB6 SampleCommands.Preprocess appends to the collection while visiting
    # its original entries. Flip and repeat measurements therefore follow the
    # entire original pass; inserting them beside InitUp flips before any read.
    return list(commands) + appended
