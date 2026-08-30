from __future__ import annotations

from dataclasses import dataclass
from typing import Iterable


@dataclass(slots=True)
class QueueSample:
    sample_name: str
    file_id: str
    hole: int
    do_up: bool = True
    do_both: bool = False
    measurement_step_count: int = 1


@dataclass(slots=True)
class QueueCommand:
    command_type: str
    hole: int = 0
    file_id: str = ""
    sample_name: str = ""


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
    for idx, item in enumerate(samples):
        if not item.sample_name:
            errors.append(f"row {idx + 1}: sample_name is required")

        if item.hole < 1:
            errors.append(f"row {idx + 1}: hole must be >= 1")

        if item.measurement_step_count < 1:
            errors.append(
                f"row {idx + 1}: measurement_step_count must be >= 1"
            )

        if item.hole in seen_holes:
            warnings.append(
                f"duplicate hole {item.hole}: rows {seen_holes[item.hole] + 1} and {idx + 1}"
            )
        else:
            seen_holes[item.hole] = idx

    if not seen_holes:
        warnings.append("queue is empty; no measurement commands will be generated")

    return QueueValidationResult(tuple(errors), tuple(warnings))


def compile_queue(
    samples: list[QueueSample],
    options: QueueOptions,
    *,
    strict: bool = False,
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
        cmds.append(QueueCommand("Meas", item.hole, item.file_id, item.sample_name))
        measured_count += 1
        if (
            options.repeat_holder
            and options.samples_between_holder > 0
            and measured_count % options.samples_between_holder == 0
            and cmds[-1].command_type != "Holder"
        ):
            cmds.append(QueueCommand("Holder", item.hole))

    goto_hole = -1 if options.use_xy_table else _first_hole(samples, options.ascending)
    if options.load_return:
        cmds.append(QueueCommand("Goto", goto_hole))

    cmds = preprocess_queue(cmds, per_file)

    if options.do_return:
        cmds.append(QueueCommand("Goto", goto_hole))

    return cmds


def preprocess_queue(commands: list[QueueCommand], per_file: dict[str, QueueSample]) -> list[QueueCommand]:
    """Apply VB6-style preprocessing for dual-side sample measurements.

    For files marked doBoth+doUp, insert a Flip marker and duplicate Meas
    entries later in the queue so both orientations are measured.
    """
    processed: list[QueueCommand] = []
    needs_second_pass: dict[str, bool] = {}

    for cmd in commands:
        processed.append(cmd)
        if cmd.command_type == "InitUp" and cmd.file_id:
            item = per_file.get(cmd.file_id)
            if item and item.do_both and item.do_up and item.measurement_step_count <= 1:
                needs_second_pass[cmd.file_id] = True
                processed.append(QueueCommand("Flip", -1, cmd.file_id, ""))
                processed.append(QueueCommand("Holder", cmd.hole))

    if not needs_second_pass:
        return processed

    second_pass: list[QueueCommand] = []
    for cmd in processed:
        second_pass.append(cmd)
        if cmd.command_type == "Meas" and cmd.file_id in needs_second_pass:
            second_pass.append(QueueCommand("Meas", cmd.hole, cmd.file_id, cmd.sample_name))

    return second_pass
