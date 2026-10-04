"""Portable specimen paths bounded by their selected source/output directory."""
from pathlib import Path, PureWindowsPath
import hashlib


_FIXED_OUTPUTS = {
    'measurements.txt', 'specimens.txt', 'provenance.json', 'artifact_index.json',
    'workflow_summary.json', 'quicklook.json', 'communication.tsv', 'susceptibility.json',
    'rockmag_run.json', 'thermal_run.json', 'simulation.txt', '.rapidpy-steps.json',
    'af_treatments', 'pulse_treatments', 'susceptibility_acquisitions',
}
_DEVICES = {'CON', 'PRN', 'AUX', 'NUL', 'CONIN$', 'CONOUT$'} | {
    prefix + number for prefix in ('COM', 'LPT') for number in '123456789¹²³'
}


def relative_specimen_path(name: str) -> Path:
    if not isinstance(name, str) or not name.strip():
        raise ValueError('A specimen path must have a name.')
    portable = PureWindowsPath(name)
    if portable.drive or portable.root:
        raise ValueError('Specimen names must be relative to the selected directory.')
    for part in portable.parts:
        if part in {'.', '..'}:
            continue
        if (any(ord(char) < 32 or char in '<>:"|?*' for char in part)
                or part.endswith((' ', '.')) or part.split('.')[0].upper() in _DEVICES):
            raise ValueError('Specimen names must use portable file-name components.')
    return Path(*portable.parts)


def contained_path(directory: str | Path, name: str, *, output: bool = False) -> Path:
    original_base = Path(directory).absolute()
    base = original_base.resolve()
    if output and base != original_base:
        raise ValueError('The admitted output directory changed or contains a link.')
    relative = relative_specimen_path(name)
    target = (base / relative).resolve()
    if target == base or not target.is_relative_to(base):
        raise ValueError('Specimen path escapes the selected directory.')
    part_path = base
    for part in relative.parts:
        part_path = part_path / part
        if not part_path.resolve().is_relative_to(base):
            raise ValueError('Specimen path escapes the selected directory.')
        if output and (part_path.is_symlink() or getattr(part_path, 'is_junction', lambda: False)()):
            raise ValueError('Output paths cannot use a link or junction below the selected directory.')
    if output and target.is_file() and target.stat().st_nlink > 1:
        raise ValueError('Output files cannot be shared through hard links.')
    return target


def specimen_output_path(directory: str | Path, name: str) -> Path:
    target = contained_path(directory, name, output=True)
    relative = target.relative_to(Path(directory).resolve()).as_posix().casefold()
    first_component = relative.split('/')[0]
    if first_component in _FIXED_OUTPUTS or first_component.startswith('.rapidpy-pending'):
        raise ValueError('Specimen name conflicts with a reserved measurement artifact.')
    return target


def specimen_run_directory(directory: str | Path, name: str, index: Path | None = None) -> Path:
    root = Path(directory).resolve()
    if index is not None:
        group = index.stem + '-' + hashlib.sha256(str(index).encode('utf-8')).hexdigest()[:12]
        root = contained_path(root, group, output=True)
    return specimen_output_path(root, name)
