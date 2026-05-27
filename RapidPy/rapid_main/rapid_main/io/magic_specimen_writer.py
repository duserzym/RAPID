from __future__ import annotations

from pathlib import Path

from rapid_main.data_model import SpecimenMeta


SPECIMEN_COLUMNS = [
    "specimen",
    "sample",
    "site",
    "location",
    "volume",
    "description",
]

_HEADER_LINE = "tab delimited\tspecimens"


def _build_row(meta: SpecimenMeta) -> list[str]:
    return [
        meta.name,
        meta.sample,
        meta.site,
        meta.location,
        f"{meta.volume:.6g}",
        meta.comment,
    ]


def append_specimen(path: str | Path, meta: SpecimenMeta) -> None:
    """Append one specimen row to MagIC specimen.txt, creating headers if needed.

    De-duplicates by specimen name when the file already exists.
    """
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)

    if path.exists():
        lines = path.read_text(encoding="utf-8", errors="replace").splitlines()
        for line in lines[2:]:
            if not line.strip():
                continue
            if line.split("\t", 1)[0] == meta.name:
                return

    write_header = not path.exists()
    with path.open("a", encoding="utf-8", newline="\n") as fh:
        if write_header:
            fh.write(_HEADER_LINE + "\n")
            fh.write("\t".join(SPECIMEN_COLUMNS) + "\n")
        fh.write("\t".join(_build_row(meta)) + "\n")
