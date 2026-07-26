"""Pure helpers for validating files received through drag and drop."""

from __future__ import annotations

from collections import Counter
from pathlib import Path
from typing import Iterable

from conversion_specs import source_kind_for_extension


SOURCE_KIND_DISPLAY = {
    "pdf": "PDF",
    "word": "Word",
    "powerpoint": "PowerPoint",
}


class MixedSourceKindsError(ValueError):
    """Raised when one drop contains more than one supported source kind."""

    def __init__(self, counts: Counter[str]):
        self.counts = counts.copy()
        details = "、".join(
            f"{SOURCE_KIND_DISPLAY[kind]} {count} 个"
            for kind, count in self.counts.items()
        )
        super().__init__(f"一次拖入了多种源格式：{details}")


def classify_dropped_paths(
    paths: Iterable[str | Path],
) -> tuple[str | None, list[Path], list[Path]]:
    """Return the single supported source kind, accepted files, and ignored paths."""
    accepted: list[Path] = []
    ignored: list[Path] = []
    counts: Counter[str] = Counter()

    for raw_path in paths:
        path = Path(raw_path)
        if not path.is_file():
            ignored.append(path)
            continue
        try:
            source_kind = source_kind_for_extension(path.suffix)
        except ValueError:
            ignored.append(path)
            continue
        accepted.append(path)
        counts[source_kind] += 1

    if len(counts) > 1:
        raise MixedSourceKindsError(counts)

    source_kind = next(iter(counts), None)
    return source_kind, accepted, ignored
