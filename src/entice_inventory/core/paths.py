"""Project paths and data-file resolution.

Auxiliary workbooks (GTAP12_matching.xlsx, GTAP12_X.xlsx, Regions_clusters.xlsx)
live in ``<project_root>/data``. ``find_data_file`` locates them there, while
still honouring an explicit ``base`` directory passed by a caller (so the
upstream ``make_inventory(repo_path=...)`` API keeps working whether it is given
the project root or the data folder directly).
"""

from __future__ import annotations

from pathlib import Path

# core/paths.py -> core -> entice_inventory -> src -> <project_root>
PROJECT_ROOT = Path(__file__).resolve().parents[3]
DATA_DIR = PROJECT_ROOT / "data"


def find_data_file(name: str, base: str | Path | None = None) -> Path:
    """Resolve a bundled data workbook by name.

    Search order: ``base/name``, ``base/data/name``, then ``DATA_DIR/name``.
    Returns the first path that exists, or ``DATA_DIR/name`` as a default so the
    caller gets a sensible (if missing) path to report.
    """
    candidates: list[Path] = []
    if base is not None:
        base = Path(base)
        candidates += [base / name, base / "data" / name]
    candidates.append(DATA_DIR / name)
    for candidate in candidates:
        if candidate.exists():
            return candidate
    return DATA_DIR / name
