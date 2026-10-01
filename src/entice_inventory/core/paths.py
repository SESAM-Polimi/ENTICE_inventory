"""Project paths and data-file resolution.

Classification workbooks live in ``<project_root>/data``. Licensed GTAP totals
live in the configured shared data root. ``find_data_file`` honours an explicit
``base`` directory first, so ``make_inventory(repo_path=...)`` accepts either
the project root or a data folder directly.
"""

from __future__ import annotations

import json
import os
from pathlib import Path

# core/paths.py -> core -> entice_inventory -> src -> <project_root>
PROJECT_ROOT = Path(__file__).resolve().parents[3]
DATA_DIR = PROJECT_ROOT / "data"


def project_data_root(required: bool = False) -> Path:
    """Resolve shared operational data from the environment or local config.

    ENTICE_DATA_ROOT takes precedence over paths.local.json's data_root.
    ENTICE_CONFIG can select a config outside the checkout (e.g. wheel installs).
    No user-specific or old SharePoint location is embedded in public code.
    """
    value = os.environ.get('ENTICE_DATA_ROOT')
    config_path = Path(os.environ.get('ENTICE_CONFIG', str(PROJECT_ROOT/'paths.local.json'))).expanduser()
    if not value and config_path.exists():
        config = json.loads(config_path.read_text())
        if not isinstance(config, dict):
            raise ValueError(f'Expected a JSON object in {config_path}')
        value = config.get('data_root')
        if value is not None and (not isinstance(value, str) or not value.strip()):
            raise ValueError(f'Invalid data_root in {config_path}')
        base = config_path.resolve().parent
    else:
        base = Path.cwd()
    if value:
        result = Path(value).expanduser()
        if not result.is_absolute():
            result = base/result
    else:
        result = PROJECT_ROOT/'external-data'
        if required:
            raise FileNotFoundError('Set ENTICE_DATA_ROOT or data_root in paths.local.json to the shared ENTICE inventory directory.')
    if required and not result.is_dir():
        raise FileNotFoundError(f'Shared inventory data directory is unavailable: {result}')
    return result


def find_data_file(name: str, base: str | Path | None = None) -> Path:
    """Resolve a classification or externally supplied data workbook by name.

    Search order: ``base/name``, ``base/data/name``, the configured shared
    location for licensed GTAP totals, then ``DATA_DIR/name``.
    Returns the first path that exists, or ``DATA_DIR/name`` as a default so the
    caller gets a sensible (if missing) path to report.
    """
    candidates: list[Path] = []
    if base is not None:
        base = Path(base)
        candidates += [base / name, base / "data" / name]
    if name == 'GTAP12_X.xlsx':
        candidates.append(project_data_root()/'Repository inputs'/name)
    candidates.append(DATA_DIR / name)
    for candidate in candidates:
        if candidate.exists():
            return candidate
    return DATA_DIR / name
