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

# Paths are relative to Inventory generation. A local override may point outside
# it, e.g. to the sibling consortium Data collection directory.
DATA_PATHS = {
    'baseline': 'Database/GTAP Power 2023/Cache',
    'gtap_totals': 'Database/GTAP Power 2023/Cache/GTAP12_X.xlsx',
    'classifications': 'Classifications/ENTICE classifications.xlsx',
    'mario_inventories': 'Data collection/Inventory cleaning/MARIO inventories',
    'reference_inventories': 'Data collection/Inventory cleaning/D2.4 inventories',
    'inventory_runs': 'Inventories',
    'reference_coefficients': 'Data collection/Inventory cleaning/D2.4 database/coefficients',
    'value_added': 'Data collection/Inventory cleaning/VA.xlsx',
    'purdue_split_targets': '../Data collection/PURDUE/Shared material/Output and trade targets/splttargs.xlsx',
    'purdue_trade': '../Data collection/PURDUE/Shared material/Trades/trade.xlsx',
    'gtap_sector_names': 'Classifications/_sources/Purdue/GTAP sectors H5.xlsx',
    'exioiot': '../Data collection/PURDUE/Shared material/Cost structures/EXIOIOT.csv',
    'copcosts': '../Data collection/PURDUE/Shared material/Cost structures/COPCOSTS.csv',
    'copper_sectors': '../Data collection/PURDUE/Shared material/Cost structures/COPSEC.csv',
}


def local_config() -> dict:
    path = Path(os.environ.get('ENTICE_CONFIG', str(PROJECT_ROOT/'paths.local.json'))).expanduser()
    if not path.exists():
        return {}
    config = json.loads(path.read_text())
    if not isinstance(config, dict):
        raise ValueError(f'Expected a JSON object in {path}')
    return config


def data_path(key: str, *, required: bool = False) -> Path:
    """Resolve an explicit dataset role, without guessing from filenames."""
    if key not in DATA_PATHS:
        raise KeyError(f'Unknown data path: {key}')
    overrides = local_config().get('paths', {})
    if not isinstance(overrides, dict):
        raise ValueError('paths must be a JSON object')
    value = overrides.get(key, DATA_PATHS[key])
    if not isinstance(value, str) or not value.strip():
        raise ValueError(f'Invalid path for {key}')
    result = Path(value).expanduser()
    if not result.is_absolute():
        result = project_data_root()/result
    result = result.resolve()
    if required and not result.exists():
        raise FileNotFoundError(f'{key} is unavailable: {result}')
    return result


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
        candidates.append(data_path('gtap_totals'))
        # Compatibility for existing installations until their config is updated.
        candidates.append(project_data_root()/'Repository inputs'/name)
    candidates.append(DATA_DIR / name)
    for candidate in candidates:
        if candidate.exists():
            return candidate
    return DATA_DIR / name
