"""Export one inventory from an existing D2.4 coefficient database.

Reads only the requested sector's parquet columns. Does not rebuild or optimise
the database. Parent policy and trade source must be explicit; unchanged source
workbooks are required. Use a fresh output directory to retain prior releases.
Requires pandas, pyarrow and openpyxl. No MARIO/CVXLAB import is needed at this
export-only stage. This is not a replacement for the full producer build.
"""
from __future__ import annotations

import argparse
import ast
import hashlib
import json
import math
import sys
import warnings
from datetime import datetime
from pathlib import Path
from types import SimpleNamespace

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / 'src'))
import pandas as pd
import pyarrow.parquet as pq
from entice_inventory.export.build_d24_inventories import (
    aggregate_sector_coefficients, build_trade_rows, create_output_workbook,
    read_sector_templates, read_total_outputs, read_purdue_trade_tables,
    read_purdue_trade_workbook, render_inventory_rows, resolve_export_region_maps,
    load_matching_parents,
)


def load_sector_frame(path: Path, code: str):
    source = pq.ParquetFile(path)
    metadata = json.loads(source.schema_arrow.metadata[b'pandas'])
    columns = []
    for column in metadata['columns']:
        try:
            label = ast.literal_eval(column['name'])
        except (ValueError, SyntaxError, TypeError):
            continue
        if isinstance(label, tuple) and len(label) == 3 and label[1:] == ('Sector', code):
            columns.append(column['field_name'])
    if not columns:
        raise ValueError(f'{code} is absent from {path.name}.')
    indexes = [i for i in metadata['index_columns'] if isinstance(i, str)]
    return source.read(columns=columns + indexes).to_pandas()


def main():
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument('--sector', required=True)
    ap.add_argument('--coefficients-dir', type=Path, required=True)
    sources = ap.add_mutually_exclusive_group(required=True)
    sources.add_argument('--source-dir', type=Path)
    sources.add_argument('--source-workbook', type=Path, help='Read only this source workbook, even if other files are damaged.')
    ap.add_argument('--trade-workbook', type=Path, required=True)
    ap.add_argument('--trade-sheet', choices=['NTSCIF', 'BLTTRD'], required=True)
    ap.add_argument('--parent-policy', choices=['source', 'matching'], required=True)
    ap.add_argument('--matching', type=Path)
    ap.add_argument('--output-dir', type=Path, required=True)
    ap.add_argument('--evidence-dir', type=Path,
                    help='Technical record directory; defaults to ignored build/reexports/<timestamp>.')
    ap.add_argument('--tolerance', type=float, default=1e-12)
    args = ap.parse_args()
    if not math.isfinite(args.tolerance) or args.tolerance < 0:
        ap.error('--tolerance must be a finite non-negative value.')
    code = args.sector.strip().upper()
    if args.parent_policy == 'matching' and args.matching is None:
        ap.error('--matching is required with --parent-policy matching.')
    if args.matching is not None and args.parent_policy == 'source':
        ap.error('--matching is unused with --parent-policy source; omit it.')
    if args.parent_policy == 'matching':
        warnings.warn('Matching changes export metadata only. The coefficient database must already use the reviewed parent assignments.')
    parents = {} if args.parent_policy == 'source' else load_matching_parents(args.matching)
    templates, skipped = read_sector_templates(
        args.source_dir or args.source_workbook.parent, parents, [code],
        [args.source_workbook] if args.source_workbook else None)
    if len(templates) != 1:
        raise ValueError(f'Expected one usable source for {code}, found {len(templates)}; skipped: {skipped}')
    template = templates[0]
    if (args.output_dir / f'{code}.xlsx').exists() or any(args.output_dir.glob(f'{code} - *.xlsx')):
        raise FileExistsError(f'{code} already exists in the output directory; choose a fresh directory.')
    frames = {name: load_sector_frame(args.coefficients_dir/f'{name}.parquet', code) for name in ('z', 'v', 'e')}
    units_frame = pd.read_parquet(args.coefficients_dir/'units.parquet')
    units = {kind: frame.droplevel(0) for kind, frame in units_frame.groupby(level=0, sort=False)}
    db = SimpleNamespace(**frames)
    outputs = read_total_outputs(template.source_path, code)
    global_regions = list(dict.fromkeys(frames['z'].index.get_level_values('Region')))
    global_regions, clusters, _ = resolve_export_region_maps(global_regions)
    orders = [units[kind].index.tolist() for kind in ('Sector', 'Factor of production', 'Satellite account')]
    payload = aggregate_sector_coefficients(db, code, outputs.keys())
    inventories, empty = render_inventory_rows(payload, {k:v['unit'].to_dict() for k,v in units.items()}, *orders, args.tolerance)
    trades, _ = (read_purdue_trade_tables(args.trade_workbook) if args.trade_sheet == 'BLTTRD'
                 else read_purdue_trade_workbook(args.trade_workbook, args.trade_sheet))
    trade_rows = build_trade_rows(code, trades.get(code), [r for r,_ in inventories], args.tolerance)
    target = create_output_workbook(template, args.output_dir, global_regions, clusters, units, outputs, inventories, trade_rows)
    manifest = {'sector': code, 'method': 'Export existing coefficient columns; no source-data rebuild or optimisation.',
                'arguments': {k:str(v) if isinstance(v,Path) else v for k,v in vars(args).items()},
                'output_sha256': hashlib.sha256(target.read_bytes()).hexdigest(), 'empty_regions': empty,
                'inventory_regions': len(inventories), 'trade_rows': len(trade_rows),
                'source_sha256': hashlib.sha256(template.source_path.read_bytes()).hexdigest(),
                'trade_workbook_sha256': hashlib.sha256(args.trade_workbook.read_bytes()).hexdigest()}
    evidence_dir = args.evidence_dir or Path(__file__).resolve().parents[1]/'build/reexports'/datetime.now().strftime('%Y%m%d-%H%M%S-%f')
    evidence_dir.mkdir(parents=True, exist_ok=True)
    (evidence_dir/f'{code}.export.json').write_text(json.dumps(manifest,indent=2)+'\n')
    print(json.dumps(manifest,indent=2))


if __name__ == '__main__':
    main()
