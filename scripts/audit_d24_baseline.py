"""Read-only audit of a published D2.4 archive and its local source candidates.

No MARIO/GTAP build is run and no input file is modified. Detailed results are
written only to --output-dir. Trade/output comparisons are screening diagnostics:
their economic interpretation requires aligned coverage, valuation and units.
"""
from __future__ import annotations

import argparse
import csv
import hashlib
import io
import json
import math
import re
import zipfile
from collections import Counter, defaultdict
from datetime import datetime, timezone
from pathlib import Path
from xml.etree import ElementTree as ET

NS = {'m': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
REL = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id'
STATIC = {'Summary', 'Master', 'Regions Clusters', 'Sectors Clusters',
          'Factors Clusters', 'Total outputs', 'Trades', 'DB units'}


def digest(data):
    return hashlib.sha256(data).hexdigest()


def text(value):
    return '' if value is None else str(value).strip()


def number(value):
    if value is None or isinstance(value, bool) or text(value) == '':
        return None
    try:
        result = float(value)
        return result if math.isfinite(result) else None
    except (TypeError, ValueError):
        return None


class Workbook:
    """Minimal cached-value OOXML reader; never calculates or saves a workbook."""

    def __init__(self, raw):
        self.archive = zipfile.ZipFile(io.BytesIO(raw))
        self.strings = []
        if 'xl/sharedStrings.xml' in self.archive.namelist():
            tree = ET.fromstring(self.archive.read('xl/sharedStrings.xml'))
            self.strings = [''.join(e.itertext()) for e in tree]
        relationships = ET.fromstring(self.archive.read('xl/_rels/workbook.xml.rels'))
        targets = {e.attrib['Id']: e.attrib['Target'] for e in relationships}
        tree = ET.fromstring(self.archive.read('xl/workbook.xml'))
        self.sheets = {}
        for e in tree.find('m:sheets', NS):
            target = targets[e.attrib[REL]]
            self.sheets[e.attrib['name']] = (target.lstrip('/') if target.startswith('/')
                                             else 'xl/' + target)
        self.cache = {}
        self.errors = []
        self.uncached_formulas = []

    def rows(self, name):
        if name in self.cache:
            return self.cache[name]
        out = []
        if name not in self.sheets:
            return out
        tree = ET.fromstring(self.archive.read(self.sheets[name]))
        for row in tree.findall('m:sheetData/m:row', NS):
            values = {}
            for cell in row.findall('m:c', NS):
                address = cell.attrib['r']
                col = re.sub(r'\d', '', address)
                value = cell.find('m:v', NS)
                kind = cell.attrib.get('t', '')
                if kind == 'inlineStr':
                    value = ''.join(e.text or '' for e in cell.findall('.//m:t', NS))
                elif value is None or value.text is None:
                    value = None
                elif kind == 's':
                    value = self.strings[int(value.text)]
                elif kind in ('str', 'e', 'd'):
                    value = value.text
                elif kind == 'b':
                    value = value.text == '1'
                else:
                    value = number(value.text)
                if kind == 'e':
                    self.errors.append((name, address, value))
                if cell.find('m:f', NS) is not None and value is None:
                    self.uncached_formulas.append((name, address))
                if value is not None and value != '':
                    values[col] = value
            if values:
                out.append((int(row.attrib['r']), values))
        self.cache[name] = out
        return out

    def cell(self, name, address):
        col = re.sub(r'\d', '', address)
        row = int(re.sub(r'\D', '', address))
        return next((r.get(col) for i, r in self.rows(name) if i == row), None)

    def semantic_hash(self):
        payload = [(name, self.rows(name)) for name in sorted(self.sheets)]
        return digest(json.dumps(payload, sort_keys=True, separators=(',', ':')).encode())


def write_csv(path, rows, fields=None):
    fields = fields or list(dict.fromkeys(k for row in rows for k in row))
    with path.open('w', newline='', encoding='utf-8') as f:
        writer = csv.DictWriter(f, fieldnames=fields)
        writer.writeheader()
        writer.writerows(rows)


def analyze(raw, filename, stage, issues, inventory_checks, trade_checks):
    w = Workbook(raw)
    code = text(w.cell('Summary', 'B2'))
    parent = text(w.cell('Summary', 'B3'))

    def issue(check, detail, sheet='', row='', severity='review'):
        issues.append(dict(stage=stage, file=filename, sector=code, check=check,
                           severity=severity, sheet=sheet, row=row, detail=detail))

    for required in ('Summary', 'Master', 'Total outputs'):
        if required not in w.sheets:
            issue('missing_sheet', required, severity='error')
    outputs = {}
    for idx, row in w.rows('Total outputs'):
        if idx == 1:
            continue
        region, value, unit = text(row.get('B')), number(row.get('C')), text(row.get('D'))
        if not region:
            continue
        if text(row.get('A')) != code:
            issue('output_sector_mismatch', text(row.get('A')), 'Total outputs', idx, 'error')
        if region in outputs:
            issue('duplicate_output_region', region, 'Total outputs', idx, 'error')
        if value is None or value < 0:
            issue('invalid_output', f'{region}: {row.get("C")}', 'Total outputs', idx, 'error')
        outputs[region] = {'value': value, 'unit': unit, 'row': idx}

    linked = set()
    master_regions = set()
    for idx, row in w.rows('Master'):
        if idx == 1:
            continue
        region, sn = text(row.get('A')), text(row.get('C'))
        if not sn:
            continue
        linked.add(sn)
        master_regions.add(region)
        if sn not in w.sheets:
            issue('missing_inventory_sheet', sn, 'Master', idx, 'error')
        if text(row.get('B')) != code or text(row.get('H')) != parent:
            issue('master_metadata_mismatch', f'{region}: sector={row.get("B")}; parent={row.get("H")}', 'Master', idx)
    actual = set(w.sheets) - STATIC
    if not actual:
        issue('no_inventory_sheets', 'No regional inventory sheets are present.', severity='error')
    for sn in sorted(actual - linked):
        issue('unlinked_inventory_sheet', sn)
    for sn in sorted(linked & actual):
        total = 0.0
        n = negative = satellite = 0
        for idx, row in w.rows(sn):
            if idx == 1:
                continue
            kind = text(row.get('D'))
            if kind == 'Satellite account':
                satellite += 1
                continue
            if kind not in ('Sector', 'Factor of production'):
                if row:
                    issue('unknown_inventory_item_type', kind, sn, idx)
                continue
            val = number(row.get('A'))
            if val is None:
                issue('invalid_coefficient', text(row.get('A')), sn, idx, 'error')
                continue
            n += 1
            total += val
            negative += val < 0
        inventory_checks.append(dict(stage=stage, file=filename, sector=code, sheet=sn,
                                     monetary_sum=total, monetary_rows=n, negative_rows=negative,
                                     satellite_rows_excluded=satellite,
                                     outside_1e_3=abs(total-1)>1e-3,
                                     outside_1e_2=abs(total-1)>1e-2))

    exports = defaultdict(float)
    imports = defaultdict(float)
    trade_units = set()
    trade_keys = set()
    valid_trades = []
    for idx, row in w.rows('Trades'):
        if idx == 1:
            continue
        sec, origin, destination = (text(row.get(c)) for c in ('A', 'B', 'C'))
        val, unit = number(row.get('D')), text(row.get('E'))
        if not (sec or origin or destination):
            continue
        key = (sec, origin, destination)
        if key in trade_keys:
            issue('duplicate_trade_key', repr(key), 'Trades', idx, 'error')
        trade_keys.add(key)
        if sec != code:
            issue('trade_sector_mismatch', sec, 'Trades', idx, 'error')
        if not origin or not destination or val is None or val < 0:
            issue('invalid_trade', repr(row), 'Trades', idx, 'error')
            continue
        if origin == destination:
            issue('domestic_trade_row', origin, 'Trades', idx)
            continue
        trade_units.add(unit)
        valid_trades.append((origin, destination, val, unit))
        exports[origin] += val
        imports[destination] += val
        for region in (origin, destination):
            if region not in outputs:
                issue('trade_endpoint_without_output', region, 'Trades', idx)

    for region in sorted(set(outputs) | set(exports) | set(imports)):
        x = outputs.get(region, {})
        val = x.get('value')
        comparable = val is not None and bool(trade_units) and trade_units == {x.get('unit')}
        ex, im = exports.get(region, 0), imports.get(region, 0)
        trade_checks.append(dict(stage=stage, file=filename, sector=code, region=region,
                                output=val, output_unit=x.get('unit'), observed_exports=ex,
                                observed_imports=im, trade_unit_labels=';'.join(sorted(trade_units)),
                                units_match=comparable, has_any_trade_rows=bool(valid_trades),
                                exports_exceed_output=(ex > val + max(1e-8, abs(val)*1e-8)) if comparable else '',
                                exports_to_output=ex/val if comparable and val > 0 else '',
                                observed_net_availability=val+im-ex if comparable else '',
                                note='Partial trade coverage; no inference of zero missing flows.'))

    semantic = w.semantic_hash()
    for sn, address, value in w.errors:
        issue('excel_error', f'{address}: {value}', sn, severity='error')
    for sn, address in w.uncached_formulas:
        issue('uncached_formula', address, sn)
    return dict(stage=stage, file=filename, sector=code, parent=parent,
                sector_name=text(w.cell('Summary', 'C2')), version=text(w.cell('Summary', 'B4')),
                sha256=digest(raw), semantic_sha256=semantic, size_bytes=len(raw),
                sheets=len(w.sheets), inventory_sheets=len(actual),
                output_rows=len(outputs), trade_rows=len(valid_trades),
                output_units=';'.join(sorted({o['unit'] for o in outputs.values()})),
                outputs=outputs, trades=valid_trades)


def main():
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument('--published-zip', type=Path, required=True)
    ap.add_argument('--cleaning-dir', type=Path, required=True)
    ap.add_argument('--matching', type=Path, required=True)
    ap.add_argument('--gtap-totals', type=Path, required=True)
    ap.add_argument('--output-dir', type=Path, required=True)
    args = ap.parse_args()
    args.output_dir.mkdir(parents=True, exist_ok=True)
    records, issues, inv_checks, trade_checks, comparison = [], [], [], [], []
    published_files = []
    archive_bytes = args.published_zip.read_bytes()
    published = zipfile.ZipFile(io.BytesIO(archive_bytes))
    for member in published.namelist():
        if member.endswith('.xlsx') and '/' in member and not member.startswith('__MACOSX'):
            print('Published:', Path(member).name, flush=True)
            raw = published.read(member)
            info = dict(file=Path(member).name, member=member, sha256=digest(raw), size_bytes=len(raw))
            try:
                records.append(analyze(raw, Path(member).name, 'published', issues, inv_checks, trade_checks))
                info['status'] = 'parsed'
            except (zipfile.BadZipFile, ET.ParseError, KeyError) as exc:
                info['status'] = 'unreadable'
                info['error'] = f'{type(exc).__name__}: {exc}'
                issues.append(dict(stage='published', file=Path(member).name, sector='',
                                   check='unreadable_workbook', severity='error', sheet='', row='',
                                   detail=info['error']))
            published_files.append(info)
    pub_by_name = {r['file']: r for r in records}
    pub_files_by_name = {r['file']: r for r in published_files}
    for stage, folder in [('workspace_final', 'D2.4 inventories'), ('mario_intermediate', 'MARIO inventories')]:
        for path in sorted((args.cleaning_dir/folder).glob('*.xlsx')):
            if path.name.startswith('~$'):
                continue
            offline = bool(getattr(path.stat(), 'st_flags', 0) & 0x40000000)
            if offline:
                comparison.append(dict(stage=stage, file=path.name, status='offline_not_read',
                                       published_match=path.name in pub_files_by_name))
                continue
            print(stage + ':', path.name, flush=True)
            raw = path.read_bytes()
            if stage == 'workspace_final':
                pub = pub_by_name.get(path.name)
                sha = digest(raw)
                pub_file = pub_files_by_name.get(path.name)
                same = bool(pub_file and sha == pub_file['sha256'])
                try:
                    semantic = pub['semantic_sha256'] if same and pub else Workbook(raw).semantic_hash()
                except (zipfile.BadZipFile, ET.ParseError, KeyError):
                    semantic = None
                comparison.append(dict(stage=stage, file=path.name, status='read',
                                       published_match=pub_file is not None, identical_bytes=same,
                                       identical_values=bool(pub and semantic == pub['semantic_sha256']) if pub else 'unavailable',
                                       local_workbook_readable=semantic is not None,
                                       sha256=sha))
            else:
                records.append(analyze(raw, path.name, stage, issues, inv_checks, trade_checks))

    matching = Workbook(args.matching.read_bytes())
    map_rows = defaultdict(list)
    for idx, row in matching.rows('Sector'):
        if idx > 1 and row.get('B'):
            map_rows[text(row['B'])].append(dict(name=text(row.get('A')), parent=text(row.get('C')),
                                               source_pipeline=text(row.get('E')) or 'template_or_unspecified', row=idx))
    for r in records:
        matches = map_rows.get(r['sector'], [])
        r['matching_rows'] = len(matches)
        r['matching_parents'] = ';'.join(sorted({m['parent'] for m in matches}))
        r['source_pipeline'] = ';'.join(sorted({m['source_pipeline'] for m in matches}))
    groups = defaultdict(list)
    for r in records:
        groups[(r['stage'], r['sector'])].append(r)
    duplicate_codes = [dict(stage=k[0], sector=k[1], files=[r['file'] for r in rows],
                            identical_bytes=len({r['sha256'] for r in rows}) == 1,
                            identical_values=len({r['semantic_sha256'] for r in rows}) == 1)
                       for k, rows in groups.items() if len(rows) > 1]

    totals_book = Workbook(args.gtap_totals.read_bytes())
    gtap = {}
    for idx, row in totals_book.rows('GTAP totals'):
        if idx > 1:
            gtap[(text(row.get('C')), text(row.get('A')))] = number(row.get('E'))
    parent_checks, combined = [], defaultdict(list)
    for r in records:
        for region, entry in r['outputs'].items():
            x, baseline = entry['value'], gtap.get((r['parent'], region))
            parent_checks.append(dict(stage=r['stage'], file=r['file'], sector=r['sector'],
                                      parent=r['parent'], region=region, output=x, parent_output=baseline,
                                      exceeds_parent=x > baseline + max(1e-8, abs(baseline)*1e-8)
                                      if x is not None and baseline is not None else '',
                                      note='Reference workbook year/valuation must be confirmed.'))
            if len(groups[(r['stage'], r['sector'])]) == 1 and x is not None:
                combined[(r['stage'], r['parent'], region)].append((r['sector'], x))
    sibling_checks = []
    ambiguous_parents = {(r['stage'], r['parent']) for r in records
                         if len(groups[(r['stage'], r['sector'])]) > 1}
    for (stage, parent, region), children in combined.items():
        total = sum(x for _, x in children)
        baseline = gtap.get((parent, region))
        sibling_checks.append(dict(stage=stage, parent=parent, region=region,
                                   sum_unambiguous_children=total, parent_output=baseline,
                                   child_codes=';'.join(c for c,_ in children),
                                   duplicate_code_omitted=(stage,parent) in ambiguous_parents,
                                   exceeds_parent=total > baseline + max(1e-8, abs(baseline)*1e-8)
                                   if baseline is not None else ''))

    lineage = []
    for r in [r for r in records if r['stage'] == 'published']:
        sources = groups.get(('mario_intermediate', r['sector']), [])
        if len(sources) != 1:
            lineage.append(dict(sector=r['sector'], file=r['file'], source_count=len(sources)))
            continue
        s = sources[0]
        common = set(s['outputs']) & set(r['outputs'])
        different = [reg for reg in common if s['outputs'][reg] != r['outputs'][reg]]
        # Row positions are presentation, not a numerical difference.
        different = [reg for reg in different if
                     (s['outputs'][reg]['value'], s['outputs'][reg]['unit']) !=
                     (r['outputs'][reg]['value'], r['outputs'][reg]['unit'])]
        lineage.append(dict(sector=r['sector'], file=r['file'], source_count=1, source_file=s['file'],
                            common_output_regions=len(common), changed_output_regions=len(different),
                            removed_output_regions=';'.join(sorted(set(s['outputs'])-set(r['outputs']))),
                            added_output_regions=';'.join(sorted(set(r['outputs'])-set(s['outputs']))),
                            changed_regions=';'.join(sorted(different))))

    short = [{k:v for k,v in r.items() if k not in ('outputs','trades')} for r in records]
    out = args.output_dir
    for name, rows in [('workbooks',short), ('issues',issues), ('inventory_sums',inv_checks),
                       ('trade_output_screening',trade_checks), ('workspace_comparison',comparison),
                       ('parent_output_screening',parent_checks), ('sibling_output_screening',sibling_checks),
                       ('output_lineage',lineage)]:
        write_csv(out/(name+'.csv'), rows)
    pub = [r for r in records if r['stage']=='published']
    summary = dict(audit_date='2026-10-01', executed_at_utc=datetime.now(timezone.utc).isoformat(),
                   published_zip_sha256=digest(archive_bytes),
                   published_zip_md5=hashlib.md5(archive_bytes).hexdigest(),
                   published_workbooks=len(published_files), parsed_published_workbooks=len(pub),
                   published_unique_codes_parsed=len({r['sector'] for r in pub}),
                   unreadable_published_workbooks=[r for r in published_files if r['status']=='unreadable'],
                   duplicate_codes=duplicate_codes,
                   missing_matching_codes=sorted({r['sector'] for r in pub if not r['matching_rows']}),
                   ambiguous_matching_codes=sorted({r['sector'] for r in pub if r['matching_rows']>1}),
                   published_without_trade=[r['sector'] for r in pub if not r['trade_rows']],
                   issue_counts=dict(Counter(i['check'] for i in issues)),
                   published_export_output_flags=sum(r['stage']=='published' and r['exports_exceed_output'] is True for r in trade_checks),
                   published_parent_flags=sum(r['stage']=='published' and r['exceeds_parent'] is True for r in parent_checks),
                   published_sibling_flags=sum(r['stage']=='published' and r['exceeds_parent'] is True for r in sibling_checks),
                   published_sum_flags_1e_3=sum(r['stage']=='published' and r['outside_1e_3'] for r in inv_checks),
                   published_sum_flags_1e_2=sum(r['stage']=='published' and r['outside_1e_2'] for r in inv_checks),
                   workspace_status_counts=dict(Counter(r['status'] for r in comparison)))
    (out/'summary.json').write_text(json.dumps(summary, indent=2)+'\n')
    (out/'published_files.json').write_text(json.dumps(published_files, indent=2)+'\n')
    (out/'workbook_details.json').write_text(json.dumps(records, indent=2)+'\n')
    print(json.dumps(summary, indent=2))


if __name__ == '__main__':
    main()
