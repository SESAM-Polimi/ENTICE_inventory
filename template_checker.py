"""
template_checker.py
====================
Parses an ENTICE data-collection template (Excel v2.x) and performs
consistency checks.

Sheets used:
  'General info'    – project metadata
  'Unit process'    – cost-structure coefficients by producing region
  'Total production'– regional output values
  'Sector'          – admissible sector list and custom clusters
  'Region'          – admissible region list and custom clusters
  'Primary input'   – admissible primary-input list and custom clusters

Sheets discarded:
  'Source', 'Emissions', 'Energy', 'USD deflator'

Public API
----------
  result = parse_template("path/to/Inventory.xlsx")

  result['general_info']      -> dict
  result['unit_process']      -> OrderedDict { region -> [row_dicts] }
  result['total_production']  -> list of dicts
  result['checks']            -> dict of { check_name -> [issue_dicts] }

Each issue dict has keys: row, sheet, col, value, reason.
"""

from __future__ import annotations

import openpyxl
from collections import OrderedDict
from pathlib import Path


# ─────────────────────────────────────────────────────────────────────────────
# Helpers
# ─────────────────────────────────────────────────────────────────────────────

def _s(v) -> str | None:
    """Strip a cell value to str, or return None."""
    return str(v).strip() if v is not None else None

def _expand_region_ids_to_gtap_names(region_ids: set,
                                     cluster2members: dict,
                                     all_gtap_names: set) -> set:
    """
    Expand template region identifiers to GTAP full names.

    Direct GTAP names are kept as-is; valid cluster identifiers expand to
    their GTAP members. GLOBAL is handled separately by callers so it can
    remain a dedicated fallback category in reports.
    """
    covered: set = set()
    for region_id in region_ids:
        if region_id in all_gtap_names:
            covered.add(region_id)
            continue

        for member in cluster2members.get(region_id, []):
            if member in all_gtap_names:
                covered.add(member)

    return covered


def _issue(row: int, sheet: str, col: str, value: str, reason: str) -> dict:
    return dict(row=row, sheet=sheet, col=col, value=value, reason=reason)


def _find_tp_columns(ws) -> tuple[int, int]:
    """Find (val_col, unit_col) in row 1 of 'Total production' sheet."""
    col_val  = 7
    col_unit = 8
    for c in range(1, ws.max_column + 1):
        h = _s(ws.cell(1, c).value)
        if h:
            h_low = h.lower()
            if h_low == '2023 value':
                col_val = c
            elif h_low == '2023 unit':
                col_unit = c
    return col_val, col_unit


# ─────────────────────────────────────────────────────────────────────────────
# 1. General info
# ─────────────────────────────────────────────────────────────────────────────

def read_general_info(wb) -> dict:
    """
    Read from 'General info' sheet:
      D5 -> new_sector   (name of the new sector)
      D6 -> parent_name  (long name of the parent GTAP sector)
      D7 -> parent_code  (short GTAP code of the parent sector)
    """
    ws = wb['General info']
    return {
        'new_sector':  _s(ws.cell(5, 4).value),
        'parent_name': _s(ws.cell(6, 4).value),
        'parent_code': _s(ws.cell(7, 4).value),
    }


# ─────────────────────────────────────────────────────────────────────────────
# 2. Unit process
# ─────────────────────────────────────────────────────────────────────────────

_IGNORE_TYPES = {'Emissions', 'Energy'}

# Tolerance for the unit-process column-sum check (must be within this of 1.0)
_TOL_SUM = 1e-3


def read_unit_process(wb) -> tuple[OrderedDict, list]:
    """
    Parse 'Unit process' and return (data, warnings).

    Layout:
      Row 1  – region names from col G (col 7) onward, one name every
                2 columns (the odd-numbered cols hold Values, the even ones
                hold Units which are discarded here).
      Row 2  – unit type per column: only columns marked 'monetary'
                (case-insensitive) are retained; the others generate a warning
                and are skipped entirely.
      Row 4  – column headers (discarded)
      Row 6+ – data rows
                  col A (1) = Input type
                  col B (2) = Input label          [discarded]
                  col C (3) = Mapped Input
                  col D (4) = Input origin
                  col E (5) = Source               [discarded]
                  col F (6) = Notes                [discarded]
                  col G, I, K, … = value for each region

    Rows whose type is 'Emissions' or 'Energy' are skipped entirely.

    After reading, the sum of all numeric values in each column is checked.
    Columns whose sum differs from 1.0 by more than _TOL_SUM are discarded,
    except columns whose sum is 0 and whose retained values are all zeroes.
    Those are kept with a warning.

    Returns:
      data     : OrderedDict { region_name -> list of {
                     row, type, mapped_input, origin, value } }
                 (only columns that pass both filters)
      warnings : list of issue dicts for every skipped column
    """
    ws = wb['Unit process']
    warnings: list = []

    # ── Discover producing regions from row 1 (col G=7 onward, step 2) ──────
    # Filter: row 2 must say 'monetary' (case-insensitive).
    region_cols: list[tuple[int, str]] = []
    c = 7
    while c <= ws.max_column:
        name_val = ws.cell(1, c).value
        if name_val is not None and str(name_val).strip():
            region_name = str(name_val).strip()
            unit_val    = ws.cell(2, c).value
            unit_str    = str(unit_val).strip() if unit_val is not None else ''
            if unit_str.lower() == 'monetary':
                region_cols.append((c, region_name))
            else:
                warnings.append(_issue(
                    2, 'Unit process', f'col {c}', region_name,
                    f"row-2 unit type is '{unit_str or '(blank)'}', "
                    "not 'monetary'; column skipped"
                ))
        c += 2

    # ── Read data rows for the accepted (monetary) columns ───────────────────
    data: OrderedDict = OrderedDict((reg, []) for _, reg in region_cols)

    for r in range(6, ws.max_row + 1):
        raw_type = ws.cell(r, 1).value
        if raw_type is None:
            continue
        input_type = str(raw_type).strip()
        if input_type in _IGNORE_TYPES:
            continue

        mapped_input = _s(ws.cell(r, 3).value)
        origin       = _s(ws.cell(r, 4).value) or 'GLOBAL'
        if not mapped_input:
            warnings.append(_issue(
                r, 'Unit process', 'C', '',
                'missing mapped input; row skipped'
            ))
            continue

        for col_idx, region in region_cols:
            raw_val = ws.cell(r, col_idx).value
            if raw_val is None:
                continue
            try:
                value = float(raw_val)
            except (ValueError, TypeError):
                continue
            data[region].append({
                'row':          r,
                'type':         input_type,
                'mapped_input': mapped_input,
                'origin':       origin,
                'value':        value,
            })

    # ── Filter: column sum must be within _TOL_SUM of 1.0 ───────────────────
    valid: OrderedDict = OrderedDict()
    for col_idx, region_name in region_cols:
        rows    = data[region_name]
        col_sum = sum(row['value'] for row in rows)
        all_zero = bool(rows) and all(abs(row['value']) <= _TOL_SUM for row in rows)
        if abs(col_sum - 1.0) <= _TOL_SUM:
            valid[region_name] = rows
        elif abs(col_sum) <= _TOL_SUM and all_zero:
            valid[region_name] = rows
            warnings.append(_issue(
                1, 'Unit process', f'col {col_idx}', region_name,
                f"column values sum to {col_sum:.6g} (expected 1.0 ± {_TOL_SUM}); "
                "kept because all retained values are zero"
            ))
        elif str(region_name).strip().upper() == 'GLOBAL':
            valid[region_name] = rows
            warnings.append(_issue(
                1, 'Unit process', f'col {col_idx}', region_name,
                f"column values sum to {col_sum:.6g} "
                f"(expected 1.0 ± {_TOL_SUM}); GLOBAL kept as fallback inventory"
            ))
        else:
            warnings.append(_issue(
                1, 'Unit process', f'col {col_idx}', region_name,
                f"column values sum to {col_sum:.6g} "
                f"(expected 1.0 ± {_TOL_SUM}); column skipped"
            ))

    return valid, warnings


# ─────────────────────────────────────────────────────────────────────────────
# 3. Total production
# ─────────────────────────────────────────────────────────────────────────────

def read_total_production(wb) -> list:
    """
    Parse 'Total production' (rows 3+).

    Column layout (row 1 headers):
      A (1) = Region
      ...
      Dynamic lookup for '2023 Value' and '2023 Unit' headers.
      Defaulting to G (7) and H (8) if not found.

    Returns list of { region, value, unit }.
    """
    ws = wb['Total production']
    col_val, col_unit = _find_tp_columns(ws)

    rows = []
    for r in range(3, ws.max_row + 1):
        region = _s(ws.cell(r, 1).value)
        if region is None:
            continue

        raw_val = ws.cell(r, col_val).value
        try:
            val_f = float(raw_val) if raw_val is not None else 0.0
        except (ValueError, TypeError):
            val_f = 0.0

        if region.strip().upper() == 'GLOBAL' and val_f != 0.0:
            continue

        unit = _s(ws.cell(r, col_unit).value)
        rows.append({'region': region, 'value': val_f, 'unit': unit})
    return rows


# ─────────────────────────────────────────────────────────────────────────────
# 4. Reference sets  (built once, shared across all checks)
# ─────────────────────────────────────────────────────────────────────────────

def _build_refs(wb) -> dict:
    """
    Pre-build all lookup sets used by the consistency checks.

    Sector sheet
      col A rows 3+         -> all admissible sector identifiers
                               (NACE codes, sector descriptions, cluster names)
      col B rows 3-185      -> canonical GTAP sector full names

    Region sheet
      col A rows 3+         -> all admissible region identifiers
                               (UN clusters, country full names, custom names)
      col B rows 3-165      -> canonical GTAP region full names

    Primary input sheet
      col A rows 3-16       -> base admissible primary-input names
                               (used for cluster-membership checks)
      col A rows 3+         -> all admissible names including cluster names
                               (used for unit-process col-C checks)
    """
    ws_s = wb['Sector']
    ws_r = wb['Region']
    ws_p = wb['Primary input']

    return {
        'sector_col_a':       {_s(ws_s.cell(r, 1).value)
                               for r in range(3, ws_s.max_row + 1)
                               if ws_s.cell(r, 1).value is not None},

        'sector_col_b_3_185': {_s(ws_s.cell(r, 2).value)
                               for r in range(3, 186)
                               if ws_s.cell(r, 2).value is not None},

        'region_col_a':       {_s(ws_r.cell(r, 1).value)
                               for r in range(3, ws_r.max_row + 1)
                               if ws_r.cell(r, 1).value is not None},

        # build region col B set using rows 3..last where 'last' is the
        # last non-empty cell in column D (4). This avoids hard-coding 165.
        'region_col_b_3_165': (lambda ws: (
            (lambda last: {_s(ws.cell(r, 2).value)
                           for r in range(3, last + 1)
                           if ws.cell(r, 2).value is not None}
            )( (lambda _ws: (lambda last: last)(next((i for i in range(_ws.max_row, 2, -1)
                                                       if _ws.cell(i, 4).value not in (None, '')), 2)) ) (ws_r) )
        ))(ws_r),

        # rows 3-16 only: for validating cluster-member entries (rows 17+)
        'pi_col_a_3_16':      {_s(ws_p.cell(r, 1).value)
                               for r in range(3, 17)
                               if ws_p.cell(r, 1).value is not None},

        # full column A: for validating Unit process col C (includes cluster names)
        'pi_col_a_all':       {_s(ws_p.cell(r, 1).value)
                               for r in range(3, ws_p.max_row + 1)
                               if ws_p.cell(r, 1).value is not None},
    }


# ─────────────────────────────────────────────────────────────────────────────
# 5. Consistency checks
# ─────────────────────────────────────────────────────────────────────────────

def check_sector_clusters(wb, refs: dict) -> list:
    """
    Sector sheet, rows 369+:
    col B values must be found in col B rows 3-185
    (i.e. every member of a custom cluster must be a known GTAP sector name).
    """
    ws    = wb['Sector']
    valid = refs['sector_col_b_3_185']
    issues = []
    for r in range(369, ws.max_row + 1):
        v = _s(ws.cell(r, 2).value)
        if v and v not in valid:
            issues.append(_issue(r, 'Sector', 'B', v,
                                 'not a valid GTAP sector name (Sector col B rows 3-185)'))
    return issues


def check_region_clusters(wb, refs: dict) -> list:
    """
    Region sheet, rows 328+:
    col B values must be found in col B rows 3-165
    (i.e. every member of a custom cluster must be a known GTAP region full name).
    """
    ws    = wb['Region']
    valid = refs['region_col_b_3_165']
    issues = []
    for r in range(328, ws.max_row + 1):
        v = _s(ws.cell(r, 2).value)
        if v and v not in valid:
            issues.append(_issue(r, 'Region', 'B', v,
                                 'not a valid GTAP region full name (Region col B rows 3-165)'))
    return issues


def check_primary_input_clusters(wb, refs: dict) -> list:
    """
    Primary input sheet, rows 17+:
    col B values must be found in col A rows 3-16
    (i.e. every member of a custom cluster must be a known primary-input name).
    """
    ws    = wb['Primary input']
    valid = refs['pi_col_a_3_16']
    issues = []
    for r in range(17, ws.max_row + 1):
        v = _s(ws.cell(r, 2).value)
        if v and v not in valid:
            issues.append(_issue(r, 'Primary input', 'B', v,
                                 'not a valid primary-input name (Primary input col A rows 3-16)'))
    return issues


def check_unit_process_region_names(wb, refs: dict) -> list:
    """
    Unit process, row 1:
    Region names (col G+, step 2) must be valid GTAP region identifiers,
    i.e. present in Region col B rows 3-165 (canonical full names) or in
    Region col A rows 3+ (UN/custom cluster names and country variants).
    """
    ws    = wb['Unit process']
    valid = refs['region_col_b_3_165'] | refs['region_col_a']
    issues = []
    c = 7
    while c <= ws.max_column:
        v = _s(ws.cell(1, c).value)
        if v and v not in valid:
            issues.append(_issue(
                1, 'Unit process', f'col {c}', v,
                'not a valid GTAP region name or known cluster '
                '(check Region sheet col A / col B rows 3-165)'
            ))
        c += 2
    return issues


def check_unit_process_inputs(wb, refs: dict) -> list:
    """
    Unit process, rows 6+:
      'Sector'        rows: col C must be in Sector col A (rows 3+)
      'Primary_input' rows: col C must be in Primary input col A (all rows 3+,
                            including cluster names defined from row 17 onward)
    """
    ws       = wb['Unit process']
    valid_s  = refs['sector_col_a']
    valid_pi = refs['pi_col_a_all']
    issues   = []

    for r in range(6, ws.max_row + 1):
        raw_type = ws.cell(r, 1).value
        if raw_type is None:
            continue
        input_type = str(raw_type).strip()
        v = _s(ws.cell(r, 3).value)
        if not v:
            continue

        if input_type == 'Sector' and v not in valid_s:
            issues.append(_issue(r, 'Unit process', 'C', v,
                                 "Sector row: not found in Sector col A"))

        elif input_type == 'Primary_input' and v not in valid_pi:
            issues.append(_issue(r, 'Unit process', 'C', v,
                                 "Primary_input row: not found in Primary input col A rows 3-16"))

    return issues


def check_total_production_regions(wb, refs: dict) -> list:
    """
    Total production, rows 3+:
    col A values must be found in Region col A (rows 3+).
    Skip rows where '2023 Value' is null.
    """
    ws    = wb['Total production']
    valid = refs['region_col_a']
    col_val, _ = _find_tp_columns(ws)
    issues = []
    for r in range(3, ws.max_row + 1):
        v = _s(ws.cell(r, 1).value)
        if not v:
            continue

        if v.strip().upper() == 'GLOBAL':
            continue

        # Skip rows without a numerical value in '2023 Value'
        raw_val = ws.cell(r, col_val).value
        try:
            val_f = float(raw_val) if raw_val is not None else 0.0
        except (ValueError, TypeError):
            val_f = 0.0

        if v not in valid:
            issues.append(_issue(r, 'Total production', 'A', v,
                                 'not found in Region col A'))
    return issues


# ─────────────────────────────────────────────────────────────────────────────
# 6. Missing-region analysis
# ─────────────────────────────────────────────────────────────────────────────

def _build_region_table(wb) -> list:
    """
    Build a list of all 162 GTAP regions (Region sheet rows 3-165).

    Returns:
      list of { full_name, code, clusters: [cluster_name, ...] }

    Cluster memberships come from two sources:
      - rows 3-165, col A : UN regional-aggregate names (e.g. CARIBBEAN,
                            WESTERN ASIA) repeated for every member country.
      - rows 328+, col A  : custom cluster names (e.g. EU28); col A is
                            carried forward when blank (merged / implicit).
    """
    ws_r = wb['Region']

    # ── GTAP base regions (rows 3..last where column D is non-empty) ──────
    regions = []
    last = ws_r.max_row
    while last >= 3 and ws_r.cell(last, 4).value in (None, ''):
        last -= 1
    for r in range(3, last + 1):
        fn   = _s(ws_r.cell(r, 2).value)   # col B = full name
        code = _s(ws_r.cell(r, 3).value)   # col C = code
        if fn:
            regions.append({'full_name': fn, 'code': code, 'clusters': []})

    name_to_idx = {reg['full_name']: i for i, reg in enumerate(regions)}

    # ── UN cluster assignments from col A, rows 3..last
    # Col A holds the UN regional-aggregate name for each country row
    # (e.g. CARIBBEAN, WESTERN ASIA).  Col B is the GTAP full name of the
    # country, already indexed in name_to_idx.
    for r in range(3, last + 1):
        cn     = _s(ws_r.cell(r, 1).value)   # col A = UN cluster name
        member = _s(ws_r.cell(r, 2).value)   # col B = GTAP full name
        if cn and member and member in name_to_idx:
            idx = name_to_idx[member]
            if cn not in regions[idx]['clusters']:
                regions[idx]['clusters'].append(cn)

    # ── Custom cluster assignments from rows 328+ ─────────────────────────
    current_cluster = None
    for r in range(328, ws_r.max_row + 1):
        cn = _s(ws_r.cell(r, 1).value)
        if cn:
            current_cluster = cn
        member = _s(ws_r.cell(r, 2).value)
        if current_cluster and member and member in name_to_idx:
            idx = name_to_idx[member]
            if current_cluster not in regions[idx]['clusters']:
                regions[idx]['clusters'].append(current_cluster)

    return regions


def check_missing_regions(wb) -> dict:
    """
    For each of the 162 GTAP regions determine whether it is covered in:
      - Unit process  (region headers in row 1, col G+, step 2)
      - Total production (col A rows 3+)

    Coverage is either *direct* (the region full name appears explicitly)
    or *via cluster* (the region belongs to a cluster that appears).

    If 'GLOBAL' appears as an identifier it is treated as a catch-all:
    every region not otherwise listed is considered covered by GLOBAL,
    and those regions are reported separately.

    Returns:
      region_table        – list of {full_name, code, clusters}
      up_ids              – sorted list of identifiers in UP row 1
      tp_ids              – sorted list of identifiers in TP col A
      has_global_up       – True if 'GLOBAL' present in UP row 1
      has_global_tp       – True if 'GLOBAL' present in TP col A
      missing_both        – regions absent from both sheets (even with GLOBAL)
      missing_up_only     – in TP but not UP  (even accounting for GLOBAL)
      missing_tp_only     – in UP but not TP  (even accounting for GLOBAL)
      global_covered_up   – not directly in UP, but covered by GLOBAL
      global_covered_tp   – not directly in TP, but covered by GLOBAL
    """
    region_table = _build_region_table(wb)

    # ── Identifiers used in Unit process row 1 (col G+, step 2) ──────────
    ws_up = wb['Unit process']
    up_ids_raw: set = set()
    c = 7
    while c <= ws_up.max_column:
        v = _s(ws_up.cell(1, c).value)
        if v:
            up_ids_raw.add(v)
        c += 2

    # ── Identifiers used in Total production col A (rows 3+) ─────────────
    ws_tp = wb['Total production']
    col_val_tp, _ = _find_tp_columns(ws_tp)
    tp_ids_raw: set = set()
    has_global_tp = False
    for r in range(3, ws_tp.max_row + 1):
        v = _s(ws_tp.cell(r, 1).value)
        if v:
            # Check value in '2023 Value'
            raw_val = ws_tp.cell(r, col_val_tp).value
            try:
                val_f = float(raw_val) if raw_val is not None else 0.0
            except (ValueError, TypeError):
                val_f = 0.0

            if val_f == 0.0:
                continue

            tp_ids_raw.add(v)

    has_global_up = 'GLOBAL' in up_ids_raw
    has_global_tp = 'GLOBAL' in tp_ids_raw

    # Work with GLOBAL removed so direct-coverage check is unambiguous
    up_ids = up_ids_raw - {'GLOBAL'}
    tp_ids = tp_ids_raw - {'GLOBAL'}
    cluster2members = _build_cluster_to_members(wb)
    all_gtap_names = {reg['full_name'] for reg in region_table}
    covered_up_names = _expand_region_ids_to_gtap_names(
        up_ids, cluster2members, all_gtap_names
    )
    covered_tp_names = _expand_region_ids_to_gtap_names(
        tp_ids, cluster2members, all_gtap_names
    )

    # ── Check coverage for each GTAP region ──────────────────────────────
    missing_both  = []
    missing_up    = []
    missing_tp    = []
    global_cov_up = []
    global_cov_tp = []

    for reg in region_table:
        fn = reg['full_name']

        in_up_direct = fn in covered_up_names
        in_tp_direct = fn in covered_tp_names

        # GLOBAL provides a fallback for anything not directly covered
        in_up = in_up_direct or has_global_up
        in_tp = in_tp_direct or has_global_tp

        if not in_up_direct and has_global_up:
            global_cov_up.append(reg)
        if not in_tp_direct and has_global_tp:
            global_cov_tp.append(reg)

        if not in_up and not in_tp:
            missing_both.append(reg)
        elif not in_up:
            missing_up.append(reg)
        elif not in_tp:
            missing_tp.append(reg)

    return {
        'region_table':      region_table,
        'up_ids':            sorted(up_ids_raw),
        'tp_ids':            sorted(tp_ids_raw),
        'covered_up_count':  len(region_table) if has_global_up else len(covered_up_names),
        'covered_tp_count':  len(region_table) if has_global_tp else len(covered_tp_names),
        'has_global_up':     has_global_up,
        'has_global_tp':     has_global_tp,
        'missing_both':      missing_both,
        'missing_up_only':   missing_up,
        'missing_tp_only':   missing_tp,
        'global_covered_up': global_cov_up,
        'global_covered_tp': global_cov_tp,
    }


# ─────────────────────────────────────────────────────────────────────────────
# 7. GTAP production cap check
# ─────────────────────────────────────────────────────────────────────────────

def _load_gtap_x(gtap_path) -> dict:
    """
    Load GTAP12_X.xlsx ('GTAP totals' sheet) and return:

      { (region_full_name, sector_code): X_value_MUSD@2023 }

    Column layout:
      A = region code (3 letters)     – not used for matching
      B = region full name            – used as lookup key
      C = sector code (3 letters)     – used as lookup key
      D = sector full name            – ignored
      E = X value (MUSD@2023)
    """
    wb_g = openpyxl.load_workbook(str(gtap_path), data_only=True)
    ws_g = wb_g['GTAP totals']
    gtap_x: dict = {}
    for r in range(2, ws_g.max_row + 1):
        fullname = _s(ws_g.cell(r, 2).value)
        sector   = _s(ws_g.cell(r, 3).value)
        x        = ws_g.cell(r, 5).value
        if fullname and sector and x is not None:
            try:
                gtap_x[(fullname, sector)] = float(x)
            except (ValueError, TypeError):
                pass
    return gtap_x


def _build_cluster_to_members(wb) -> dict:
    """
    { cluster_name -> [member_full_names] }  from Region sheet.

    Sources:
      rows 3-165  : col A = UN aggregate name, col B = GTAP full name.
                    Every country row carries the name of its UN sub-region.
      rows 328+   : col A = custom cluster name (carry-forward when blank),
                    col B = GTAP full name 

    Both sets are merged so that _resolve_to_names can handle UN aggregate
    identifiers (e.g. CARIBBEAN) as well as user-defined clusters (e.g. EU28).
    """
    ws_r = wb['Region']
    clusters: dict = {}

    # UN regional aggregates (rows 3..last where column D is non-empty)
    last = ws_r.max_row
    while last >= 3 and ws_r.cell(last, 4).value in (None, ''):
        last -= 1
    for r in range(3, last + 1):
        cn     = _s(ws_r.cell(r, 1).value)
        member = _s(ws_r.cell(r, 2).value)
        if cn and member:
            clusters.setdefault(cn, [])
            if member not in clusters[cn]:
                clusters[cn].append(member)

    # Custom clusters (rows 328+)
    current = None
    for r in range(328, ws_r.max_row + 1):
        cn = _s(ws_r.cell(r, 1).value)
        if cn:
            current = cn
        member = _s(ws_r.cell(r, 2).value)
        if current and member:
            clusters.setdefault(current, [])
            if member not in clusters[current]:
                clusters[current].append(member)

    return clusters


def _resolve_to_names(region_id: str,
                      cluster2members: dict,
                      all_gtap_names: set) -> tuple[list, str]:
    """
    Resolve a Total-production region identifier to a list of GTAP full names
    (matching col B of GTAP12_X).

    Returns (names_list, status) where status is:
      'ok'         – one or more GTAP names found
      'unresolved' – not a known GTAP name or cluster
                     (e.g. countries rolled into "Rest of..." in GTAP 12)
    """
    rid = region_id.strip()

    # 1. GLOBAL → all GTAP region names
    if rid.upper() == 'GLOBAL':
        return sorted(all_gtap_names), 'ok'

    # 2. Direct match against GTAP full names (col B)
    if rid in all_gtap_names:
        return [rid], 'ok'

    # 3. Cluster defined in the template (e.g. EU28)
    if rid in cluster2members:
        members = [m for m in cluster2members[rid] if m in all_gtap_names]
        if members:
            return members, 'ok'

    return [], 'unresolved'


def check_total_production_vs_gtap(wb, gtap_path) -> dict:
    """
    Compare each Total-production row (Inventory col G, MUSD@2023) against the
    corresponding parent-sector output in GTAP12_X.xlsx (col E, MUSD@2023).

    Region matching
    ---------------
    Template col A identifiers are matched directly against GTAP12_X col B
    (full region names).  For cluster identifiers (e.g. EU28), the GTAP X
    values of all member regions are summed.  Rows whose identifier cannot be
    resolved to any GTAP name (countries aggregated into "Rest of..." in
    GTAP 12) are reported as 'unresolved' without raising a violation.

    Returns
    -------
    {
      'violations' : [ {row, region, members, tp_val, gtap_sum, ratio} ]
                     rows where col-G value > GTAP parent-sector sum
      'unresolved' : [ {row, region} ]
                     rows whose region could not be mapped to GTAP names
      'gtap_missing': [ {row, region, members} ]
                     region mapped to GTAP names but parent sector has no
                     GTAP X entry (sum = 0)
    }
    """
    parent_code     = read_general_info(wb)['parent_code']
    gtap_x          = _load_gtap_x(gtap_path)
    all_gtap_names  = {fn for (fn, _) in gtap_x}
    cluster2members = _build_cluster_to_members(wb)

    ws_tp = wb['Total production']
    col_val, _ = _find_tp_columns(ws_tp)

    violations   = []
    unresolved   = []
    gtap_missing = []

    for r in range(3, ws_tp.max_row + 1):
        region_id = _s(ws_tp.cell(r, 1).value)
        if not region_id:
            continue

        if region_id.strip().upper() == 'GLOBAL':
            continue

        # Template value: MUSD@2023
        raw = ws_tp.cell(r, col_val).value
        if raw is None:
            continue
        try:
            tp_val = float(raw)
        except (ValueError, TypeError):
            continue

        if tp_val == 0.0:
            continue

        # Resolve region identifier to GTAP full names
        names, status = _resolve_to_names(region_id, cluster2members, all_gtap_names)

        if status == 'unresolved':
            unresolved.append({'row': r, 'region': region_id})
            continue

        # Sum GTAP X for the parent sector across all matched region names
        gtap_sum = sum(gtap_x.get((n, parent_code), 0.0) for n in names)

        if gtap_sum == 0.0:
            gtap_missing.append({'row': r, 'region': region_id, 'members': names})
            continue

        if tp_val > gtap_sum:
            violations.append({
                'row':      r,
                'region':   region_id,
                'members':  names,
                'tp_val':   tp_val,
                'gtap_sum': gtap_sum,
                'ratio':    tp_val / gtap_sum,
            })

    return {
        'violations':   violations,
        'unresolved':   unresolved,
        'gtap_missing': gtap_missing,
    }


# ─────────────────────────────────────────────────────────────────────────────
# 8. Top-level entry point
# ─────────────────────────────────────────────────────────────────────────────

def parse_template(path: str | Path, gtap_path: str | Path = None) -> dict:
    """
    Load the workbook at *path* and return a result dict:

      {
        'general_info':     { new_sector, parent_name, parent_code },
        'unit_process':     OrderedDict { region -> [row_dicts] },
        'total_production': [ { region, value, unit } ],
        'checks': {
            'sector_clusters':              [issues],
            'region_clusters':              [issues],
            'primary_input_clusters':       [issues],
            'unit_process_region_names':    [issues],   ← new
            'unit_process_inputs':          [issues],
            'unit_process_columns':         [issues],   ← new (non-monetary / bad sum)
            'total_production_regions':     [issues],
            'total_production_vs_gtap': { violations, unresolved, gtap_missing }
                                         (only present when gtap_path is given)
        },
        'missing_regions': { region_table, up_ids, tp_ids, has_global_up,
                             has_global_tp, missing_both, missing_up_only,
                             missing_tp_only, global_covered_up,
                             global_covered_tp },
      }

    gtap_path : optional path to GTAP12_X.xlsx; when supplied, the
                'total_production_vs_gtap' check is included.
    """
    wb   = openpyxl.load_workbook(str(path), data_only=True)
    refs = _build_refs(wb)

    # read_unit_process now returns (filtered_data, column_warnings)
    up_data, up_col_warnings = read_unit_process(wb)

    checks = {
        'sector_clusters':           check_sector_clusters(wb, refs),
        'region_clusters':           check_region_clusters(wb, refs),
        'primary_input_clusters':    check_primary_input_clusters(wb, refs),
        'unit_process_region_names': check_unit_process_region_names(wb, refs),
        'unit_process_inputs':       check_unit_process_inputs(wb, refs),
        'unit_process_columns':      up_col_warnings,
        'total_production_regions':  check_total_production_regions(wb, refs),
    }

    if gtap_path is not None:
        checks['total_production_vs_gtap'] = check_total_production_vs_gtap(
            wb, Path(gtap_path)
        )

    return {
        'general_info':     read_general_info(wb),
        'unit_process':     up_data,
        'total_production': read_total_production(wb),
        'checks':           checks,
        'missing_regions':  check_missing_regions(wb),
    }
