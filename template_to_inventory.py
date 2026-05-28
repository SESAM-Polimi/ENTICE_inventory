"""
template_to_inventory.py  (v3)
==============================
Converts an ENTICE data-collection template (Excel v2.x) into a MARIO
add_sector inventory, using template_checker.parse_template() as the
data source.

API
---
  from template_to_inventory import make_inventory
  make_inventory("Inventory_Cultivation of barley.xlsx",
                 sector_code="BAR", version="Y26M05")

CLI
---
  python template_to_inventory.py <template.xlsx> --sector-code BAR --version Y26M05

Sheets produced in the inventory:
  Summary                     ← sector metadata, sources, page links
  Master
  Regions Clusters
  Sectors Clusters
  Factors Clusters            ← Primary input clusters
  {sector_code}_{region}      ← one per producing region in Unit process
  DB units                    ← left unchanged
  Total outputs

"""

import argparse
import importlib
import openpyxl
import statistics
import template_checker
from collections import OrderedDict
from pathlib import Path


# ─────────────────────────────────────────────────────────────────────────────
# DB units  (unchanged — leave as-is)
# ─────────────────────────────────────────────────────────────────────────────

DB_UNITS_DATA = [
    (None,                  None,         'unit'),
    ('Sector',              'AFS',        'M USD'),
    ('Sector',              'ATP',        'M USD'),
    ('Sector',              'B_T',        'M USD'),
    ('Sector',              'BPH',        'M USD'),
    ('Sector',              'C_B',        'M USD'),
    ('Sector',              'CHM',        'M USD'),
    ('Sector',              'CMN',        'M USD'),
    ('Sector',              'CMT',        'M USD'),
    ('Sector',              'CNS',        'M USD'),
    ('Sector',              'COA',        'M USD'),
    ('Sector',              'CoalBL',     'M USD'),
    ('Sector',              'CTL',        'M USD'),
    ('Sector',              'DWE',        'M USD'),
    ('Sector',              'EDU',        'M USD'),
    ('Sector',              'EEQ',        'M USD'),
    ('Sector',              'ELE',        'M USD'),
    ('Sector',              'FMP',        'M USD'),
    ('Sector',              'FRS',        'M USD'),
    ('Sector',              'FSH',        'M USD'),
    ('Sector',              'GAS',        'M USD'),
    ('Sector',              'GasBL',      'M USD'),
    ('Sector',              'GasP',       'M USD'),
    ('Sector',              'GDT',        'M USD'),
    ('Sector',              'GRO',        'M USD'),
    ('Sector',              'HHT',        'M USD'),
    ('Sector',              'HydroBL',    'M USD'),
    ('Sector',              'HydroP',     'M USD'),
    ('Sector',              'I_S',        'M USD'),
    ('Sector',              'INS',        'M USD'),
    ('Sector',              'LEA',        'M USD'),
    ('Sector',              'LUM',        'M USD'),
    ('Sector',              'MIL',        'M USD'),
    ('Sector',              'MVH',        'M USD'),
    ('Sector',              'NFM',        'M USD'),
    ('Sector',              'NMM',        'M USD'),
    ('Sector',              'NuclearBL',  'M USD'),
    ('Sector',              'OAP',        'M USD'),
    ('Sector',              'OBS',        'M USD'),
    ('Sector',              'OCR',        'M USD'),
    ('Sector',              'OFD',        'M USD'),
    ('Sector',              'OFI',        'M USD'),
    ('Sector',              'OIL',        'M USD'),
    ('Sector',              'OilBL',      'M USD'),
    ('Sector',              'OilP',       'M USD'),
    ('Sector',              'OME',        'M USD'),
    ('Sector',              'OMF',        'M USD'),
    ('Sector',              'OMT',        'M USD'),
    ('Sector',              'OSD',        'M USD'),
    ('Sector',              'OSG',        'M USD'),
    ('Sector',              'OtherBL',    'M USD'),
    ('Sector',              'OTN',        'M USD'),
    ('Sector',              'OTP',        'M USD'),
    ('Sector',              'OXT',        'M USD'),
    ('Sector',              'P_C',        'M USD'),
    ('Sector',              'PCR',        'M USD'),
    ('Sector',              'PDR',        'M USD'),
    ('Sector',              'PFB',        'M USD'),
    ('Sector',              'PPP',        'M USD'),
    ('Sector',              'RMK',        'M USD'),
    ('Sector',              'ROS',        'M USD'),
    ('Sector',              'RPP',        'M USD'),
    ('Sector',              'RSA',        'M USD'),
    ('Sector',              'SGR',        'M USD'),
    ('Sector',              'SolarP',     'M USD'),
    ('Sector',              'TEX',        'M USD'),
    ('Sector',              'TnD',        'M USD'),
    ('Sector',              'TRD',        'M USD'),
    ('Sector',              'V_F',        'M USD'),
    ('Sector',              'VOL',        'M USD'),
    ('Sector',              'WAP',        'M USD'),
    ('Sector',              'WHS',        'M USD'),
    ('Sector',              'WHT',        'M USD'),
    ('Sector',              'WindBL',     'M USD'),
    ('Sector',              'WOL',        'M USD'),
    ('Sector',              'WTP',        'M USD'),
    ('Sector',              'WTR',        'M USD'),
    ('Factor of production', 'Tax',       'M USD'),
    ('Factor of production', 'TTM',       'M USD'),
    ('Factor of production', 'Capital',   'M USD'),
    ('Factor of production', 'Land',      'M USD'),
    ('Factor of production', 'NatRes',    'M USD'),
    ('Factor of production', 'Labor',     'M USD'),
    ('Satellite account',    'BC',        'M ton'),
    ('Satellite account',    'CO2',       'M ton'),
    ('Satellite account',    'CO',        'M ton'),
    ('Satellite account',    'NH3',       'M ton'),
    ('Satellite account',    'NMVOC',     'M ton'),
    ('Satellite account',    'NOX',       'M ton'),
    ('Satellite account',    'OC',        'M ton'),
    ('Satellite account',    'PM10',      'M ton'),
    ('Satellite account',    'PM2.5',     'M ton'),
    ('Satellite account',    'SO2',       'M ton'),
]


# ─────────────────────────────────────────────────────────────────────────────
# Internal helpers
# ─────────────────────────────────────────────────────────────────────────────

def _name_to_code_map(wb) -> dict:
    """
    Build {region_full_name -> 3-letter code} from Region sheet
    col B (full name) → col C (code), rows 3-165.

    Used to convert full region names from Unit process row 1 into
    compact codes for sheet names and Master rows.
    Cluster names that are not in this table are returned unchanged.
    """
    ws = wb['Region']
    m = {}
    for r in range(3, 166):
        fn   = ws.cell(r, 2).value   # col B = full name
        code = ws.cell(r, 3).value   # col C = code
        if fn and code:
            m[str(fn).strip()] = str(code).strip()
    return m


def _safe_sheet_name(name: str) -> str:
    """Return a valid Excel sheet name (max 31 chars, no \ / * ? : [ ])."""
    for ch in r'\/*?:[]':
        name = name.replace(ch, '_')
    return name[:31]


def _read_clusters(ws, start_row: int) -> OrderedDict:
    """
    Read cluster definitions starting at *start_row* of worksheet *ws*.

    Layout (shared by Region, Sector, and Primary input sheets):
      col A = cluster name (carried forward when blank, i.e. merged cells)
      col B = member

    Returns OrderedDict {cluster_name: [members]}  (members deduplicated,
    order preserved).
    """
    clusters: OrderedDict = OrderedDict()
    current = None
    for r in range(start_row, ws.max_row + 1):
        raw_cn = ws.cell(r, 1).value
        if raw_cn and str(raw_cn).strip():
            current = str(raw_cn).strip()
        raw_mb = ws.cell(r, 2).value
        if current and raw_mb and str(raw_mb).strip():
            member = str(raw_mb).strip()
            clusters.setdefault(current, [])
            if member not in clusters[current]:
                clusters[current].append(member)
    return clusters


def _write_clusters_sheet(ws, clusters: OrderedDict) -> None:
    """
    Write clusters in columnar format:
      row 1 = cluster names (one per column)
      rows 2+ = members, padded with None where clusters differ in length
    """
    if not clusters:
        return
    names = list(clusters.keys())
    ws.append(names)
    max_len = max(len(v) for v in clusters.values())
    for i in range(max_len):
        ws.append([clusters[n][i] if i < len(clusters[n]) else None
                   for n in names])


# ─────────────────────────────────────────────────────────────────────────────
# Synthetic unit-process clusters for uncovered GTAP regions
# ─────────────────────────────────────────────────────────────────────────────

def _build_synthetic_up_clusters(
        inv_by_region:    OrderedDict,
        n2c:              dict,
        all_gtap_regions: list,
        up_data_keys:     set,
        tpl_cluster_mbrs: dict,
        rc_path:          Path,
        gtap_x:           dict,
        parent_code:      str,
) -> tuple[OrderedDict, OrderedDict]:
    """
    Build synthetic 'X_*' unit-process clusters for GTAP regions not covered
    by any explicit unit process in the template.

    Coverage categories
    -------------------
    1 – Direct : region full name is a key of up_data_keys that matches a GTAP
                 full name → uses its own inv_by_region entry.
    2 – Cluster: region is a member of a cluster that is a key of up_data_keys
                 → uses that cluster's inv_by_region entry as effective UP.
    3 – Missing: all others; this function assigns them to a synthetic cluster.

    Algorithm (for category 3)
    --------------------------
    For each uncovered region, Regions_clusters.xlsx provides a hierarchy of
    clusters ordered from finest (col C) to broadest (rightmost column).
    The function walks that hierarchy until it finds a cluster containing at
    least one covered region, then groups all category-3 regions that share
    the same finest available cluster into one 'X_{cluster}' entry.
    GLOBAL is the ultimate fallback when no cluster level works.

    The unit process for each X cluster is the GTAP-output-weighted average of
    the covered contributors' unit processes.  If a contributor has zero GTAP
    output, equal weights are used instead.

    Parameters
    ----------
    inv_by_region    : {region_code → agg_dict}  (already built)
    n2c              : {full_name → 3-letter code}
    all_gtap_regions : [(full_name, code)]  163 entries
    up_data_keys     : set of region/cluster names present in the template UP
    tpl_cluster_mbrs : {cluster_name → [member_full_names]}  (from template)
    rc_path          : path to Regions_clusters.xlsx
    gtap_x           : {(full_name, parent_code) → output_MUSD}
    parent_code      : parent GTAP sector code (e.g. 'GRO')

    Returns
    -------
    new_up   : OrderedDict {x_cluster_name → agg_dict}
    x_members: OrderedDict {x_cluster_name → [member_codes]}
    """
    def _rc(name: str) -> str:
        return n2c.get(name, name)

    def _members_from_identifier(region_id: str,
                                 cluster_members: dict[str, set]) -> set[str]:
        if region_id in gtap_fn_set:
            return {region_id}
        return set(cluster_members.get(region_id, set()))

    # GLOBAL is a special identifier, not a real GTAP region.
    gtap_fn_set = {fn for fn, _ in all_gtap_regions
                   if str(fn).upper() != 'GLOBAL'}

    # ── Step 1: load hierarchy from Regions_clusters.xlsx ─────────────────────
    # Col A = GTAP full name, col C = level-1 cluster, col D = level-2, …
    hierarchy: dict[str, list]          = {}   # full_name → [lvl1, lvl2, …]
    rc_cl_to_members: dict[str, set]    = {}   # cluster   → set of full_names

    try:
        wb_rc = openpyxl.load_workbook(str(rc_path), data_only=True)
        ws_rc = wb_rc.active
        for _ri in range(2, ws_rc.max_row + 1):   # row 1 = header
            fn_val = ws_rc.cell(_ri, 1).value      # col A = full name
            if not fn_val:
                continue
            fn = str(fn_val).strip()
            levels: list[str] = []
            _ci = 3                                # col C = level 1
            while _ci <= ws_rc.max_column:
                cv = ws_rc.cell(_ri, _ci).value
                if cv and str(cv).strip():
                    lvl = str(cv).strip()
                    levels.append(lvl)
                    rc_cl_to_members.setdefault(lvl, set()).add(fn)
                _ci += 1
            hierarchy[fn] = levels
    except Exception as e:
        print(f"  Warning: could not read {rc_path.name}: {e}")
        return OrderedDict(), OrderedDict()

    for cluster_name, members in tpl_cluster_mbrs.items():
        rc_cl_to_members.setdefault(cluster_name, set()).update(
            member for member in members if member in gtap_fn_set
        )

    # GLOBAL as a valid identifier and ultimate fallback
    rc_cl_to_members['GLOBAL'] = gtap_fn_set.copy()

    # ── Step 2: determine covered regions and their effective agg dict ─────────
    covered: dict[str, dict] = {}   # full_name → agg_dict
    for region_id in up_data_keys:
        region_code = _rc(region_id)
        if region_code not in inv_by_region:
            continue

        agg_dict = inv_by_region[region_code]
        for member_fn in _members_from_identifier(region_id, rc_cl_to_members):
            if member_fn in gtap_fn_set and member_fn not in covered:
                covered[member_fn] = agg_dict

    uncovered_fn = gtap_fn_set - set(covered.keys())
    if not uncovered_fn:
        print("  All 163 GTAP regions are covered; no synthetic clusters needed.")
        return OrderedDict(), OrderedDict()

    # ── Step 3: assign each uncovered region to the finest viable X cluster ────
    region_to_x_cluster: dict[str, str] = {}
    for fn in uncovered_fn:
        assigned = False
        for cluster_name in hierarchy.get(fn, []):
            if rc_cl_to_members.get(cluster_name, set()) & set(covered.keys()):
                region_to_x_cluster[fn] = f'X_{cluster_name}'
                assigned = True
                break
        if not assigned:
            region_to_x_cluster[fn] = 'X_GLOBAL'

    # ── Step 4: group by X cluster and compute weighted average unit process ────
    x_uncov_mbrs: dict[str, list] = {}
    for fn, x_cl in region_to_x_cluster.items():
        x_uncov_mbrs.setdefault(x_cl, []).append(fn)

    new_up:    OrderedDict = OrderedDict()
    x_members: OrderedDict = OrderedDict()

    for x_cl in sorted(x_uncov_mbrs.keys()):
        base_cl    = x_cl[2:]   # strip 'X_' prefix
        contributors = (set(covered.keys()) if base_cl == 'GLOBAL'
                        else rc_cl_to_members.get(base_cl, set()) & set(covered.keys()))

        if not contributors:
            print(f"  Warning: '{x_cl}' has no covered contributors; skipped.")
            continue

        # Weights = GTAP parent-sector output; fall back to equal weights if all 0
        weights: dict[str, float] = {
            fn: gtap_x.get((fn, parent_code), 0.0) for fn in contributors
        }
        if not any(w > 0 for w in weights.values()):
            weights = {fn: 1.0 for fn in contributors}
        total_w = sum(w for w in weights.values() if w > 0)
        weights  = {fn: w for fn, w in weights.items() if w > 0}

        agg: dict = {}
        for fn, w in weights.items():
            share = w / total_w
            for key, val in covered[fn].items():
                agg[key] = agg.get(key, 0.0) + share * val

        uncov_list = x_uncov_mbrs[x_cl]

        if len(uncov_list) == 1:
            # Single uncovered region: register directly with its 3-letter code
            # instead of wrapping it in a cluster.
            solo_code = _rc(uncov_list[0])
            new_up[solo_code] = agg
            print(f"  {solo_code:<10}  (direct synthetic via {x_cl}, "
                  f"← {len(contributors)} contributors)")
        else:
            # Multiple uncovered regions: create named X cluster.
            # GLOBAL must not appear as a cluster member (it is only valid as
            # a cluster name / catch-all identifier).
            member_codes = [_rc(fn) for fn in sorted(uncov_list)
                            if str(_rc(fn)).upper() != 'GLOBAL']
            new_up[x_cl]    = agg
            x_members[x_cl] = member_codes
            print(f"  {x_cl:<40}  "
                  f"{len(uncov_list)} uncovered  "
                  f"← {len(contributors)} contributors")

    n_direct  = sum(1 for k in new_up if not k.startswith('X_'))
    n_cluster = sum(1 for k in new_up if k.startswith('X_'))
    print(f"  Synthetic entries : {n_direct} direct + {n_cluster} X-clusters")
    return new_up, x_members


# ─────────────────────────────────────────────────────────────────────────────
# Public API
# ─────────────────────────────────────────────────────────────────────────────

def make_inventory(inventory_path: str | Path,
                   sector_code: str,
                   repo_path: str,
                   output_path: str | Path = None,
                   version: str = '',
                   ):
    """
    Parse *template_path* with template_checker and write a MARIO
    add_sector inventory to *output_path*.

    Parameters
    ----------
    inventory_path : path to the ENTICE data-collection Excel inventory
    sector_code   : code for the new sector (e.g. 'BAR', 'CEM')
    output_path   : output file; defaults to inventory_{sector_code}.xlsx
                    placed next to the template
    version       : inventory version string (e.g. 'Y26M05'); shown in Summary
    repo_path     : path to repository where auxiliary excel files are stored (GTAP12_X.xlsx, GTAP12_matching.xlsx, Regions_clusters.xlsx)

    Returns
    -------
    Path of the written inventory file.
    """
    inventory_path = Path(inventory_path)
    repo_path = Path(repo_path)
    output_path = (Path(output_path) if output_path
                   else inventory_path.parent / f"inventory_{sector_code}.xlsx")

    # ── Parse template ────────────────────────────────────────────────────────
    importlib.reload(template_checker)
    result = template_checker.parse_template(inventory_path)

    gi          = result['general_info']
    parent_code = gi['parent_code']   # e.g. 'GRO'
    print(f"  New sector : {gi['new_sector']}")
    print(f"  Parent     : {gi['parent_name']}  [{parent_code}]")

    # ── Extra reads from workbook (clusters, name→code lookup) ────────────────
    wb   = openpyxl.load_workbook(str(inventory_path), data_only=True)
    n2c  = _name_to_code_map(wb)   # region full name → code

    # ── GTAP12 sector / factor-of-production code mapping ────────────────────
    # GTAP12_matching.xlsx:
    #   "Sector" sheet          col A = inventory name, col B = GTAP12 code
    #   "Factor of production"  col A = inventory name, col B = GTAP12 code
    _matching_path = repo_path / 'GTAP12_matching.xlsx'
    _sector_to_gtap: dict[str, str] = {}
    _factprod_to_gtap: dict[str, str] = {}
    if _matching_path.exists():
        _wb_m = openpyxl.load_workbook(str(_matching_path), data_only=True)
        if 'Sector' in _wb_m.sheetnames:
            _ws_s = _wb_m['Sector']
            for _r in range(2, _ws_s.max_row + 1):
                _k = _ws_s.cell(_r, 1).value
                _v = _ws_s.cell(_r, 2).value
                if _k and _v:
                    _sector_to_gtap[str(_k).strip()] = str(_v).strip()
        if 'Factor of production' in _wb_m.sheetnames:
            _ws_f = _wb_m['Factor of production']
            for _r in range(2, _ws_f.max_row + 1):
                _k = _ws_f.cell(_r, 1).value
                _v = _ws_f.cell(_r, 2).value
                if _k and _v:
                    _factprod_to_gtap[str(_k).strip()] = str(_v).strip()
        print(f"  GTAP matching: {len(_sector_to_gtap)} sectors, "
              f"{len(_factprod_to_gtap)} factors loaded")
    else:
        print(f"  Warning: {_matching_path.name} not found – "
              "sector/factor codes will be used as-is.")

    # Lead from General info D4
    _lead = wb['General info'].cell(4, 4).value
    data_lead = str(_lead).strip() if _lead else ''

    # Sector-specific sources from Source sheet rows 20+ (col A = name, col B = URL)
    sources = []
    if 'Source' in wb.sheetnames:
        ws_src = wb['Source']
        for r in range(20, ws_src.max_row + 1):
            src_name = ws_src.cell(r, 1).value
            src_url  = ws_src.cell(r, 2).value
            if src_name and str(src_name).strip():
                sources.append((str(src_name).strip(),
                                str(src_url).strip() if src_url else ''))

    def rc(name: str) -> str:
        """Resolve a region full-name to its code; unknowns / clusters as-is."""
        return n2c.get(name, name)

    def _pick_residual_up_name(existing_names: set[str]) -> str:
        """Choose a readable cluster name for the residual GLOBAL inventory."""
        if 'GLOBAL_ROW' not in existing_names:
            return 'GLOBAL_ROW'
        return 'GLOBAL_ROW2'

    regions_clusters_raw = _read_clusters(wb['Region'],        start_row=328)
    sectors_clusters_raw = _read_clusters(wb['Sector'],        start_row=369)
    factprod_clusters_raw = _read_clusters(wb['Primary input'], start_row=17)

    # Convert Regions Clusters members from full names to 3-letter codes.
    # Members that are not found in n2c (e.g. already a code) are kept as-is.
    # GLOBAL is a special catch-all identifier, not a real GTAP region code:
    # it must never appear as the content of a cluster, only as a cluster name.
    regions_clusters = OrderedDict(
        (cluster, [rc(m) for m in members
                   if str(m).upper() != 'GLOBAL'])
        for cluster, members in regions_clusters_raw.items()
    )

    # Convert Sectors Clusters and Factors Clusters members to GTAP12 codes,
    # deduplicate (multiple inventory names may share the same GTAP code), and
    # handle single-item results: a cluster with exactly one unique GTAP code is
    # not a real cluster — any reference to it as a db_item in the inventory
    # should be replaced directly with that code.
    def _resolve_clusters(raw: dict, mapping: dict) -> tuple[OrderedDict, dict]:
        """
        Returns (clusters, direct_map) where:
          clusters   – multi-item clusters only (OrderedDict cluster→[codes])
          direct_map – single-item clusters: {cluster_name: single_gtap_code}
        """
        clusters: OrderedDict = OrderedDict()
        direct_map: dict = {}
        for cluster, members in raw.items():
            codes = list(dict.fromkeys(mapping.get(m, m) for m in members))
            if len(codes) == 1:
                direct_map[cluster] = codes[0]
                print(f"  Single-item cluster '{cluster}' → collapsed to '{codes[0]}'")
            else:
                clusters[cluster] = codes
        return clusters, direct_map

    sectors_clusters,  _sector_single_map  = _resolve_clusters(
        sectors_clusters_raw,  _sector_to_gtap)
    factprod_clusters, _factprod_single_map = _resolve_clusters(
        factprod_clusters_raw, _factprod_to_gtap)

    # ── UN regional aggregate clusters (Region sheet, col A rows 3-165) ──
    # If a unit-process column or a total-production row is defined for one
    # of the built-in UN sub-regional aggregates (e.g. CARIBBEAN, WESTERN
    # ASIA), include that cluster in the Regions Clusters sheet so MARIO can
    # expand it to the individual country codes.
    # Only clusters that are actually used are added; clusters already defined
    # as custom entries (rows 328+) are left untouched.
    _ws_r = wb['Region']
    _gtap_region_names: set[str] = set()
    _un_clusters_raw: OrderedDict = OrderedDict()
    for _r in range(3, 166):
        _cn = _ws_r.cell(_r, 1).value
        _mb = _ws_r.cell(_r, 2).value
        if _mb and str(_mb).strip():
            _gtap_region_names.add(str(_mb).strip())
        if _cn and str(_cn).strip() and _mb and str(_mb).strip():
            _cn = str(_cn).strip()
            _mb = str(_mb).strip()
            _un_clusters_raw.setdefault(_cn, [])
            if _mb not in _un_clusters_raw[_cn]:
                _un_clusters_raw[_cn].append(_mb)

    # Identifiers actually used in Unit process (row-1 headers) and Total production (col A)
    _raw_up_ids: set = set(result['unit_process'].keys())
    _has_global_up = 'GLOBAL' in _raw_up_ids

    _used_ids: set = set(_raw_up_ids)
    _used_ids |= {tp['region'] for tp in result['total_production'] if tp['region']}

    for _cn, _members in _un_clusters_raw.items():
        if _cn in _used_ids and _cn not in regions_clusters:
            regions_clusters[_cn] = [rc(_m) for _m in _members
                                     if str(_m).upper() != 'GLOBAL']

    print(f"  Region clusters    : {list(regions_clusters.keys())}")
    print(f"  Sector clusters    : {list(sectors_clusters.keys())}")
    print(f"  Fact-prod clusters : {list(factprod_clusters.keys())}")

    # ── Shared reference data (reused by synthetic UP and Total outputs) ───────
    # Load once here so we don't read the same files twice later.
    db_X_path = (repo_path / 'GTAP12_X.xlsx')
    if db_X_path.exists():
        _shared_gtap_x = template_checker._load_gtap_x(db_X_path)
    else:
        _shared_gtap_x = {}
        print(f"  Warning: {db_X_path.name} not found – "
              "GTAP weights unavailable.")

    _shared_gtap_regions: list[tuple[str, str]] = []
    _ws_reg = wb['Region']
    for _r in range(3, 166):
        _fn_s   = _ws_reg.cell(_r, 2).value
        _code_s = _ws_reg.cell(_r, 3).value
        if _fn_s and str(_fn_s).strip().upper() != 'GLOBAL':
            _shared_gtap_regions.append((str(_fn_s).strip(),
                                         str(_code_s).strip() if _code_s else ''))

    _shared_cluster_mbrs = template_checker._build_cluster_to_members(wb)

    # ── Aggregate inventory rows per producing region ─────────────────────────
    #
    # Source: result['unit_process']  →  OrderedDict {
    #     region_name: [ {row, type, mapped_input, origin, value} ]
    # }
    #
    # Rules:
    #   Item type  = 'Factor of production'  if type == 'Primary_input'
    #              = 'Sector'                 otherwise
    #   DB Item    = mapped_input  (col C of Unit process) — used as-is
    #   DB Region  = 'GLOBAL'      if origin is blank / 'GLOBAL'
    #              = region code   otherwise (map full name → code)
    #   Change type= 'Update'      always
    #   Rows with the same (Item type, DB Item, DB Region) are summed.

    up_data = result['unit_process']

    inv_by_region: OrderedDict = OrderedDict()

    for region_name, rows in up_data.items():
        r_code = rc(region_name)
        agg: dict = {}
        for row in rows:
            mario_type = ('Factor of production'
                          if row['type'] == 'Primary_input' else 'Sector')
            db_region = 'GLOBAL'
            raw_item = row['mapped_input']
            if mario_type == 'Factor of production':
                db_item = _factprod_to_gtap.get(raw_item, raw_item)
                # If the name resolved to a cluster that collapsed to 1 item,
                # use the direct GTAP code instead of the cluster name.
                db_item = _factprod_single_map.get(db_item, db_item)
            else:
                db_item = _sector_to_gtap.get(raw_item, raw_item)
                db_item = _sector_single_map.get(db_item, db_item)
            key = (mario_type, db_item, db_region)
            agg[key] = agg.get(key, 0.0) + row['value']
        inv_by_region[r_code] = agg
        print(f"  {r_code:<10}  {len(agg)} inventory rows")

    if _has_global_up and 'GLOBAL' in inv_by_region:
        _explicit_up_ids = set(up_data.keys()) - {'GLOBAL'}
        if _explicit_up_ids:
            _explicit_gtap_names = template_checker._expand_region_ids_to_gtap_names(
                _explicit_up_ids, _shared_cluster_mbrs, _gtap_region_names
            )
            _residual_gtap_names = [
                _fn for _fn, _ in _shared_gtap_regions
                if _fn not in _explicit_gtap_names
            ]
            _global_agg = inv_by_region.pop('GLOBAL')

            if _residual_gtap_names:
                _residual_name = _pick_residual_up_name(
                    set(inv_by_region.keys()) | set(regions_clusters.keys()) | _raw_up_ids
                )
                inv_by_region[_residual_name] = _global_agg
                regions_clusters[_residual_name] = [rc(_name) for _name in _residual_gtap_names]
                print(
                    "  GLOBAL inventory converted to residual cluster "
                    f"'{_residual_name}' excluding explicit inventories."
                )
            else:
                print("  GLOBAL inventory fully overlapped by explicit inventories; removed.")

    # ── Synthetic unit processes for GTAP regions with no explicit UP ─────────
    # Regions_clusters.xlsx provides a cluster hierarchy (col C = finest level,
    # rightward = broader) used to locate the finest cluster that contains at
    # least one covered region.  A weighted-average unit process is computed
    # from those covered contributors and registered as 'X_{cluster}'.
    _rc_xlsx = repo_path / 'Regions_clusters.xlsx'
    if _has_global_up:
        print("  GLOBAL inventory present: synthetic unit processes skipped.")
    elif _rc_xlsx.exists():
        _syn_up, _syn_mbrs = _build_synthetic_up_clusters(
            inv_by_region,
            n2c,
            _shared_gtap_regions,
            set(up_data.keys()),
            _shared_cluster_mbrs,
            _rc_xlsx,
            _shared_gtap_x,
            parent_code,
        )
        for _x_cl, _x_agg in _syn_up.items():
            inv_by_region[_x_cl] = _x_agg        # new UP entry
        for _x_cl, _x_codes in _syn_mbrs.items():
            regions_clusters[_x_cl] = _x_codes   # new cluster definition
    else:
        print(f"  Warning: Regions_clusters.xlsx not found next to template – "
              "synthetic unit processes not computed.")

    # ── Write output workbook ─────────────────────────────────────────────────
    wb_out = openpyxl.Workbook()
    wb_out.remove(wb_out.active)

    # Build the ordered list of all sheet names for the Summary links column.
    # Populated as sheets are created below.
    all_sheet_names: list = []

    # 0. Summary ───────────────────────────────────────────────────────────────
    ws_sum = wb_out.create_sheet('Summary')

    # -- Info block (A1:C5) ---------------------------------------------------
    ws_sum.cell(1, 1).value = 'INFO'

    ws_sum.cell(2, 1).value = 'Sector'
    ws_sum.cell(2, 2).value = sector_code
    ws_sum.cell(2, 3).value = gi['new_sector']

    ws_sum.cell(3, 1).value = 'Parent sector'
    ws_sum.cell(3, 2).value = parent_code
    ws_sum.cell(3, 3).value = gi['parent_name']

    ws_sum.cell(4, 1).value = 'Inventory version'
    ws_sum.cell(4, 2).value = version

    ws_sum.cell(5, 1).value = 'Data collection lead'
    ws_sum.cell(5, 2).value = data_lead

    # -- Sources block (A7+) --------------------------------------------------
    ws_sum.cell(7, 1).value = 'Sources'
    for i, (src_name, src_url) in enumerate(sources):
        row = 8 + i
        ws_sum.cell(row, 1).value = src_name
        ws_sum.cell(row, 2).value = src_url

    # Links column (I) will be filled after all sheets are known (see below).

    # 1. Master ────────────────────────────────────────────────────────────────
    ws_m = wb_out.create_sheet('Master')
    all_sheet_names.append('Master')
    ws_m.append(['Region', 'Sector', 'Inventory sheet', 'Quantity', 'Unit',
                  'Final consumption', 'Consumption category',
                  'Parent Sector', 'Leave empty', 'Source', 'Notes', 'Add or Split'])
    for r_code in inv_by_region:
        sn = _safe_sheet_name(f"{sector_code}_{r_code}")
        ws_m.append([r_code, sector_code, sn, None, 'M USD',
                     None, None, parent_code, None, None, None, 'Split'])

    # 2. Regions Clusters ──────────────────────────────────────────────────────
    regions_clusters = OrderedDict(
        (k, v) for k, v in regions_clusters.items() if v
    )
    _write_clusters_sheet(wb_out.create_sheet('Regions Clusters'),
                          regions_clusters)
    all_sheet_names.append('Regions Clusters')

    # 3. Sectors Clusters ──────────────────────────────────────────────────────
    _write_clusters_sheet(wb_out.create_sheet('Sectors Clusters'),
                          sectors_clusters)
    all_sheet_names.append('Sectors Clusters')

    # 4. Factors Clusters ──────────────────────────────────────────────────────
    _write_clusters_sheet(wb_out.create_sheet('Factors Clusters'),
                          factprod_clusters)
    all_sheet_names.append('Factors Clusters')

    # 5. Inventory sheets  (one per producing region) ─────────────────────────
    #    Column layout:
    #      A  Quantity   B  Unit   C  Input (empty)   D  Item type
    #      E  DB Item    F  DB Region                 G  Change type
    for r_code, agg in inv_by_region.items():
        sn = _safe_sheet_name(f"{sector_code}_{r_code}")
        ws_inv = wb_out.create_sheet(sn)
        ws_inv.append(['Quantity', 'Unit', 'Input', 'Item type',
                        'DB Item', 'DB Region', 'Change type'])
        for (mario_type, db_item, db_region), value in agg.items():
            ws_inv.append([value, 'M USD', None, mario_type,
                            db_item, db_region, 'Update'])
        all_sheet_names.append(sn)

    # 6. DB units  (unchanged) ─────────────────────────────────────────────────
    ws_db = wb_out.create_sheet('DB units')
    for row in DB_UNITS_DATA:
        ws_db.append(list(row))
    all_sheet_names.append('DB units')

    # 7. Total outputs ─────────────────────────────────────────────────────────
    # One row for every GTAP region (163 total).  Values are resolved in order:
    #
    #   Case 1 – Direct  : region full name appears in template Total production.
    #   Case 2 – Cluster : region belongs to a cluster present in Total
    #                      production; share is weighted by GTAP parent-sector
    #                      output (from GTAP12_X.xlsx).
    #                      Example: cluster REST OF EUROPE has TP value 100;
    #                      Albania accounts for 50 of the cluster's 200 GTAP
    #                      parent output → Albania gets 100 × 50/200 = 25.
    #   Case 2b – GLOBAL : treated as a cluster covering all 163 GTAP regions,
    #                      distributed with the same weighting logic.
    #   Case 3 – Median  : region absent from direct and cluster coverage;
    #                      value = region_gtap_x × median(new/parent ratio),
    #                      where the median is over all Case-1/2 regions with
    #                      new ≤ parent (values where new > parent are data
    #                      errors and are excluded from the ratio calculation).

    # ── GTAP reference data (reuse shared variables computed earlier) ────────
    _gtap_x_to       = _shared_gtap_x
    _gtap_regions_to = _shared_gtap_regions
    _cluster_mbrs_to = _shared_cluster_mbrs

    def _gx(full_name: str) -> float:
        """GTAP parent-sector output for a region (0.0 if unknown)."""
        return _gtap_x_to.get((full_name, parent_code), 0.0)

    # ── Template Total production → {identifier: float value} ────────────────
    _tp_vals_to: dict[str, float] = {}
    _has_global_to  = False
    _global_val_to: float | None = None

    for _tp in result['total_production']:
        if _tp['value'] is None:
            continue
        try:
            _v = float(_tp['value'])
        except (ValueError, TypeError):
            continue
        _rid = _tp['region']
        if not _rid:
            continue
        if str(_rid).strip().upper() == 'GLOBAL':
            _has_global_to = True
            _global_val_to = _v
        else:
            _tp_vals_to[str(_rid).strip()] = _v

    # ── Resolve each GTAP region ──────────────────────────────────────────────
    _resolved_to: dict[str, tuple] = {}   # full_name → (value, method)

    for _fn, _rcode in _gtap_regions_to:

        # Case 1: direct match
        if _fn in _tp_vals_to:
            _resolved_to[_fn] = (_tp_vals_to[_fn], 'direct')
            continue

        # Case 2: region belongs to a specific cluster in Total production
        _matched_cl: str | None = None
        for _cname, _cmbrs in _cluster_mbrs_to.items():
            if _cname in _tp_vals_to and _fn in _cmbrs:
                _matched_cl = _cname
                break


        # Case 3: unknown → placeholder until median is computed
        _resolved_to[_fn] = (None, 'pending')

    # ── Median new/parent ratio (Cases 1+2, excluding new > parent) ──────────
    _ratios_to: list[float] = []
    for _fn, _rcode in _gtap_regions_to:
        _entry = _resolved_to.get(_fn)
        if _entry is None or _entry[1] not in ('direct', 'cluster'):
            continue
        _v_e, _gv_e = _entry[0], _gx(_fn)
        if _gv_e <= 0 or _v_e > _gv_e:   # zero GTAP data or data error
            continue
        _ratios_to.append(_v_e / _gv_e)

    _median_ratio_to = statistics.median(_ratios_to) if _ratios_to else 0.0
    print(f"  Median new/parent ratio : {_median_ratio_to:.4%}  "
          f"(n = {len(_ratios_to)})")

    # ── Fill Case-3 regions with median estimate ──────────────────────────────
    for _fn, _rcode in _gtap_regions_to:
        if _resolved_to[_fn][1] == 'pending':
            _resolved_to[_fn] = (_median_ratio_to * _gx(_fn), 'median')

    # ── Cap values exceeding the GTAP parent-sector output ───────────────────
    # Direct and cluster values that exceed the parent's GTAP output are
    # capped at that output.  These regions were already excluded from the
    # median calculation above (via the _v_e > _gv_e guard), so the cap does
    # not affect the median.
    for _fn, _rcode in _gtap_regions_to:
        _v_cap, _meth_cap = _resolved_to[_fn]
        if _meth_cap in ('direct', 'cluster') and _v_cap is not None:
            _gv_cap = _gx(_fn)
            if _gv_cap > 0 and _v_cap > _gv_cap:
                _resolved_to[_fn] = (_gv_cap, 'capped')

    # ── Write sheet ───────────────────────────────────────────────────────────
    ws_to = wb_out.create_sheet('Total outputs')
    ws_to.append(['Sector', 'Region', 'Quantity', 'Unit', 'Source','Notes'])
    _to_counts: dict[str, int] = {}
    for _fn, _rcode in _gtap_regions_to:
        _v_out, _meth = _resolved_to[_fn]
        ws_to.append([sector_code, _rcode or rc(_fn), _v_out, 'M USD'])
        _to_counts[_meth] = _to_counts.get(_meth, 0) + 1

    n_tp = len(_gtap_regions_to)
    print(f"  Total outputs rows : {n_tp}  "
          f"(direct={_to_counts.get('direct', 0)}, "
          f"cluster={_to_counts.get('cluster', 0)}, "
          f"cluster-zero={_to_counts.get('cluster-zero', 0)}, "
          f"capped={_to_counts.get('capped', 0)}, "
          f"median={_to_counts.get('median', 0)})")
    all_sheet_names.append('Total outputs')

    # ── Fill Summary column I: hyperlinks to every other sheet ────────────────
    ws_sum.cell(1, 9).value = 'LINKS TO PAGES'
    for i, sn in enumerate(all_sheet_names):
        cell = ws_sum.cell(2 + i, 9)
        cell.value = sn
        # Internal Excel hyperlink: navigate to cell A1 of the target sheet.
        # Sheet names with spaces or special chars must be wrapped in single quotes.
        safe = sn.replace("'", "''")          # escape any literal apostrophes
        cell.hyperlink = f"#'{safe}'!A1"

    wb_out.save(str(output_path))
    print(f"\nInventory saved: {output_path.name}")
    return output_path


# ─────────────────────────────────────────────────────────────────────────────
# CLI entry point
# ─────────────────────────────────────────────────────────────────────────────

if __name__ == '__main__':
    ap = argparse.ArgumentParser(
        description='Convert ENTICE template to MARIO add_sector inventory'
    )
    ap.add_argument('template',
                    help='Path to the ENTICE Excel data-collection template')
    ap.add_argument('--sector-code', required=True,
                    help='Code for the new sector (e.g. BAR, CEM, STE)')
    ap.add_argument('--output', default=None,
                    help='Output file path (default: inventory_{code}.xlsx)')
    ap.add_argument('--version', default='',
                    help='Inventory version string shown in Summary (e.g. Y26M05)')
    ap.add_argument('--gtap', default=None,
                    help='Path to GTAP12_X.xlsx (default: auto-detected next to template)')
    args = ap.parse_args()
    make_inventory(args.template, args.sector_code.upper(),
                   args.output, args.version, args.gtap)
