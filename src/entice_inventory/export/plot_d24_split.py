"""
Visualise the GTAP12 ENTICE disaggregation of the original GTAP sectors for D2.4.

One subplot per "parent" GTAP sector. Each subplot has two stacked columns -- the
total output X and the total value added VA of that sector group -- each broken
down into the carved-out ENTICE sub-sectors plus a residual block. The group
total (in T USD) is annotated on top of each column, so X and VA are shown
together in a single figure.

Everything is aggregated over regions (or restricted to one via `region=`).

Data sources
------------
X  (total output):
    * group total : GTAP12_X.xlsx, "GTAP totals" sheet, column X.
    * sub-sector  : "Total outputs" sheet of each D2.4 inventory.

VA  (total value added):
    * group total : VA.xlsx ("VA" sheet); VA = Tax + TTM + Capital + Labor.
    * sub-sector  : the D2.4 inventories are stored as coefficients, so the
                    sub-sector VA is rebuilt as (VA coefficient) x (sub-sector X),
                    region by region. The VA coefficient is the column sum of the
                    value-added matrix v.parquet of the D2.4 coefficients database;
                    the per-region X comes from the "Total outputs" sheets.

The residual = group total minus the carved-out sub-sectors (e.g. for P_C the
residual is labelled PEP, see DEFAULT_RESIDUAL_LABELS).

Typical use
-----------
    from plot_d24_split import plot_sector_split

    plot_sector_split("all")                         # every split sector
    plot_sector_split(["EEQ", "MVH"])                # selected sectors
    plot_sector_split("all", region="ITA")           # one region
    plot_sector_split("all", normalize=True)         # shares (%) per column
"""

from __future__ import annotations

import glob
import math
import os
import warnings

import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
from openpyxl import load_workbook

from entice_inventory.core.matching_utils import load_parent_map
from entice_inventory.core.paths import DATA_DIR, PROJECT_ROOT, find_data_file, data_path

# --------------------------------------------------------------------------- #
# Default locations
# --------------------------------------------------------------------------- #
DEFAULT_GTAP_X_PATH = str(find_data_file('GTAP12_X.xlsx'))

DEFAULT_D24_DIR = str(data_path('reference_inventories'))
DEFAULT_MARIO_DIR = str(data_path('mario_inventories'))
DEFAULT_VA_PATH = str(data_path('value_added'))
DEFAULT_V_PARQUET_PATH = str(data_path('reference_coefficients') / 'v.parquet')

# authoritative sub-sector -> parent (GTAP12) mapping
DEFAULT_MATCHING_PATH = str(DATA_DIR / "GTAP12_matching.xlsx")

# small on-disk cache so we don't re-open ~60 OneDrive files on every call
DEFAULT_CACHE_PATH = str(PROJECT_ROOT / ".d24_split_cache.pkl")
DEFAULT_MIXED_CACHE_PATH = str(PROJECT_ROOT / ".d24_split_mixed_cache.pkl")

# value-added components in VA.xlsx that sum to total VA
_VA_COMPONENTS = ["Tax", "TTM", "Capital", "Labor"]

_INVENTORY_META_SHEETS = {
    "Summary",
    "Master",
    "Regions Clusters",
    "Sectors Clusters",
    "Factors Clusters",
    "Total outputs",
    "Trades",
    "DB units",
}

# GTAP12_matching.xlsx is the single source of truth for sub-sector codes, so we
# do NOT alias mismatches (e.g. inventory NMC vs matching NCM). Such codes are
# surfaced by check_parent_consistency as "not found" so they get reconciled at
# the source rather than hidden. Kept as an empty hook in case it is ever needed.
_CODE_ALIASES: dict = {}

# Some parents leave a meaningful sector as their residual rather than a generic
# "Rest of <parent>". E.g. once coke (PCP) is carved out of P_C the residual is
# petroleum products (PEP); SOP out of ELE leaves XEL. (RBR is now a real
# sub-sector sized as the RPP residual, so it is no longer a residual label.)
# Override / extend via the `residual_labels` argument.
DEFAULT_RESIDUAL_LABELS = {"P_C": "PEP", "ELE": "XEL"}


# --------------------------------------------------------------------------- #
# Original (pre-split) quantities
# --------------------------------------------------------------------------- #
def _read_original_X(gtap_x_path: str, region: str | None = None) -> pd.Series:
    """Original total output per GTAP sector code.

    Summed over all regions, or restricted to `region` if given.
    """
    df = pd.read_excel(gtap_x_path, sheet_name="GTAP totals")
    # columns: Region | Full name regions | Sector to | Sector to full name | X
    df.columns = ["Region", "RegName", "Sector", "SectorName", "X"][: df.shape[1]]
    df["X"] = pd.to_numeric(df["X"], errors="coerce")
    df["Region"] = df["Region"].astype(str).str.strip()
    if region is not None:
        df = df[df["Region"] == region]
    return df.groupby("Sector")["X"].sum()


def _read_original_VA(va_path: str, region: str | None = None) -> pd.Series:
    """Original total value added per GTAP sector code.

    Summed over all regions (or restricted to `region`).
    VA.xlsx columns: Region | Level | Item | Tax | TTM | Capital | Labor
    Total VA = Tax + TTM + Capital + Labor.
    """
    df = pd.read_excel(va_path, sheet_name="VA")
    comps = [c for c in _VA_COMPONENTS if c in df.columns]
    if not comps:
        raise ValueError(
            f"None of the VA components {_VA_COMPONENTS} found in {va_path} "
            f"(columns: {list(df.columns)})"
        )
    df["_VA"] = df[comps].apply(pd.to_numeric, errors="coerce").sum(axis=1)
    df["Region"] = df["Region"].astype(str).str.strip()
    if region is not None:
        df = df[df["Region"] == region]
    return df.groupby("Item")["_VA"].sum()


# --------------------------------------------------------------------------- #
# Sub-sector data collection (region level, so VA can be weighted by X)
# --------------------------------------------------------------------------- #
def _summary_lookup(summary: pd.DataFrame, label: str):
    """Return (code, full_name) for a row whose first column equals `label`."""
    col0 = summary.iloc[:, 0].astype(str).str.strip()
    hit = summary[col0 == label]
    if hit.empty:
        return None, None
    row = hit.iloc[0]
    code = row.iloc[1] if summary.shape[1] > 1 else None
    name = row.iloc[2] if summary.shape[1] > 2 else None
    code = None if pd.isna(code) else str(code).strip()
    name = None if pd.isna(name) else str(name).strip()
    return code, name


def _summary_lookup_ws(ws, label: str):
    """Return (code, full_name) for a Summary-sheet row whose first cell matches."""
    for row in ws.iter_rows(values_only=True):
        if not row:
            continue
        head = row[0]
        if head is None or str(head).strip() != label:
            continue
        code = row[1] if len(row) > 1 else None
        name = row[2] if len(row) > 2 else None
        code = None if code is None else str(code).strip()
        name = None if name is None else str(name).strip()
        return code, name
    return None, None


def _sheet_header_map(ws) -> dict[str, int]:
    header = next(ws.iter_rows(min_row=1, max_row=1, values_only=True))
    return {
        str(value).strip(): idx
        for idx, value in enumerate(header)
        if value is not None and str(value).strip()
    }


def _safe_float(value) -> float:
    try:
        return float(value)
    except (TypeError, ValueError):
        return 0.0


def _cluster_map_from_regions_sheet(ws) -> dict[str, list[str]]:
    """Map a Master-sheet region label to the GTAP region codes it covers."""
    rows = list(ws.iter_rows(values_only=True))
    if not rows:
        return {}

    mapping: dict[str, list[str]] = {}
    ncols = max(len(row) for row in rows)
    for col_idx in range(ncols):
        label = rows[0][col_idx] if col_idx < len(rows[0]) else None
        if label is None or not str(label).strip():
            continue
        codes = []
        for row in rows[1:]:
            value = row[col_idx] if col_idx < len(row) else None
            if value is None or not str(value).strip():
                continue
            codes.append(str(value).strip())
        mapping[str(label).strip()] = codes
    return mapping


def _master_region_sheet_map(master_ws, cluster_map: dict[str, list[str]]) -> dict[str, str]:
    """Resolve each GTAP region code to the inventory sheet that should serve it."""
    headers = _sheet_header_map(master_ws)
    region_idx = headers.get("Region")
    sheet_idx = headers.get("Inventory sheet")
    if region_idx is None or sheet_idx is None:
        return {}

    mapping: dict[str, str] = {}
    for row in master_ws.iter_rows(min_row=2, values_only=True):
        if not row:
            continue
        region_label = row[region_idx] if region_idx < len(row) else None
        sheet_name = row[sheet_idx] if sheet_idx < len(row) else None
        if region_label is None or sheet_name is None:
            continue

        sheet_name = str(sheet_name).strip()
        members = cluster_map.get(str(region_label).strip(), [str(region_label).strip()])
        for region in members:
            mapping.setdefault(region, sheet_name)
    return mapping


def _total_outputs_from_sheet(ws) -> pd.DataFrame:
    headers = _sheet_header_map(ws)
    region_idx = headers.get("Region")
    quantity_idx = headers.get("Quantity")
    if region_idx is None or quantity_idx is None:
        return pd.DataFrame(columns=["region", "X"])

    rows = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        if not row:
            continue
        region = row[region_idx] if region_idx < len(row) else None
        if region is None or not str(region).strip():
            continue
        quantity = row[quantity_idx] if quantity_idx < len(row) else None
        rows.append({"region": str(region).strip(), "X": _safe_float(quantity)})
    return pd.DataFrame(rows, columns=["region", "X"])


def _sheet_va_coefficient(ws) -> float:
    headers = _sheet_header_map(ws)
    quantity_idx = headers.get("Quantity")
    db_item_idx = headers.get("DB Item")
    if quantity_idx is None or db_item_idx is None:
        return 0.0

    total = 0.0
    for row in ws.iter_rows(min_row=2, values_only=True):
        if not row:
            continue
        db_item = row[db_item_idx] if db_item_idx < len(row) else None
        if db_item is None or str(db_item).strip() not in _VA_COMPONENTS:
            continue
        quantity = row[quantity_idx] if quantity_idx < len(row) else None
        total += _safe_float(quantity)
    return total


def _read_inventory_workbook(path: str, source: str) -> pd.DataFrame:
    """Read one inventory workbook and return one row per GTAP region."""
    try:
        wb = load_workbook(path, read_only=True, data_only=True)
        summary_ws = wb["Summary"]
        master_ws = wb["Master"]
        regions_ws = wb["Regions Clusters"]
        total_outputs_ws = wb["Total outputs"]
    except Exception as exc:  # noqa: BLE001 - report and keep going
        warnings.warn(f"Could not read {os.path.basename(path)}: {exc}")
        return pd.DataFrame()

    sub, sub_name = _summary_lookup_ws(summary_ws, "Sector")
    parent, _parent_name = _summary_lookup_ws(summary_ws, "Parent sector")
    if not sub or not parent:
        warnings.warn(f"Missing Sector/Parent in {os.path.basename(path)}")
        return pd.DataFrame()

    total_outputs = _total_outputs_from_sheet(total_outputs_ws)
    if total_outputs.empty:
        warnings.warn(f"Missing Total outputs data in {os.path.basename(path)}")
        return pd.DataFrame()

    cluster_map = _cluster_map_from_regions_sheet(regions_ws)
    region_to_sheet = _master_region_sheet_map(master_ws, cluster_map)
    sheet_coef_cache: dict[str, float] = {}
    missing_region_sheets = []

    rows = []
    for region, X in total_outputs[["region", "X"]].itertuples(index=False, name=None):
        sheet_name = region_to_sheet.get(region)
        va_coef = 0.0
        if sheet_name is None:
            if abs(X) > 1e-9:
                missing_region_sheets.append(region)
        else:
            if sheet_name not in sheet_coef_cache:
                if sheet_name not in wb.sheetnames:
                    warnings.warn(
                        f"Inventory sheet {sheet_name!r} not found in "
                        f"{os.path.basename(path)}; VA treated as 0."
                    )
                    sheet_coef_cache[sheet_name] = 0.0
                else:
                    sheet_coef_cache[sheet_name] = _sheet_va_coefficient(wb[sheet_name])
            va_coef = sheet_coef_cache[sheet_name]

        rows.append(
            {
                "region": region,
                "X": X,
                "va_coef": va_coef,
                "subsector": sub,
                "subsector_name": sub_name,
                "parent_inventory": parent,
                "inventory_source": source,
                "inventory_file": os.path.basename(path),
            }
        )

    if missing_region_sheets:
        warnings.warn(
            f"No inventory-sheet mapping found in {os.path.basename(path)} for "
            f"regions {sorted(set(missing_region_sheets))}; VA treated as 0 there."
        )

    return pd.DataFrame(rows)


def _collect_inventory_dir(inventory_dir: str, source: str) -> pd.DataFrame:
    frames = []
    seen_subsectors: dict[str, str] = {}

    files = sorted(glob.glob(os.path.join(inventory_dir, "*.xlsx")))
    for path in files:
        if os.path.basename(path).startswith("~$"):
            continue
        frame = _read_inventory_workbook(path, source=source)
        if frame.empty:
            continue

        subsector = str(frame["subsector"].iat[0]).strip()
        if subsector in seen_subsectors:
            warnings.warn(
                f"Multiple {source} inventories found for {subsector}; keeping "
                f"{seen_subsectors[subsector]} and skipping {os.path.basename(path)}."
            )
            continue
        seen_subsectors[subsector] = os.path.basename(path)
        frames.append(frame)

    if not frames:
        return pd.DataFrame(
            columns=[
                "region",
                "X",
                "va_coef",
                "subsector",
                "subsector_name",
                "parent_inventory",
                "inventory_source",
                "inventory_file",
            ]
        )
    return pd.concat(frames, ignore_index=True)


def _select_inventory_sources(
    primary: pd.DataFrame,
    fallback: pd.DataFrame | None,
    prefer_fallback_subsectors=(),
) -> pd.DataFrame:
    """Choose one source per sub-sector: primary by default, fallback if needed."""
    if fallback is None or fallback.empty:
        return primary.copy()
    if primary.empty:
        return fallback.copy()

    prefer_fallback = {str(code).strip() for code in prefer_fallback_subsectors}
    combined = pd.concat([primary, fallback], ignore_index=True)

    options = (
        combined[["subsector", "inventory_source"]]
        .drop_duplicates()
        .groupby("subsector")["inventory_source"]
        .agg(list)
    )

    chosen = {}
    for subsector, sources in options.items():
        if subsector in prefer_fallback and "mario" in sources:
            chosen[subsector] = "mario"
        elif "d24" in sources:
            chosen[subsector] = "d24"
        else:
            chosen[subsector] = sources[0]

    selected = combined["subsector"].map(chosen)
    return combined.loc[combined["inventory_source"] == selected].copy()


def _parent_map_from_matching(matching_path: str) -> dict:
    """Authoritative sub-sector code -> parent GTAP12 code, from GTAP12_matching.

    Delegates to ``matching_utils.load_parent_map`` so the matching workbook is
    parsed in exactly one place across the whole pipeline.
    """
    return load_parent_map(matching_path)


def collect_split_data(
    d24_dir: str = DEFAULT_D24_DIR,
    mario_dir: str | None = None,
    prefer_mario_subsectors=(),
    cache_path: str | None = None,
    matching_path: str | None = None,
    refresh: bool = False,
) -> pd.DataFrame:
    """
    Scan the selected inventory workbooks and return a tidy long-form DataFrame
    with one row per (sub-sector, region):

        subsector | subsector_name | parent | parent_inventory | region | X | va_coef

    where X is the sub-sector output in that region (from the "Total outputs"
    sheet) and va_coef is the workbook-level value-added coefficient rebuilt by
    summing Tax/TTM/Capital/Labor from the regional inventory sheet serving that
    GTAP region.

    If `mario_dir` is passed, the D2.4 inventories remain the default source but
    missing sub-sectors are taken from MARIO, and `prefer_mario_subsectors`
    forces selected sub-sectors to come from MARIO even if they exist in D2.4.

    Results are cached to `cache_path` (pickle); pass refresh=True to rebuild.
    """
    if cache_path is None:
        cache_path = DEFAULT_MIXED_CACHE_PATH if mario_dir else DEFAULT_CACHE_PATH

    expected_cols = {
        "region",
        "X",
        "va_coef",
        "subsector",
        "subsector_name",
        "parent_inventory",
        "inventory_source",
        "inventory_file",
    }
    if cache_path and not refresh and os.path.exists(cache_path):
        cached = pd.read_pickle(cache_path)
        if expected_cols.issubset(cached.columns):
            data = cached.copy()
        else:
            warnings.warn(f"Ignoring stale cache {cache_path}; rebuilding.")
            data = pd.DataFrame()
    else:
        data = pd.DataFrame()

    if data.empty:
        primary = _collect_inventory_dir(d24_dir, source="d24")
        fallback = (
            _collect_inventory_dir(mario_dir, source="mario")
            if mario_dir is not None
            else None
        )
        data = _select_inventory_sources(
            primary,
            fallback,
            prefer_fallback_subsectors=prefer_mario_subsectors,
        )

    # authoritative parent from GTAP12_matching (fallback: inventory parent)
    if matching_path and not data.empty:
        pmap = _parent_map_from_matching(matching_path)

        def _resolve(code, fallback):
            return pmap.get(_CODE_ALIASES.get(code, code), fallback)

        in_matching = data["subsector"].map(
            lambda c: _CODE_ALIASES.get(c, c) in pmap
        )
        missing = sorted(data.loc[~in_matching, "subsector"].unique())
        if missing:
            warnings.warn(
                "Sub-sectors not found in GTAP12_matching (kept inventory "
                f"parent): {missing}"
            )
        data["parent"] = [
            _resolve(c, p)
            for c, p in zip(data["subsector"], data["parent_inventory"])
        ]
    else:
        data["parent"] = data["parent_inventory"]

    if cache_path:
        try:
            data.to_pickle(cache_path)
        except Exception as exc:  # noqa: BLE001
            warnings.warn(f"Could not write cache {cache_path}: {exc}")
    return data


def check_parent_consistency(
    d24_dir: str = DEFAULT_D24_DIR,
    matching_path: str = DEFAULT_MATCHING_PATH,
    refresh: bool = False,
    only_mismatches: bool = True,
) -> pd.DataFrame:
    """
    Compare each inventory's declared parent ("Parent sector" in its Summary)
    against the authoritative GTAP12_matching.xlsx mapping.

    Returns a DataFrame: subsector | parent_inventory | parent_matching | ok
    (NaN parent_matching = sub-sector code absent from the matching workbook,
    possibly a code-spelling difference, see _CODE_ALIASES). With
    only_mismatches=True (default) only the rows to fix are returned.
    """
    split = collect_split_data(d24_dir, matching_path=None, refresh=refresh)
    inv = (
        split[["subsector", "parent_inventory"]]
        .drop_duplicates()
        .set_index("subsector")["parent_inventory"]
    )
    pmap = _parent_map_from_matching(matching_path)

    rows = []
    for code, parent_inv in inv.items():
        matched = pmap.get(_CODE_ALIASES.get(code, code))
        rows.append(
            {
                "subsector": code,
                "parent_inventory": parent_inv,
                "parent_matching": matched,
                "ok": matched == parent_inv,
            }
        )
    out = pd.DataFrame(rows).sort_values("subsector").reset_index(drop=True)
    if only_mismatches:
        out = out[~out["ok"]].reset_index(drop=True)
    return out


# --------------------------------------------------------------------------- #
# VA coefficients of the split database (v.parquet)
# --------------------------------------------------------------------------- #
def _va_coefficient(v_parquet_path: str) -> pd.Series:
    """Total value-added coefficient per producing (region, sector).

    v.parquet (mario coefficients export) has the value-added accounts on the
    rows and the producing sectors on the columns (a MultiIndex whose levels
    typically are Region / Level / Item). Summing over the rows gives the total
    VA generated per unit of output.

    Returns a Series indexed by a MultiIndex (region, sector).
    """
    v = pd.read_parquet(v_parquet_path)
    coef = v.sum(axis=0)  # total VA coefficient per producing column

    cols = coef.index
    if isinstance(cols, pd.MultiIndex):
        names = list(cols.names)
        region = (
            cols.get_level_values("Region")
            if "Region" in names
            else cols.get_level_values(0)
        )
        sector = (
            cols.get_level_values("Item")
            if "Item" in names
            else cols.get_level_values(-1)
        )
    else:  # single level -> assume it is the sector code, region unknown
        region = pd.Index(["GLOBAL"] * len(cols))
        sector = cols

    out = pd.Series(
        np.asarray(coef.values, dtype=float),
        index=pd.MultiIndex.from_arrays(
            [np.asarray(region), np.asarray(sector)], names=["region", "sector"]
        ),
    )
    return out.groupby(level=["region", "sector"]).sum()


# --------------------------------------------------------------------------- #
# Aggregate the chosen quantity per (parent, subsector)
# --------------------------------------------------------------------------- #
def _aggregate_quantity(
    split: pd.DataFrame,
    quantity: str,
    v_parquet_path: str,
    region: str | None = None,
) -> pd.DataFrame:
    """
    Collapse the long-form `split` table into one row per sub-sector carrying the
    chosen quantity:  parent | subsector | subsector_name | value

    If `region` is given, only that region's rows are used.
    """
    if region is not None:
        split = split[split["region"] == region]
    keys = ["parent", "subsector", "subsector_name"]

    if quantity == "X":
        agg = split.groupby(keys, as_index=False)["X"].sum()
        agg = agg.rename(columns={"X": "value"})
        return agg

    if "va_coef" in split.columns:
        split = split.copy()
        split["va_coef"] = pd.to_numeric(split["va_coef"], errors="coerce").fillna(0.0)
        split["value"] = split["va_coef"] * split["X"]
        return split.groupby(keys, as_index=False)["value"].sum()

    # quantity == "VA": value = sum_region  coef[region, subsector] * X[region]
    coef = _va_coefficient(v_parquet_path).rename("coef").reset_index()
    merged = split.merge(
        coef,
        left_on=["region", "subsector"],
        right_on=["region", "sector"],
        how="left",
    )
    missing = merged["coef"].isna()
    if missing.any():
        miss_subs = sorted(merged.loc[missing, "subsector"].unique())
        warnings.warn(
            "No VA coefficient found for some (region, sub-sector) pairs "
            f"(sub-sectors: {miss_subs}); treated as 0."
        )
    merged["coef"] = merged["coef"].fillna(0.0)
    merged["value"] = merged["coef"] * merged["X"]
    return merged.groupby(keys, as_index=False)["value"].sum()


# --------------------------------------------------------------------------- #
# Plot
# --------------------------------------------------------------------------- #
def plot_sector_split(
    sectors,
    gtap_x_path: str = DEFAULT_GTAP_X_PATH,
    va_path: str = DEFAULT_VA_PATH,
    v_parquet_path: str = DEFAULT_V_PARQUET_PATH,
    d24_dir: str = DEFAULT_D24_DIR,
    mario_dir: str | None = DEFAULT_MARIO_DIR,
    prefer_mario_subsectors=("CPP", "RCP"),
    region: str | None = None,
    ncols: int | None = None,
    normalize: bool = False,
    show_residual: bool = True,
    sharey: bool = True,
    residual_labels: dict | None = None,
    matching_path: str | None = DEFAULT_MATCHING_PATH,
    refresh_cache: bool = False,
    save_path: str | None = None,
    figsize_per_plot: tuple[float, float] = (3.4, 3.6),
    cmap: str = "tab20",
):
    """
    One subplot per parent GTAP sector. Each subplot shows the GTAP12 ENTICE
    disaggregation as two stacked columns -- total output X and total value added
    VA -- broken down into the carved-out sub-sectors plus a residual block. The
    group total is annotated on top of each column.

    Parameters
    ----------
    sectors : list[str] | str
        Parent GTAP sector codes to plot (e.g. ["EEQ", "MVH", "CHM"]). The
        subplot grid adapts automatically to the length of this list.
        Pass "all" to plot every parent sector that was actually split.
    gtap_x_path, va_path, v_parquet_path : str
        Sources for, respectively, original X, original VA, and the VA
        coefficients of the split database.
    d24_dir, mario_dir : str
        Inventory folders. D2.4 remains the default source; if mario_dir is
        provided, missing sub-sectors fall back to MARIO.
    prefer_mario_subsectors : iterable[str]
        Sub-sectors to force from MARIO even if they also exist in D2.4.
    region : str, optional
        GTAP region code (e.g. "ITA"). If given, plot that single region's values
        instead of the world totals.
    ncols : int, optional
        Number of subplot columns. Defaults to ceil(sqrt(n)).
    normalize : bool
        If True, each column is expressed as a share of its own total (%) instead
        of absolute T USD (the absolute total is still annotated on top).
    show_residual : bool
        If True, add the residual block (group total minus carved-out
        sub-sectors) to each column.
    sharey : bool
        If True (default), all subplots share the same y-axis scale.
    residual_labels : dict, optional
        Per-parent label for the residual block (default DEFAULT_RESIDUAL_LABELS,
        e.g. P_C -> PEP). Anything not listed falls back to "Rest of <parent>".
    refresh_cache : bool
        Force re-reading every inventory file instead of using the cache.
    save_path : str, optional
        If given, save the figure there (png/pdf/...).

    Returns
    -------
    (fig, axes)
    """
    # residual labels: defaults (e.g. P_C -> PEP) overridden by caller
    res_labels = {**DEFAULT_RESIDUAL_LABELS, **(residual_labels or {})}

    # original (pre-split) group totals per parent sector (optionally one region)
    original = {
        "X": _read_original_X(gtap_x_path, region=region),
        "VA": _read_original_VA(va_path, region=region),
    }
    if region is not None and original["X"].empty and original["VA"].empty:
        warnings.warn(f"No data found for region {region!r}.")

    # sub-sector X and VA, merged side by side (one row per parent/sub-sector)
    split = collect_split_data(
        d24_dir=d24_dir,
        mario_dir=mario_dir,
        prefer_mario_subsectors=prefer_mario_subsectors,
        matching_path=matching_path,
        refresh=refresh_cache,
    )
    agg_X = _aggregate_quantity(split, "X", v_parquet_path, region=region)
    agg_VA = _aggregate_quantity(split, "VA", v_parquet_path, region=region)
    merged = agg_X.rename(columns={"value": "X"}).merge(
        agg_VA.rename(columns={"value": "VA"})[["parent", "subsector", "VA"]],
        on=["parent", "subsector"],
        how="left",
    )
    merged["VA"] = merged["VA"].fillna(0.0)

    # "all" -> every parent sector that was actually split
    if isinstance(sectors, str) and sectors.lower() == "all":
        sectors = merged["parent"].dropna().unique().tolist()
        # subplots ordered alphabetically only for the automatic "all" mode
        sectors = sorted(sectors)
    elif isinstance(sectors, str):
        sectors = [sectors]
    else:
        # preserve caller order while dropping duplicates
        sectors = list(dict.fromkeys(sectors))

    ylabel = "Share of column total [%]" if normalize else "T USD"
    columns = ["X", "VA"]  # x=0 -> X, x=1 -> VA

    # ----- grid geometry --------------------------------------------------- #
    n = len(sectors)
    if ncols is None:
        # prefer a taller grid (more rows than columns) when n is not a square
        ncols = max(1, math.floor(math.sqrt(n)))
    nrows = math.ceil(n / ncols)

    fig, axes = plt.subplots(
        nrows,
        ncols,
        figsize=(figsize_per_plot[0] * ncols, figsize_per_plot[1] * nrows),
        squeeze=False,
        sharey=sharey,
    )
    axes_flat = axes.ravel()
    cmap_obj = plt.get_cmap(cmap)

    tops: dict = {}  # ax -> tallest bar (for y-headroom so totals don't clip)

    for ax, parent in zip(axes_flat, sectors):
        sub = merged[merged["parent"] == parent]
        # drop sub-sectors that are ~0 in both X and VA (e.g. an empty inventory
        # such as RBR, which is instead shown as the RPP residual)
        sub = sub[(sub["X"].abs() > 1e-9) | (sub["VA"].abs() > 1e-9)]
        sub = sub.sort_values("X", ascending=False).reset_index(drop=True)
        orig = {q: original[q].get(parent, np.nan) for q in columns}

        if all(np.isnan(orig[q]) for q in columns):
            ax.set_title(f"{parent}\n(not found)", fontsize=9)
            ax.axis("off")
            continue

        bar_top = 0.0
        for xpos, q in enumerate(columns):
            o = orig[q]
            if np.isnan(o):
                continue
            sub_sum = sub[q].sum()
            residual = max(o - sub_sum, 0.0)
            scale = 100.0 / o if (normalize and o) else 1e-6
            add_legend = xpos == 0  # legend entries only once, from the X column

            bottom = 0.0
            for i, row in sub.iterrows():
                ax.bar(
                    xpos,
                    row[q] * scale,
                    width=0.6,
                    bottom=bottom,
                    color=cmap_obj(i % cmap_obj.N),
                    edgecolor="white",
                    label=row["subsector"] if add_legend else "_nolegend_",
                )
                bottom += row[q] * scale
            if show_residual and residual > 0:
                # a named residual (e.g. P_C -> PEP) is a real sub-sector: solid
                # colour, no hatch; a generic "Rest of <parent>" stays grey hatched
                named_residual = parent in res_labels
                ax.bar(
                    xpos,
                    residual * scale,
                    width=0.6,
                    bottom=bottom,
                    color=cmap_obj(len(sub) % cmap_obj.N) if named_residual else "0.55",
                    edgecolor="white",
                    hatch=None if named_residual else "//",
                    label=(
                        res_labels.get(parent, f"Rest of {parent}")
                        if add_legend
                        else "_nolegend_"
                    ),
                )
                bottom += residual * scale

            # total on top of the column (always the absolute group total, T USD)
            ax.annotate(
                f"{o * 1e-6:.2f}",
                (xpos, bottom),
                textcoords="offset points",
                xytext=(0, 2.5),
                ha="center",
                va="bottom",
                fontsize=7,
            )
            bar_top = max(bar_top, bottom)

        tops[ax] = bar_top

        # ---- cosmetics ---------------------------------------------------- #
        ax.set_xticks([0, 1])
        ax.set_xticklabels(["Production", "Value added"], fontsize=8)
        ax.set_title(parent, fontsize=10)  # three-letter code only
        # widen the x-axis so the legend sits in the empty space to the right of
        # the two bars (bars span x in [-0.3, 1.3]) and never overlaps them
        ax.set_xlim(-0.6, 2.7)
        ax.legend(
            fontsize=6.5,
            loc="upper right",
            bbox_to_anchor=(1.0, 1.0),
            frameon=False,
            handlelength=1.2,
            borderaxespad=0.2,
        )

    # hide any unused axes
    for ax in axes_flat[n:]:
        ax.axis("off")

    # y-headroom so the on-top totals are not clipped
    if tops:
        if sharey:
            axes_flat[0].set_ylim(0, max(tops.values()) * 1.15)
        else:
            for a, t in tops.items():
                a.set_ylim(0, t * 1.15)

    # y-axis label on the whole first column (every used row of column 0)
    for r in range(nrows):
        if r * ncols < n:  # axis is in use
            axes[r, 0].set_ylabel(ylabel)

    fig.tight_layout(pad=0.5, w_pad=0.4, h_pad=0.8)

    if save_path:
        fig.savefig(save_path, dpi=200, bbox_inches="tight")
        print(f"Saved figure to {save_path}")

    return fig, axes


# --------------------------------------------------------------------------- #
# Example
# --------------------------------------------------------------------------- #
if __name__ == "__main__":
    plot_sector_split("all", save_path=str(PROJECT_ROOT / "d24_split_X_VA.png"))
    plt.show()
