from __future__ import annotations

from collections import OrderedDict, defaultdict
from dataclasses import dataclass
from pathlib import Path

from matching_utils import load_matching_rows


def _string(value) -> str:
    return str(value).strip() if value is not None else ''


def _upper(value) -> str:
    return _string(value).upper()


def _dedupe(values) -> list[str]:
    seen: set[str] = set()
    ordered: list[str] = []
    for value in values:
        if value and value not in seen:
            ordered.append(value)
            seen.add(value)
    return ordered


@dataclass(frozen=True)
class ResidualClusterUpdate:
    sheet_name: str
    cluster_name: str
    missing_codes: list[str]
    cluster_members: list[str]


def _read_sector_clusters(ws) -> OrderedDict[str, list[str]]:
    clusters: OrderedDict[str, list[str]] = OrderedDict()
    for column in range(1, ws.max_column + 1):
        cluster_name = _string(ws.cell(1, column).value)
        if not cluster_name:
            continue

        members: list[str] = []
        for row in range(2, ws.max_row + 1):
            member = _string(ws.cell(row, column).value)
            if member:
                members.append(member)
        clusters[cluster_name] = members
    return clusters


def _append_sector_cluster(ws, cluster_name: str, members: list[str]) -> None:
    if ws.max_column == 1 and ws.cell(1, 1).value is None:
        next_column = 1
    else:
        next_column = ws.max_column + 1

    ws.cell(1, next_column).value = cluster_name
    for row_offset, member in enumerate(members, start=2):
        ws.cell(row_offset, next_column).value = member


def _cluster_signature(members: list[str]) -> tuple[str, ...]:
    return tuple(sorted({_upper(member) for member in members if _string(member)}))


def _unique_cluster_name(existing_names: set[str], base_name: str) -> str:
    if base_name not in existing_names:
        return base_name

    counter = 2
    while True:
        candidate = f"{base_name}_{counter}"
        if candidate not in existing_names:
            return candidate
        counter += 1


def _scale_inventory_rows(ws, scale_factor: float) -> None:
    for row in range(2, ws.max_row + 1):
        item_type = _string(ws.cell(row, 4).value)
        if item_type not in {'Sector', 'Factor of production'}:
            continue

        change_type = _string(ws.cell(row, 7).value)
        if change_type and change_type != 'Update':
            continue

        quantity_cell = ws.cell(row, 1)
        try:
            quantity = float(quantity_cell.value)
        except (TypeError, ValueError):
            continue

        quantity_cell.value = quantity * scale_factor


def _expand_member_codes(
    member: str,
    clusters: OrderedDict[str, list[str]],
    gtap12_to_entice_codes: dict[str, list[str]],
    direct_entice_codes: set[str],
    seen: set[str] | None = None,
) -> set[str]:
    if seen is None:
        seen = set()

    normalized = _upper(member)
    if not normalized or normalized in seen:
        return set()
    seen.add(normalized)

    if normalized in clusters:
        covered: set[str] = set()
        for nested_member in clusters[normalized]:
            covered.update(
                _expand_member_codes(
                    nested_member,
                    clusters,
                    gtap12_to_entice_codes,
                    direct_entice_codes,
                    seen,
                )
            )
        return covered

    if normalized in direct_entice_codes:
        return {normalized}

    return set(gtap12_to_entice_codes.get(normalized, []))


def apply_residual_sector_clusters(
    workbook,
    repo_path: str | Path,
    residual_share: float,
    sector_code: str | None = None,
) -> list[ResidualClusterUpdate]:
    if residual_share <= 0:
        return []
    if residual_share >= 1:
        raise ValueError('residual_share must be >= 0 and < 1.')

    repo_path = Path(repo_path)
    sector_rows, _ = load_matching_rows(repo_path)

    all_entice_codes = _dedupe(
        _upper(row.entice_code)
        for row in sector_rows
        if _string(row.entice_code)
    )
    all_entice_code_set = set(all_entice_codes)

    gtap12_to_entice_codes: dict[str, list[str]] = defaultdict(list)
    entice_to_gtap12: dict[str, str] = {}
    for row in sector_rows:
        entice_code = _upper(row.entice_code)
        gtap12_lookup = _upper(row.gtap12)
        gtap12_display = _string(row.gtap12)
        if not entice_code:
            continue

        if gtap12_lookup and entice_code not in gtap12_to_entice_codes[gtap12_lookup]:
            gtap12_to_entice_codes[gtap12_lookup].append(entice_code)
        entice_to_gtap12[entice_code] = gtap12_display

    excluded_codes = {_upper(sector_code)} if _string(sector_code) else set()
    universe_codes = [code for code in all_entice_codes if code not in excluded_codes]

    master_ws = workbook['Master']
    sectors_clusters_ws = workbook['Sectors Clusters']
    sectors_clusters = OrderedDict(
        (_upper(name), members)
        for name, members in _read_sector_clusters(sectors_clusters_ws).items()
    )
    existing_cluster_names = set(sectors_clusters.keys())
    cluster_name_by_signature = {
        _cluster_signature(members): cluster_name
        for cluster_name, members in sectors_clusters.items()
        if _cluster_signature(members)
    }

    updates: list[ResidualClusterUpdate] = []
    scale_factor = 1.0 - residual_share

    for row in range(2, master_ws.max_row + 1):
        sheet_name = _string(master_ws.cell(row, 3).value)
        if not sheet_name or sheet_name not in workbook.sheetnames:
            continue

        ws_inv = workbook[sheet_name]
        covered_codes: set[str] = set()
        for inv_row in range(2, ws_inv.max_row + 1):
            item_type = _string(ws_inv.cell(inv_row, 4).value)
            if item_type != 'Sector':
                continue

            db_item = _string(ws_inv.cell(inv_row, 5).value)
            if not db_item:
                continue

            covered_codes.update(
                _expand_member_codes(
                    db_item,
                    sectors_clusters,
                    gtap12_to_entice_codes,
                    all_entice_code_set,
                )
            )

        missing_codes = [code for code in universe_codes if code not in covered_codes]
        if not missing_codes:
            continue

        cluster_members = _dedupe(
            entice_to_gtap12.get(code, code)
            for code in missing_codes
            if entice_to_gtap12.get(code, code)
        )
        if not cluster_members:
            continue

        cluster_signature = _cluster_signature(cluster_members)
        cluster_name = cluster_name_by_signature.get(cluster_signature)
        if cluster_name is None:
            base_name = _upper(f"ALL_OTHER_{sheet_name}")
            cluster_name = _unique_cluster_name(existing_cluster_names, base_name)
            existing_cluster_names.add(cluster_name)
            sectors_clusters[cluster_name] = cluster_members
            cluster_name_by_signature[cluster_signature] = cluster_name
            _append_sector_cluster(sectors_clusters_ws, cluster_name, cluster_members)

        _scale_inventory_rows(ws_inv, scale_factor)
        ws_inv.append([residual_share, 'M USD', None, 'Sector', cluster_name, 'GLOBAL', 'Update'])

        updates.append(
            ResidualClusterUpdate(
                sheet_name=sheet_name,
                cluster_name=cluster_name,
                missing_codes=missing_codes,
                cluster_members=cluster_members,
            )
        )

    return updates