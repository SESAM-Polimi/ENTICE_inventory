# ENTICE inventory — GTAP12 disaggregation toolkit (D2.4)

Tooling for the ENTICE **D2.4 — Enhanced GTAP database** deliverable. It turns
sector-level cost inventories into MARIO `add_sector` workbooks, applies them to
the GTAP database, and exports the per-sector D2.4 inventory workbooks plus QA
plots.

The [maintained project report](docs/baseline/README.md) is the single place for
baseline findings, current status, agreed decisions and next steps towards D2.8
and T2.6. This README covers installation and usage. The report is updated in
place; new narrative audit/follow-up reports are not created. Historical build
reproduction remains unverified; preserved source and release manifests are in
`manifests/`.

`data/GTAP12_matching.xlsx` currently supplies the runner's GTAP parent mapping.
It is not an independently validated NACE registry: the reconstruction identified
aliases and classification/scope conflicts. New consumers use the canonical
registry described below. The historical parent mapping has not been replaced
globally: parent-dependent coefficients and quantities must be rebuilt together.

For one-sector re-export from saved coefficients, use
`python scripts/export_d24_sector.py --help`. The
[recovery guide](docs/baseline/README.md#regenerating-one-sector)
includes a concrete command. The export-only dependencies are pinned in
`requirements/d24-export.txt`; they do not certify the upstream solver environment.

## Layout

```text
data/                         bundled input workbooks (matching, GTAP12_X, region clusters)
data/registry/                sector identities, GTAP12 adapter and explicit region groups
scripts/run_d24_pipeline.py   end-to-end driver (Jupyter #%% cells)
src/entice_inventory/
  core/        matching_utils, residual_sector_cluster, paths
  inventory/   template_/exioiot_/copcosts_to_inventory, inventory_writer, template_checker
  export/      build_d24_inventories, plot_d24_split
```

## Install

```bash
pip install -e .            # core (openpyxl)
pip install -e '.[plots]'   # + pandas/numpy/matplotlib for the QA plots
```

For the tested MARIO/CVXLAB workflow, use `python3.13 scripts/bootstrap_producer.py`
to create an isolated environment from pinned revisions, dependencies and explicit
compatibility patches. See the [producer verification and commands](docs/baseline/README.md#producer-environment-and-representative-verification)
for the representative HYE check and its remaining numerical warnings. This is
separate from the lightweight export-only installation.

## Two pipelines

### 1. Inventory creation — `entice_inventory.inventory`

Build `Add_sector_*.xlsx`. Three entry points, one per data source, each
exposing `make_inventory(...)`:

| Module | Source | Sectors |
| --- | --- | --- |
| `template_to_inventory` | ENTICE data-collection Excel templates | hand-collected inventories |
| `exioiot_to_inventory` | `EXIOIOT.csv` (Purdue) | sectors flagged `Pipeline = EXIOIOT` |
| `copcosts_to_inventory` | `COPCOSTS.csv` (Purdue) | copper sectors not in EXIOIOT |

Shared building blocks: `inventory_writer` (`write_inventory_workbook`, the one
place that emits the shared `Add_sector_*.xlsx` layout), `template_checker`
(parse/validate templates), `core.residual_sector_cluster` (append the "rest of
parent" residual clusters), `core.matching_utils` (read `GTAP12_matching.xlsx`).

### 2. D2.4 export — `entice_inventory.export`

| Module | Role |
| --- | --- |
| `build_d24_inventories` | `export_d24_inventories(...)` + `update_trades_in_exported_inventories(...)` — read the applied MARIO database and write the per-sector D2.4 workbooks. |
| `plot_d24_split` | `plot_sector_split(...)` — QA plots of the X / VA disaggregation per parent sector. |

## Running it

End-to-end: open `scripts/run_d24_pipeline.py`, set `BASE` to your local OneDrive
copy of the WP2 data folder, and run the `#%%` cells in order.

The modules with a `__main__` block can also be run from the command line, e.g.:

```bash
python -m entice_inventory.export.build_d24_inventories --help
python -m entice_inventory.inventory.template_to_inventory --help
```

## Canonical registry

Three editable JSON files have distinct roles: `sectors.json` holds activity
identities, source aliases and NACE Rev. 2 scopes; `gtap12.json` holds the reviewed
NACE-to-GTAP bridge and historical mappings; `regions.json` holds geography and
explicit cluster membership. They include source references and hashes. The
older NACE CSV is preserved review evidence, not a second live mapping.

```bash
python scripts/registry.py sector BVL
python scripts/registry.py sector WNT --scope turbine_sets
python scripts/registry.py region GLOBAL
python scripts/registry.py --output build/registry-check.json check
python scripts/registry.py --output build/hye-check.json inspect-inventory '/path/to/Add_sector_HYE.xlsx'
```

Use `--registry-dir /path/to/registry` before the subcommand to supply an explicit
registry location. Outputs are JSON; `--output` refuses to overwrite an existing
file and prints a concise summary. Diagnostic checks are optional: `check
--checks classification regions` runs only those checks; `check --checks` skips
all optional checks and lists them as skipped. Invalid registry structure,
unreadable workbooks and ambiguous lookups produce a failed result and exit 1.
Ordinary diagnostic warnings return normally, with an action for each finding.

Python consumers can call `Registry.load()`, `registry.sector(code)`,
`registry.parent(code, target='GTAP12', scope=None)` and
`registry.members(group, geography='GTAP12')` from
`entice_inventory.core.registry`. An unresolved parent is `None` with an explicit
status; there is no fallback to a published or legacy parent. A scope selected
for a mixed activity gives a conditional mapping, not proof that its source
observations match that activity. Querying never changes data or activates a
catalogue-only sector. EXIOBASE has no adapter yet.

The template importer now takes its GTAP12 region universe from this registry,
including MRT. Its optional `registry_dir` argument selects a different registry
directory. Regenerate affected intermediates before rebuilding coefficients;
editing a published `GLOBAL` list alone would leave derived values stale.
Newly restored regions have output estimation notes in `Total outputs`.

The GTAP baseline is preserved as supplied. Producer verification diagnoses its
existing residuals without correcting matrices; the earlier experimental
baseline-closure option has been removed. Full build acceptance remains open.

## Notes

- `Run.py` + `paths.yml` are a personal scratch runner for the inventory-creation
  pipeline (per-user paths); `Run.py` is git-ignored.
- `.d24_split_cache.pkl` / `.d24_split_mixed_cache.pkl` are regenerable caches for
  `plot_d24_split` (`refresh_cache=True` rebuilds them); git-ignored.
