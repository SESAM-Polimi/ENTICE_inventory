"""D2.4 builder — end-to-end driver for the enhanced GTAP database (D2.4).

Run as Jupyter-style cells (#%%). The flow is:

  1. Parse the baseline GTAP database (MARIO, flows).
  2. Read the MARIO add_sector inventories and apply them (add_sectors).
  3. Save the enhanced database, then re-parse it as coefficients.
  4. Export the per-sector D2.4 inventory workbooks (export_d24_inventories).
  5. Inject Purdue bilateral trade data into the exported workbooks.
  6. Plot the X / VA disaggregation of each split parent sector for QA.

Set ENTICE_DATA_ROOT or data_root in paths.local.json to the shared ENTICE
inventory archive. Every operational path is derived from that directory.
"""

#%%
from pathlib import Path
from datetime import datetime

import mario

from entice_inventory.export.build_d24_inventories import (
    export_d24_inventories,
    update_trades_in_exported_inventories,
)
from entice_inventory.export import plot_d24_split
from entice_inventory.core.paths import project_data_root, data_path, PROJECT_ROOT
from entice_inventory.core.baseline import load_baseline

# --- Paths (all derived from BASE) ----------------------------------------- #
BASE = project_data_root(required=True)
RUN_ID = datetime.now().strftime('%Y%m%d-%H%M%S')
GTAP_DB = data_path('baseline', required=True)
MARIO_INVENTORIES = data_path('mario_inventories', required=True)
D24_INVENTORIES = data_path('inventory_runs') / RUN_ID
# Numerical work files and logs stay outside the partner-facing output folder.
D24_WORK = PROJECT_ROOT / 'build' / 'runs' / RUN_ID
D24_DATABASE = D24_WORK / 'database'
PURDUE_TRADE = data_path('purdue_trade', required=True)
D24_INVENTORIES.mkdir(parents=True, exist_ok=False)
D24_WORK.mkdir(parents=True, exist_ok=False)

#%%
# 1. Baseline GTAP database (flows).
db = load_baseline(GTAP_DB)

#%%
# 2. Read the MARIO add_sector inventories and apply them.
db.read_add_sectors_excel(
    path=str(MARIO_INVENTORIES),
    read_inventories=True,
    split=False,
)
db.add_sectors(split=False, VA_fix=True, accept_non_unitary_sum=True)

#%%
# 3. Save the enhanced database, then re-parse it as coefficients.
db.to_parquet(str(D24_DATABASE), coefficients=True, flows=False)
db = mario.parse_from_parquet(
    str(D24_DATABASE / "coefficients"), table="IOT", mode="coefficients"
)

#%%
# 4. Export the per-sector D2.4 inventory workbooks.
export_d24_inventories(
    db,
    source_dir=str(MARIO_INVENTORIES),
    output_dir=str(D24_INVENTORIES),
    tolerance=1e-12,
    inventory_sum_check_tolerance=0.01,
    report_dir=D24_WORK,
)

#%%
# 5. Inject Purdue bilateral trade data into the exported workbooks.
updated_files = update_trades_in_exported_inventories(
    output_dir=str(D24_INVENTORIES),
    purdue_trade_path=str(PURDUE_TRADE),
)

#%%
# 6. QA plot: X / VA disaggregation of each split parent sector.
plot_d24_split.plot_sector_split(
    [
        "OXT", "P_C", "CHM", "RPP", "NMM",
        "I_S", "NFM", "ELE", "EEQ", "MVH", "OTP",
    ],
    mario_dir=str(MARIO_INVENTORIES),
    prefer_mario_subsectors=("CPP", "RCP"),
    refresh_cache=True,
    save_path="X_VA.png",
    sharey=True,
)
