"""ENTICE inventory toolkit — GTAP12 disaggregation for deliverable D2.4.

Sub-packages:
    core       shared helpers (matching workbook, residual clusters, paths)
    inventory  build Add_sector_*.xlsx inventories from templates / EXIOIOT / COPCOSTS
    export     read the applied MARIO database and write the D2.4 workbooks + QA plots
"""

from entice_inventory.core.paths import DATA_DIR, PROJECT_ROOT

__all__ = ["DATA_DIR", "PROJECT_ROOT"]
