#%%
import mario
from build_d24_inventories import export_d24_inventories

#%%
db_path='/Users/lorenzorinaldi/Library/CloudStorage/OneDrive-SharedLibraries-eNextGen/ENTICE - Documents/WPs, Tasks & Deliverables/WP2 - Data/T2.2 & T2.3 - GTAP disaggregation/Database/GTAP 2023/2023entice'
# db_path=r"C:\Users\camic\eNextGen\ENTICE - Documenti\WPs, Tasks & Deliverables\WP2 - Data\T2.2 & T2.3 - GTAP disaggregation\Database\GTAP 2023\2023_start"
db = mario.parse_from_parquet(
    db_path,
    table="IOT",
    mode="flows"
    )

# %%
db.read_add_sectors_excel(
    path = "/Users/lorenzorinaldi/Library/CloudStorage/OneDrive-SharedLibraries-eNextGen/ENTICE - Documents/WPs, Tasks & Deliverables/WP2 - Data/T2.2 & T2.3 - GTAP disaggregation/Data collection/Inventory cleaning/MARIO inventories",
    # path = r"C:\Users\camic\eNextGen\ENTICE - Documenti\WPs, Tasks & Deliverables\WP2 - Data\T2.2 & T2.3 - GTAP disaggregation\Data collection\Inventory cleaning\MARIO inventories\Add_sector_Manufacture of solar panels.xlsx",
    read_inventories = True,
    split = False,
)

# %%
db.add_sectors(
    split=False,
    VA_fix=True,
    accept_non_unitary_sum=True,
    )

#%%
db.to_parquet(
    "/Users/lorenzorinaldi/Library/CloudStorage/OneDrive-SharedLibraries-eNextGen/ENTICE - Documents/WPs, Tasks & Deliverables/WP2 - Data/T2.2 & T2.3 - GTAP disaggregation/Data collection/Inventory cleaning/D2.4 database", 
    coefficients=True, 
    flows=False
    )

#%%
db = mario.parse_from_parquet(
    "/Users/lorenzorinaldi/Library/CloudStorage/OneDrive-SharedLibraries-eNextGen/ENTICE - Documents/WPs, Tasks & Deliverables/WP2 - Data/T2.2 & T2.3 - GTAP disaggregation/Data collection/Inventory cleaning/D2.4 database/coefficients",
    table="IOT",
    mode="coefficients"
    )

# %%
export_d24_inventories(
    db,
    source_dir="/Users/lorenzorinaldi/Library/CloudStorage/OneDrive-SharedLibraries-eNextGen/ENTICE - Documents/WPs, Tasks & Deliverables/WP2 - Data/T2.2 & T2.3 - GTAP disaggregation/Data collection/Inventory cleaning/MARIO inventories",
    output_dir="/Users/lorenzorinaldi/Library/CloudStorage/OneDrive-SharedLibraries-eNextGen/ENTICE - Documents/WPs, Tasks & Deliverables/WP2 - Data/T2.2 & T2.3 - GTAP disaggregation/Data collection/Inventory cleaning/D2.4 inventories",
    tolerance=1e-12,
    inventory_sum_check_tolerance=0.01
)
# %%
