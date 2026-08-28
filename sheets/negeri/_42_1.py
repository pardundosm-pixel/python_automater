import pandas as pd
from src.data_provider import get_metrics_dict
from src.excel_utils import safe_write

# ==========================================
# 1. MAPPING & PAGINATION CONFIGURATION
# ==========================================
# TODO: Update these column indices based on where data lives in the new Jadual
COL_DISTRICT = 2  # e.g., Column B
COL_YEAR = 5      # e.g., Column E
COL_VALUE = 8     # e.g., Column H
COL_PCT = 9       # e.g., Column I

# TODO: Adjust the years being mapped. 
# The dictionary key (0, 1, 2) represents how many rows down from the block's start row the year sits.
YEAR_OFFSETS = {
    0: "2018",
    1: "2019",
    2: "2020"
}

# TODO: Define the exact row where the State (Negeri) total is printed.
STATE_START_ROW = 11

# TODO: Recalculate these block start rows based on the new template's layout.
# - The first number in the bracket (e.g., 15) is the starting row for that specific page.
# - The multiplier (e.g., 4) is the height of each district block (3 years + 1 blank separator row).
# - The range (e.g., 15) is the maximum number of districts that can fit on that page.
BLOCK_START_ROWS = (
    [15 + (i * 4) for i in range(15)] +  # Page 1 blocks
    [85 + (i * 4) for i in range(13)] +  # Page 2 blocks
    [151 + (i * 4) for i in range(14)]   # Page 3 blocks
)

# TODO: Update these cell coordinates if the titles are placed in different cells in the new template.
TITLE_COORDINATES = [
    {"bm": (2, 3), "en": (3, 3), "is_samb": False},     # Page 1 Title
    {"bm": (76, 3), "en": (77, 3), "is_samb": True},    # Page 2 Title
    {"bm": (142, 3), "en": (143, 3), "is_samb": True}   # Page 3 Title
]

# ==========================================
# 2. REPORT INJECTION ENGINE
# ==========================================
def populate_jadual_42_1(sheet, hierarchy, report_type):
    state_name = hierarchy.get('state_name', 'Unknown State')
    state_code = hierarchy.get('state_code', '00')
    districts = hierarchy.get('districts', [])
    print(f"  -> Populating Jadual 42.1 (KDNK Daerah) untuk {state_name}")

    # 1. Title Generation
    title_bm_base = f": Keluaran Dalam Negeri Kasar (KDNK) mengikut daerah pentadbiran, {state_name}, 2018 - 2020"
    title_en_base = f": Gross Domestic Product (GDP) by administrative district, {state_name}, 2018 - 2020"
    
    for coords in TITLE_COORDINATES:
        suffix_bm = " (samb.)" if coords["is_samb"] else ""
        suffix_en = " (cont'd)" if coords["is_samb"] else ""
        safe_write(sheet, coords["bm"][0], coords["bm"][1], title_bm_base + suffix_bm)
        safe_write(sheet, coords["en"][0], coords["en"][1], title_en_base + suffix_en)

    # 2. State-Level Data Injection
    safe_write(sheet, STATE_START_ROW, COL_DISTRICT, state_name.upper())
    
    # Fetch State Total using the specialized empty-cell trigger in the daerah table
    state_metrics = get_metrics_dict("STATE_TOTAL", level='daerah', parent_code=state_code)
    
    for offset, year in YEAR_OFFSETS.items():
        target_row = STATE_START_ROW + offset
        year_data = state_metrics.get(year, {})
        
        val_kdnk = year_data.get("kdnk_mengikut_jenis_aktiviti_ekonomi_pada_harga_malar_2015_rm_juta", "n.a")
        val_pct = year_data.get("peratus_perubahan_peratusan_tahunan", "n.a")
        
        final_kdnk = float(val_kdnk) if pd.notna(val_kdnk) and val_kdnk not in ["n.a", ""] else "n.a"
        final_pct = float(val_pct) if pd.notna(val_pct) and val_pct not in ["n.a", ""] else "n.a"
        
        safe_write(sheet, target_row, COL_VALUE, final_kdnk)
        safe_write(sheet, target_row, COL_PCT, final_pct)

    # 3. District-Level Data Injection & Cleanup 
    total_districts = len(districts)
    
    for district_idx, base_row in enumerate(BLOCK_START_ROWS):
        if district_idx < total_districts:
            # --- Inject District Data ---
            dist_code = districts[district_idx]['code']
            safe_write(sheet, base_row, COL_DISTRICT, districts[district_idx]['name'])
            
            dist_metrics = get_metrics_dict(dist_code, level='daerah', parent_code=state_code)
            
            for offset, year in YEAR_OFFSETS.items():
                target_row = base_row + offset
                year_data = dist_metrics.get(year, {})
                
                val_kdnk = year_data.get("kdnk_mengikut_jenis_aktiviti_ekonomi_pada_harga_malar_2015_rm_juta", "n.a")
                val_pct = year_data.get("peratus_perubahan_peratusan_tahunan", "n.a")
                
                final_kdnk = float(val_kdnk) if pd.notna(val_kdnk) and val_kdnk not in ["n.a", ""] else "n.a"
                final_pct = float(val_pct) if pd.notna(val_pct) and val_pct not in ["n.a", ""] else "n.a"
                
                safe_write(sheet, target_row, COL_VALUE, final_kdnk)
                safe_write(sheet, target_row, COL_PCT, final_pct)
        else:
            # --- Openpyxl Safe XML Wiping ---
            # Iterates through every row in the block and safely wipes columns B through I
            for r in range(base_row, base_row + len(YEAR_OFFSETS)):
                for c in range(COL_DISTRICT, COL_PCT + 1):
                    safe_write(sheet, r, c, None)

    # 4. XML-Level Page Deletion
    # Syntax: sheet.delete_rows(start_row_index, number_of_rows_to_delete)
    # Deleting ~225 rows safely clears out all artifacts to the bottom of the template.
    if total_districts <= 15:
        print("     [FORMAT] State only needs 1 page. Eliminating Page 2 and 3.")
        sheet.delete_rows(75, 225) 
        
    elif total_districts <= 28:
        print("     [FORMAT] State needs 2 pages. Eliminating Page 3.")
        sheet.delete_rows(141, 160)