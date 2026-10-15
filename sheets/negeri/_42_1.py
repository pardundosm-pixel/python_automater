import pandas as pd
from src.data_provider import get_metrics_dict

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
    {"bm": "C2", "en": "C3", "is_samb": False},     # Page 1 Title
    {"bm": "C76", "en": "C77", "is_samb": True},    # Page 2 Title
    {"bm": "C142", "en": "C143", "is_samb": True}   # Page 3 Title
]

# ==========================================
# 2. REPORT INJECTION ENGINE
# ==========================================
# TODO: Rename this function to match your new Jadual (e.g., populate_jadual_54_0)
def populate_jadual_42_1(sheet, hierarchy, report_type):
    state_name = hierarchy.get('state_name', 'Unknown State')
    state_code = hierarchy.get('state_code', '00')
    districts = hierarchy.get('districts', [])

    # 1. Title Generation
    # TODO: Update the base title text to reflect the metrics of the new Jadual
    title_bm_base = f": Keluaran Dalam Negeri Kasar (KDNK) mengikut daerah pentadbiran, {state_name}, 2018 - 2020"
    title_en_base = f": Gross Domestic Product (GDP) by administrative district, {state_name}, 2018 - 2020"
    
    for coords in TITLE_COORDINATES:
        suffix_bm = " (samb.)" if coords["is_samb"] else ""
        suffix_en = " (cont'd)" if coords["is_samb"] else ""
        sheet.range(coords["bm"]).value = title_bm_base + suffix_bm
        sheet.range(coords["en"]).value = title_en_base + suffix_en

    # 2. State-Level Data Injection
    sheet.range((STATE_START_ROW, COL_DISTRICT)).value = state_name.upper()
    state_metrics = get_metrics_dict("STATE_TOTAL", level='daerah', parent_code=state_code)
    
    for offset, year in YEAR_OFFSETS.items():
        target_row = STATE_START_ROW + offset
        year_data = state_metrics.get(year, {})
        
        # TODO: Change the dictionary keys here to match the specific database metrics you want to pull
        val_kdnk = year_data.get("kdnk_mengikut_jenis_aktiviti_ekonomi_pada_harga_malar_2015_rm_juta", "n.a")
        val_pct = year_data.get("peratus_perubahan_peratusan_tahunan", "n.a")
        
        sheet.range((target_row, COL_VALUE)).value = float(val_kdnk) if pd.notna(val_kdnk) and val_kdnk not in ["n.a", ""] else "n.a"
        sheet.range((target_row, COL_PCT)).value = float(val_pct) if pd.notna(val_pct) and val_pct not in ["n.a", ""] else "n.a"

    # 3. District-Level Data Injection & Cleanup 
    total_districts = len(districts)
    
    for district_idx, base_row in enumerate(BLOCK_START_ROWS):
        if district_idx < total_districts:
            dist_code = districts[district_idx]['code']
            sheet.range((base_row, COL_DISTRICT)).value = districts[district_idx]['name']
            
            dist_metrics = get_metrics_dict(dist_code, level='daerah', parent_code=state_code)
            
            for offset, year in YEAR_OFFSETS.items():
                target_row = base_row + offset
                year_data = dist_metrics.get(year, {})
                
                # TODO: Change these dictionary keys to match the state-level metrics above
                val_kdnk = year_data.get("kdnk_mengikut_jenis_aktiviti_ekonomi_pada_harga_malar_2015_rm_juta", "n.a")
                val_pct = year_data.get("peratus_perubahan_peratusan_tahunan", "n.a")
                
                # TODO: If your new Jadual has different columns or requires less variables, update this writing block
                sheet.range((target_row, COL_VALUE)).value = float(val_kdnk) if pd.notna(val_kdnk) and val_kdnk not in ["n.a", ""] else "n.a"
                sheet.range((target_row, COL_PCT)).value = float(val_pct) if pd.notna(val_pct) and val_pct not in ["n.a", ""] else "n.a"
        else:
            # Wipe the dummy row clean if no district exists for this slot
            sheet.range((base_row, COL_DISTRICT), (base_row + len(YEAR_OFFSETS), COL_PCT)).value = None

    # 4. Dynamic Page Elimination (Cleanup unused pages)
    # TODO: Adjust the `total_districts` threshold limits based on how many blocks fit on your new template's pages.
    # TODO: Adjust the `.delete()` row ranges (e.g., '75:300') to precisely target where Page 2 and Page 3 begin in the new template.
    if total_districts <= 15:
        # State only needs 1 page. Delete Page 2 and Page 3 entirely.
        print("     [FORMAT] State only needs 1 page. Eliminating Page 2 and 3.")
        sheet.range('75:300').delete()
        
    elif total_districts <= 28:
        # State needs 2 pages. Delete Page 3 entirely.
        print("     [FORMAT] State needs 2 pages. Eliminating Page 3.")
        sheet.range('141:300').delete()