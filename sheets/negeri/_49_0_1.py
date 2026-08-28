import pandas as pd
from src.data_provider import get_metrics_dict, get_jkm_hierarchy
from src.excel_utils import safe_write

# ==========================================
# 1. MAPPING & PAGINATION CONFIGURATION
# ==========================================
COL_DISTRICT = 2  # Column B

# Map the EXCEL COLUMN NUMBER to the YEAR STRING.
COL_YEAR_MAP = {
    5: "2022", # Column E
    6: "2023", # Column F
    7: "2024"  # Column G
}

STATE_START_ROW = 11

# Recalculated block start rows based on the template's layout.
BLOCK_START_ROWS = (
    [13 + (i * 2) for i in range(20)] + # Page 1: 20 slots starting at Row 13
    [68 + (i * 2) for i in range(20)]   # Page 2: 20 slots starting at Row 68
)

# Coordinates converted to tuples (Row, Col) for openpyxl safe_write
TITLE_COORDINATES = [
    {"bm": (2, 3), "en": (3, 3), "is_samb": True},     # Page 1 Titles
    {"bm": (63, 3), "en": (64, 3), "is_samb": True}    # Page 2 Titles
]

# ==========================================
# SANITIZER HELPER
# ==========================================
def sanitize_value(val):
    """
    Safely attempts to convert a value to float. 
    Catches common Excel accounting formats like dashes ('-') and weird spaces,
    forcing them into a standard 'n.a' string to prevent critical engine crashes.
    """
    if pd.notna(val):
        clean_val = str(val).strip()
        if clean_val not in ["", "n.a", "n.a.", "-", "na"]:
            try:
                return float(clean_val)
            except (ValueError, TypeError):
                pass
    return "n.a"

# ==========================================
# 2. REPORT INJECTION ENGINE
# ==========================================
def populate_jadual_49_0_1(sheet, hierarchy, report_type):
    state_name = hierarchy.get('state_name', 'Unknown State')
    state_code = hierarchy.get('state_code', '00')
    
    jkm_branches = get_jkm_hierarchy(state_code)
    print(f"  -> Populating Jadual 49.0 (Cawangan JKM - Bilangan Kanak-Kanak)(samb.) untuk {state_name}")

    # 1. Title Generation
    title_bm_base = f": Statistik taman asuhan kanak-kanak mengikut cawangan JKM, {state_name}, 2022 - 2024"
    title_en_base = f": Statistics of kindergarten by JKM branch, {state_name}, 2022 - 2024"
    
    for coords in TITLE_COORDINATES:
        suffix_bm = " (samb.)" if coords["is_samb"] else ""
        suffix_en = " (cont'd)" if coords["is_samb"] else ""
        safe_write(sheet, coords["bm"][0], coords["bm"][1], title_bm_base + suffix_bm)
        safe_write(sheet, coords["en"][0], coords["en"][1], title_en_base + suffix_en)

    # 2. State-Level Data Injection (Row 11)
    safe_write(sheet, STATE_START_ROW, COL_DISTRICT, state_name.upper())
    
    # Fetch State Total directly from the general Negeri fact table
    state_metrics = get_metrics_dict(state_code, level='negeri')
    
    for col_idx, year in COL_YEAR_MAP.items():
        year_data = state_metrics.get(year, {})
        raw_val = year_data.get("bilangan_kanak_kanak", "n.a")
        
        # Apply the robust sanitizer before writing to the sheet
        safe_write(sheet, STATE_START_ROW, col_idx, sanitize_value(raw_val))

    # 3. Branch-Level Data Injection & Cleanup 
    total_branches = len(jkm_branches)
    
    for branch_idx, target_row in enumerate(BLOCK_START_ROWS):
        if branch_idx < total_branches:
            # --- Inject Branch Data ---
            branch_code = jkm_branches[branch_idx]['code']
            safe_write(sheet, target_row, COL_DISTRICT, jkm_branches[branch_idx]['name'])
            
            # Fetch using isolated jkm dimension
            branch_metrics = get_metrics_dict(branch_code, level='jkm', parent_code=state_code)
            
            for col_idx, year in COL_YEAR_MAP.items():
                year_data = branch_metrics.get(year, {})
                raw_val = year_data.get("bilangan_kanak_kanak", "n.a")
                
                safe_write(sheet, target_row, col_idx, sanitize_value(raw_val))
        else:
            # --- Openpyxl Safe XML Wiping ---
            safe_write(sheet, target_row, COL_DISTRICT, None)
            for c in range(min(COL_YEAR_MAP.keys()), max(COL_YEAR_MAP.keys()) + 1):
                safe_write(sheet, target_row, c, None)

    # 4. XML-Level Page Deletion
    # Syntax: sheet.delete_rows(start_row_index, number_of_rows_to_delete)
    if total_branches <= 20:
        print("     [FORMAT] State only needs 1 page. Eliminating Page 2.")
        sheet.delete_rows(62, 238) # Deletes ~238 rows starting at row 62 to cleanly wipe the footer and Page 2