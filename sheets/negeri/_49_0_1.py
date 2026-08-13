import pandas as pd
from src.data_provider import get_metrics_dict, get_jkm_hierarchy

# ==========================================
# 1. MAPPING & PAGINATION CONFIGURATION
# ==========================================
# TODO: Update this column index to match where the District/Branch names are listed.
COL_DISTRICT = 2  # Column B

# TODO: Map the EXCEL COLUMN NUMBER to the YEAR STRING.
# Adjust the indices (5, 6, 7) if the years are located in different columns in a new template.
COL_YEAR_MAP = {
    5: "2022", # Column E
    6: "2023", # Column F
    7: "2024"  # Column G
}

# TODO: Define the exact row where the State (Negeri) total is printed.
STATE_START_ROW = 11

# TODO: Recalculate these block start rows based on the new template's layout.
# - The first number (e.g., 13) is the starting row for that specific page.
# - The multiplier (e.g., 2) is the row spacing/step (1 row for data + 1 blank separator row).
# - The range (e.g., 20) is the maximum number of branches that can fit on that page.
BLOCK_START_ROWS = (
    [13 + (i * 2) for i in range(20)] + # Page 1: 20 slots starting at Row 13
    [69 + (i * 2) for i in range(20)]   # Page 2: 20 slots starting at Row 69
)

# TODO: Update these cell coordinates if the bilingual titles are placed differently.
TITLE_COORDINATES = [
    {"bm": "C2", "en": "C3", "is_samb": False},     # Page 1 Titles
    {"bm": "C63", "en": "C64", "is_samb": True}     # Page 2 Titles
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
        # Check if it's not in our list of known "empty" string formats
        if clean_val not in ["", "n.a", "n.a.", "-", "na"]:
            try:
                return float(clean_val)
            except (ValueError, TypeError):
                pass # If it still fails to convert, fall through to returning "n.a"
    return "n.a"

# ==========================================
# 2. REPORT INJECTION ENGINE
# ==========================================
# TODO: Rename this function to match your new Jadual (e.g., populate_jadual_50)
def populate_jadual_49(sheet, hierarchy, report_type):
    state_name = hierarchy.get('state_name', 'Unknown State')
    state_code = hierarchy.get('state_code', '00')
    
    # TODO: If reusing this pattern for a standard District (Daerah) table, 
    # replace this line with: districts = hierarchy.get('districts', [])
    jkm_branches = get_jkm_hierarchy(state_code)

    # 1. Title Generation
    # TODO: Update the base title text to reflect the metrics of the new Jadual
    title_bm_base = f": Statistik taman asuhan kanak-kanak mengikut cawangan JKM, {state_name}, 2022 - 2024"
    title_en_base = f": Statistics of kindergarten by JKM branch, {state_name}, 2022 - 2024"
    
    for coords in TITLE_COORDINATES:
        suffix_bm = " (samb.)" if coords["is_samb"] else ""
        suffix_en = " (cont'd)" if coords["is_samb"] else ""
        sheet.range(coords["bm"]).value = title_bm_base + suffix_bm
        sheet.range(coords["en"]).value = title_en_base + suffix_en

    # 2. State-Level Data Injection (Row 11)
    sheet.range((STATE_START_ROW, COL_DISTRICT)).value = state_name.upper()
    
    # Fetch State Total directly from the general Negeri fact table
    state_metrics = get_metrics_dict(state_code, level='negeri')
    
    for col_idx, year in COL_YEAR_MAP.items():
        year_data = state_metrics.get(year, {})
        
        # TODO: Change "bilangan_taman_asuhan" to the exact metric name used in fact_metrics_negeri
        raw_val = year_data.get("bilangan_kanak_kanak", "n.a")
        
        # Apply the robust sanitizer before writing to the sheet
        sheet.range((STATE_START_ROW, col_idx)).value = sanitize_value(raw_val)

    # 3. Branch-Level Data Injection & Cleanup 
    total_branches = len(jkm_branches)
    
    for branch_idx, target_row in enumerate(BLOCK_START_ROWS):
        if branch_idx < total_branches:
            # --- Inject Branch Data ---
            branch_code = jkm_branches[branch_idx]['code']
            sheet.range((target_row, COL_DISTRICT)).value = jkm_branches[branch_idx]['name']
            
            # TODO: If reusing for standard districts, change level='jkm' to level='daerah'
            # Note: parent_code=state_code acts as a safety lock to prevent inter-state collisions
            branch_metrics = get_metrics_dict(branch_code, level='jkm', parent_code=state_code)
            
            for col_idx, year in COL_YEAR_MAP.items():
                year_data = branch_metrics.get(year, {})
                
                # TODO: Change "bilangan_taman_asuhan" to the exact metric name used in fact_metrics_cawangan_jkm
                raw_val = year_data.get("bilangan_kanak_kanak", "n.a")
                
                sheet.range((target_row, col_idx)).value = sanitize_value(raw_val)
        else:
            # --- Cleanup Unused Row ---
            # Wipes the Branch Name and the year columns to leave a pristine blank row
            sheet.range((target_row, COL_DISTRICT)).value = None
            sheet.range((target_row, min(COL_YEAR_MAP.keys())), (target_row, max(COL_YEAR_MAP.keys()))).value = None

    # 4. Dynamic Page Elimination (Cleanup unused pages)
    # TODO: Adjust the `total_branches` threshold limit based on how many slots fit on Page 1.
    if total_branches <= 20:
        print("     [FORMAT] State only needs 1 page. Eliminating Page 2.")
        # TODO: Adjust the `.delete()` row range (e.g., '58:300') to precisely target 
        # where Page 1's footer ends and Page 2 begins in the new template.
        sheet.range('58:300').delete()