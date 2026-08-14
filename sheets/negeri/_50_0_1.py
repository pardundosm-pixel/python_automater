import pandas as pd
from src.data_provider import get_metrics_dict, get_pdrm_hierarchy

# ==========================================
# 1. MAPPING & PAGINATION CONFIGURATION
# ==========================================
# Column B holds the PDRM District Names
COL_DISTRICT = 2  

# Map the EXCEL COLUMN NUMBERS to the exact Metric Names in your database
# TODO: Update the dictionary values to match the metric names in fact_metrics_daerah_pdrm
COL_METRICS_MAP = {
    5: "jumlah_jenayah_harta_benda",  
    6: "pecah_rumah_curi",           
    7: "kecurian_lori",           
    8: "kecurian_kereta",           
    9: "kecurian_moto",
    10: "kecurian_lain" 
}

# The dictionary key (0, 1, 2) represents how many rows down from the block's start row the year sits.
YEAR_OFFSETS = {
    0: "2023",
    1: "2024",
    2: "2025"
}

# The exact row where the State (Negeri) total block begins
STATE_START_ROW = 9

# Recalculated block start rows based on the Jadual 50 screenshots:
# - Page 1 holds 13 districts. Starts at Row 14, spaced every 4 rows (14, 18, 22...).
# - Page 2 starts at Row 81, spaced every 4 rows.
# - Page 3 (estimated) starts around Row 148, spaced every 4 rows.
BLOCK_START_ROWS = (
    [13 + (i * 4) for i in range(13)] +  
    [77 + (i * 4) for i in range(15)] +  
    [141 + (i * 4) for i in range(15)]   
)

# Coordinates for the bilingual titles
TITLE_COORDINATES = [
    {"bm": "C3", "en": "C4", "is_samb": False},     # Page 1 Titles
    {"bm": "C74", "en": "C75", "is_samb": True}     # Page 2 Titles
]

# ==========================================
# SANITIZER HELPER
# ==========================================
def sanitize_value(val):
    """Safely attempts to convert a value to float, catching dashes and spaces."""
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
# TODO: Rename this function to match your engine.py registry (e.g., populate_jadual_50_0)
def populate_jadual_50_0_1(sheet, hierarchy, report_type):
    state_name = hierarchy.get('state_name', 'Unknown State')
    state_code = hierarchy.get('state_code', '00')
    print(f"  -> Populating Jadual 50.1 (Jenayah PDRM) untuk {state_name}")
    
    # Fetch PDRM districts directly using the isolated domain dimension
    pdrm_districts = get_pdrm_hierarchy(state_code)

    # 1. Title Generation
    title_bm_base = f": Jenayah harta benda mengikut daerah PDRM dan jenis jenayah, {state_name}, 2023 - 2025"
    title_en_base = f": Property crime by PDRM district and type of crime, {state_name}, 2023 - 2025"
    
    for coords in TITLE_COORDINATES:
        suffix_bm = " (samb.)" if coords["is_samb"] else ""
        suffix_en = " (cont'd)" if coords["is_samb"] else ""
        sheet.range(coords["bm"]).value = title_bm_base + suffix_bm
        sheet.range(coords["en"]).value = title_en_base + suffix_en

    # 2. State-Level Data Injection (Row 10 - 12)
    sheet.range((STATE_START_ROW, COL_DISTRICT)).value = state_name.upper()
    
    # Fetch State Total directly from the general Negeri fact table
    # TODO: If your state total is inside fact_pdrm instead, you will need to add the 
    # "STATE_TOTAL" logic to the 'pdrm' block in data_provider.py and switch this to level='pdrm'.
    state_metrics = get_metrics_dict(state_code, level='negeri')
    
    for offset, year in YEAR_OFFSETS.items():
        target_row = STATE_START_ROW + offset
        year_data = state_metrics.get(year, {})
        
        for col_idx, metric_name in COL_METRICS_MAP.items():
            raw_val = year_data.get(metric_name, "n.a")
            sheet.range((target_row, col_idx)).value = sanitize_value(raw_val)

    # 3. District-Level Data Injection & Cleanup 
    total_districts = len(pdrm_districts)
    
    for district_idx, base_row in enumerate(BLOCK_START_ROWS):
        if district_idx < total_districts:
            # --- Inject Branch Data ---
            dist_code = pdrm_districts[district_idx]['code']
            sheet.range((base_row, COL_DISTRICT)).value = pdrm_districts[district_idx]['name']
            
            # Fetch using the new 'pdrm' level routing with collision prevention
            branch_metrics = get_metrics_dict(dist_code, level='pdrm', parent_code=state_code)
            
            for offset, year in YEAR_OFFSETS.items():
                target_row = base_row + offset
                year_data = branch_metrics.get(year, {})
                
                for col_idx, metric_name in COL_METRICS_MAP.items():
                    raw_val = year_data.get(metric_name, "n.a")
                    sheet.range((target_row, col_idx)).value = sanitize_value(raw_val)
        else:
            # --- Cleanup Unused Row ---
            # Wipes the District Name and all metric columns to leave a pristine blank block
            sheet.range((base_row, COL_DISTRICT)).value = None
            for offset in YEAR_OFFSETS.keys():
                sheet.range((base_row + offset, min(COL_METRICS_MAP.keys())), 
                            (base_row + offset, max(COL_METRICS_MAP.keys()))).value = None

    # 4. Dynamic Page Elimination (Cleanup unused pages)
    if total_districts <= 13:
        print("     [FORMAT] State only needs 1 page. Eliminating Page 2.")
        # Row 67 is the footer for Page 1. Page 2 begins around 74.
        sheet.range('71:300').delete()
        
    elif total_districts <= 28:
        print("     [FORMAT] State needs 2 pages. Eliminating Page 3.")
        sheet.range('138:300').delete()