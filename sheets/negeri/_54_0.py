import pandas as pd
from src.data_provider import get_metrics_dict
from src.excel_utils import safe_write

# ==========================================
# 1. MAPPING & PAGINATION CONFIGURATION
# ==========================================
COL_DISTRICT = 2   # Column B
# No COL_YEAR / YEAR_OFFSETS - this table only covers a single year (2024), one row per district.

# Data columns (alternating D, F, H, J, L, N, P, R - odd columns between are unused/merged)
COL_DURIAN = 4        
COL_MANGGA = 6        
COL_NANAS = 8         
COL_CILI = 10         
COL_KOBIS_BULAT = 12  
COL_TIMUN = 14        
COL_TOMATO = 16       
COL_SAWI = 18         

STATE_START_ROW = 13

BLOCK_START_ROWS = (
    [15 + (i * 2) for i in range(20)] +   # Page 1 blocks
    [82 + (i * 2) for i in range(20)]     # Page 2 blocks
)

# Converted to (Row, Col) tuples for openpyxl
TITLE_COORDINATES = [
    {"bm": (3, 3), "en": (4, 3), "is_samb": False},      
    {"bm": (72, 3), "en": (73, 3), "is_samb": True}       
]

# ==========================================
# SANITIZER HELPER
# ==========================================
def sanitize_value(val):
    if pd.notna(val):
        clean_val = str(val).strip()
        if clean_val not in ["", "n.a", "n.a.", "-", "na"]:
            try: return float(clean_val)
            except (ValueError, TypeError): pass
    return "n.a"

# ==========================================
# 2. REPORT INJECTION ENGINE
# ==========================================
def populate_jadual_54_0(sheet, hierarchy, report_type):
    state_name = hierarchy.get('state_name', 'Unknown State')
    state_code = hierarchy.get('state_code', '00')
    districts = hierarchy.get('districts', [])
    print(f"  -> Populating Jadual 54.0 (Penggunaan Per Kapita) untuk {state_name}")

    # 1. Title Generation
    title_bm_base = f": Statistik terpilih Penggunaan Per Kapita item pertanian mengikut daerah pentadbiran, {state_name}, 2024"
    title_en_base = f": Selected statistics of Per Capita Consumption of agricultural item by administrative district, {state_name}, 2024"

    for coords in TITLE_COORDINATES:
        suffix_bm = " (samb.)" if coords["is_samb"] else ""
        suffix_en = " (cont'd)" if coords["is_samb"] else ""
        safe_write(sheet, coords["bm"][0], coords["bm"][1], title_bm_base + suffix_bm)
        safe_write(sheet, coords["en"][0], coords["en"][1], title_en_base + suffix_en)

    # 2. State-Level Data Injection (single row, no year loop)
    safe_write(sheet, STATE_START_ROW, COL_DISTRICT, state_name.upper())
    state_metrics = get_metrics_dict("STATE_TOTAL", level='daerah', parent_code=state_code)
    year_data = state_metrics.get("2024", {})

    safe_write(sheet, STATE_START_ROW, COL_DURIAN, sanitize_value(year_data.get("durian", "n.a")))
    safe_write(sheet, STATE_START_ROW, COL_MANGGA, sanitize_value(year_data.get("mangga", "n.a")))
    safe_write(sheet, STATE_START_ROW, COL_NANAS, sanitize_value(year_data.get("nanas", "n.a")))
    safe_write(sheet, STATE_START_ROW, COL_CILI, sanitize_value(year_data.get("cili", "n.a")))
    safe_write(sheet, STATE_START_ROW, COL_KOBIS_BULAT, sanitize_value(year_data.get("kobis_bulat", "n.a")))
    safe_write(sheet, STATE_START_ROW, COL_TIMUN, sanitize_value(year_data.get("timun", "n.a")))
    safe_write(sheet, STATE_START_ROW, COL_TOMATO, sanitize_value(year_data.get("tomato", "n.a")))
    safe_write(sheet, STATE_START_ROW, COL_SAWI, sanitize_value(year_data.get("sawi", "n.a")))

    # 3. District-Level Data Injection & Cleanup
    total_districts = len(districts)

    for district_idx, base_row in enumerate(BLOCK_START_ROWS):
        if district_idx < total_districts:
            dist_code = districts[district_idx]['code']
            safe_write(sheet, base_row, COL_DISTRICT, districts[district_idx]['name'])

            dist_metrics = get_metrics_dict(dist_code, level='daerah', parent_code=state_code)
            year_data = dist_metrics.get("2024", {})

            safe_write(sheet, base_row, COL_DURIAN, sanitize_value(year_data.get("durian", "n.a")))
            safe_write(sheet, base_row, COL_MANGGA, sanitize_value(year_data.get("mangga", "n.a")))
            safe_write(sheet, base_row, COL_NANAS, sanitize_value(year_data.get("nanas", "n.a")))
            safe_write(sheet, base_row, COL_CILI, sanitize_value(year_data.get("cili", "n.a")))
            safe_write(sheet, base_row, COL_KOBIS_BULAT, sanitize_value(year_data.get("kobis_bulat", "n.a")))
            safe_write(sheet, base_row, COL_TIMUN, sanitize_value(year_data.get("timun", "n.a")))
            safe_write(sheet, base_row, COL_TOMATO, sanitize_value(year_data.get("tomato", "n.a")))
            safe_write(sheet, base_row, COL_SAWI, sanitize_value(year_data.get("sawi", "n.a")))
        else:
            # XML-Level Cleanup
            safe_write(sheet, base_row, COL_DISTRICT, None)
            for c in [COL_DURIAN, COL_MANGGA, COL_NANAS, COL_CILI, COL_KOBIS_BULAT, COL_TIMUN, COL_TOMATO, COL_SAWI]:
                safe_write(sheet, base_row, c, None)

    # 4. XML-Level Page Deletion
    if total_districts <= 20:
        print("     [FORMAT] State only needs 1 page. Eliminating Page 2.")
        sheet.delete_rows(71, 230)