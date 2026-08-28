import pandas as pd
from src.data_provider import get_metrics_dict
from src.excel_utils import safe_write

# ==========================================
# 1. MAPPING & PAGINATION CONFIGURATION
# ==========================================
COL_DISTRICT = 2   
COL_PENDAPATAN_PURATA = 5     
COL_PENDAPATAN_PENENGAH = 6   
COL_PERBELANJAAN = 7           
COL_KEMISKINAN = 8              
COL_GINI = 9                     

YEAR_OFFSETS = {0: "2022", 1: "2024"}
STATE_START_ROW = 9

BLOCK_START_ROWS = (
    [12 + (i * 3) for i in range(20)] +   
    [86 + (i * 3) for i in range(20)] +   
    [158 + (i * 3) for i in range(20)]    
)

TITLE_COORDINATES = [
    {"bm": (1, 3), "en": (2, 3), "is_samb": False},       
    {"bm": (78, 3), "en": (79, 3), "is_samb": True},      
    {"bm": (150, 3), "en": (151, 3), "is_samb": True}      
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
def populate_jadual_53_0(sheet, hierarchy, report_type):
    state_name = hierarchy.get('state_name', 'Unknown State')
    state_code = hierarchy.get('state_code', '00')
    districts = hierarchy.get('districts', [])
    print(f"  -> Populating Jadual 53.0 (Pendapatan & Kemiskinan) untuk {state_name}")

    # 1. Title Generation
    title_bm_base = f": Pendapatan, perbelanjaan dan kemiskinan, {state_name}, 2022 dan 2024"
    title_en_base = f": Income, expenditure and poverty, {state_name}, 2022 and 2024"

    for coords in TITLE_COORDINATES:
        suffix_bm = " (samb.)" if coords["is_samb"] else ""
        suffix_en = " (cont'd)" if coords["is_samb"] else ""
        safe_write(sheet, coords["bm"][0], coords["bm"][1], title_bm_base + suffix_bm)
        safe_write(sheet, coords["en"][0], coords["en"][1], title_en_base + suffix_en)

    # 2. State-Level Data Injection
    safe_write(sheet, STATE_START_ROW, COL_DISTRICT, state_name.upper())
    state_metrics = get_metrics_dict(state_code, level='negeri')

    for offset, year in YEAR_OFFSETS.items():
        target_row = STATE_START_ROW + offset
        year_data = state_metrics.get(year, {})

        safe_write(sheet, target_row, COL_PENDAPATAN_PURATA, sanitize_value(year_data.get("pendapatan_purata", "n.a")))
        safe_write(sheet, target_row, COL_PENDAPATAN_PENENGAH, sanitize_value(year_data.get("pendapatan_penengah", "n.a")))
        safe_write(sheet, target_row, COL_PERBELANJAAN, sanitize_value(year_data.get("perbelanjaan", "n.a")))
        safe_write(sheet, target_row, COL_KEMISKINAN, sanitize_value(year_data.get("kemiskinan", "n.a")))
        safe_write(sheet, target_row, COL_GINI, sanitize_value(year_data.get("gini", "n.a")))

    # 3. District-Level Data Injection & Cleanup
    total_districts = len(districts)

    for district_idx, base_row in enumerate(BLOCK_START_ROWS):
        if district_idx < total_districts:
            dist_code = districts[district_idx]['code']
            safe_write(sheet, base_row, COL_DISTRICT, districts[district_idx]['name'])
            dist_metrics = get_metrics_dict(dist_code, level='daerah', parent_code=state_code)

            for offset, year in YEAR_OFFSETS.items():
                target_row = base_row + offset
                year_data = dist_metrics.get(year, {})

                safe_write(sheet, target_row, COL_PENDAPATAN_PURATA, sanitize_value(year_data.get("pendapatan_purata", "n.a")))
                safe_write(sheet, target_row, COL_PENDAPATAN_PENENGAH, sanitize_value(year_data.get("pendapatan_penengah", "n.a")))
                safe_write(sheet, target_row, COL_PERBELANJAAN, sanitize_value(year_data.get("perbelanjaan", "n.a")))
                safe_write(sheet, target_row, COL_KEMISKINAN, sanitize_value(year_data.get("kemiskinan", "n.a")))
                safe_write(sheet, target_row, COL_GINI, sanitize_value(year_data.get("gini", "n.a")))
        else:
            # XML-Level Cleanup
            safe_write(sheet, base_row, COL_DISTRICT, None)
            for r in range(base_row, base_row + len(YEAR_OFFSETS)):
                for c in range(COL_PENDAPATAN_PURATA, COL_GINI + 1):
                    safe_write(sheet, r, c, None)

    # 4. XML-Level Page Deletion
    if total_districts <= 20:
        print("     [FORMAT] State only needs 1 page. Eliminating Page 2 and 3.")
        sheet.delete_rows(77, 224) 
    elif total_districts <= 40:
        print("     [FORMAT] State needs 2 pages. Eliminating Page 3.")
        sheet.delete_rows(149, 152)