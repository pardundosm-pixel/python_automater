import pandas as pd
from src.data_provider import get_metrics_dict
from src.excel_utils import safe_write

# ==========================================
# 1. MAPPING & PAGINATION CONFIGURATION
# ==========================================
COL_DISTRICT = 2   
# No COL_YEAR / YEAR_OFFSETS - this table only covers a single year (2024), one row per district.

# Umur 5-17 tahun
COL_UMUR_17_DOS1 = 5        
COL_UMUR_17_DOS2 = 6         
COL_UMUR_17_BOOSTER1 = 7      
COL_UMUR_17_BOOSTER2 = 8       

# Umur 18-59 tahun
COL_UMUR_18_DOS1 = 10       
COL_UMUR_18_DOS2 = 11        
COL_UMUR_18_BOOSTER1 = 12     
COL_UMUR_18_BOOSTER2 = 13      

# Umur 60 dan lebih
COL_UMUR_60_DOS1 = 15       
COL_UMUR_60_DOS2 = 16        
COL_UMUR_60_BOOSTER1 = 17     
COL_UMUR_60_BOOSTER2 = 18      

STATE_START_ROW = 10

BLOCK_START_ROWS = (
    [12 + (i * 2) for i in range(20)] +   
    [66 + (i * 2) for i in range(20)]     
)

TITLE_COORDINATES = [
    {"bm": (2, 3), "en": (3, 3), "is_samb": False},       
    {"bm": (58, 3), "en": (59, 3), "is_samb": True}        
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
def populate_jadual_55_0(sheet, hierarchy, report_type):
    state_name = hierarchy.get('state_name', 'Unknown State')
    state_code = hierarchy.get('state_code', '00')
    districts = hierarchy.get('districts', [])
    print(f"  -> Populating Jadual 55.0 (Vaksin COVID-19) untuk {state_name}")

    # 1. Title Generation
    title_bm_base = f": Statistik penerima vaksin COVID-19 mengikut umur, {state_name}, 2024"
    title_en_base = f": Statistics of COVID-19 vaccine recipient by age, {state_name}, 2024"

    for coords in TITLE_COORDINATES:
        suffix_bm = " (samb.)" if coords["is_samb"] else ""
        suffix_en = " (cont'd)" if coords["is_samb"] else ""
        safe_write(sheet, coords["bm"][0], coords["bm"][1], title_bm_base + suffix_bm)
        safe_write(sheet, coords["en"][0], coords["en"][1], title_en_base + suffix_en)

    # 2. State-Level Data Injection
    safe_write(sheet, STATE_START_ROW, COL_DISTRICT, state_name.upper())
    state_metrics = get_metrics_dict("STATE_TOTAL", level='daerah', parent_code=state_code)
    year_data = state_metrics.get("2024", {})

    safe_write(sheet, STATE_START_ROW, COL_UMUR_17_DOS1, sanitize_value(year_data.get("umur_17_dos1", "n.a")))
    safe_write(sheet, STATE_START_ROW, COL_UMUR_17_DOS2, sanitize_value(year_data.get("umur_17_dos2", "n.a")))
    safe_write(sheet, STATE_START_ROW, COL_UMUR_17_BOOSTER1, sanitize_value(year_data.get("umur_17_booster1", "n.a")))
    safe_write(sheet, STATE_START_ROW, COL_UMUR_17_BOOSTER2, sanitize_value(year_data.get("umur_17_booster2", "n.a")))
    safe_write(sheet, STATE_START_ROW, COL_UMUR_18_DOS1, sanitize_value(year_data.get("umur_18_dos1", "n.a")))
    safe_write(sheet, STATE_START_ROW, COL_UMUR_18_DOS2, sanitize_value(year_data.get("umur_18_dos2", "n.a")))
    safe_write(sheet, STATE_START_ROW, COL_UMUR_18_BOOSTER1, sanitize_value(year_data.get("umur_18_booster1", "n.a")))
    safe_write(sheet, STATE_START_ROW, COL_UMUR_18_BOOSTER2, sanitize_value(year_data.get("umur_18_booster2", "n.a")))
    safe_write(sheet, STATE_START_ROW, COL_UMUR_60_DOS1, sanitize_value(year_data.get("umur_60_dos1", "n.a")))
    safe_write(sheet, STATE_START_ROW, COL_UMUR_60_DOS2, sanitize_value(year_data.get("umur_60_dos2", "n.a")))
    safe_write(sheet, STATE_START_ROW, COL_UMUR_60_BOOSTER1, sanitize_value(year_data.get("umur_60_booster1", "n.a")))
    safe_write(sheet, STATE_START_ROW, COL_UMUR_60_BOOSTER2, sanitize_value(year_data.get("umur_60_booster2", "n.a")))

    # 3. District-Level Data Injection & Cleanup
    total_districts = len(districts)
    data_cols = [
        COL_UMUR_17_DOS1, COL_UMUR_17_DOS2, COL_UMUR_17_BOOSTER1, COL_UMUR_17_BOOSTER2,
        COL_UMUR_18_DOS1, COL_UMUR_18_DOS2, COL_UMUR_18_BOOSTER1, COL_UMUR_18_BOOSTER2,
        COL_UMUR_60_DOS1, COL_UMUR_60_DOS2, COL_UMUR_60_BOOSTER1, COL_UMUR_60_BOOSTER2
    ]

    for district_idx, base_row in enumerate(BLOCK_START_ROWS):
        if district_idx < total_districts:
            dist_code = districts[district_idx]['code']
            safe_write(sheet, base_row, COL_DISTRICT, districts[district_idx]['name'])

            dist_metrics = get_metrics_dict(dist_code, level='daerah', parent_code=state_code)
            year_data = dist_metrics.get("2024", {})

            safe_write(sheet, base_row, COL_UMUR_17_DOS1, sanitize_value(year_data.get("umur_17_dos1", "n.a")))
            safe_write(sheet, base_row, COL_UMUR_17_DOS2, sanitize_value(year_data.get("umur_17_dos2", "n.a")))
            safe_write(sheet, base_row, COL_UMUR_17_BOOSTER1, sanitize_value(year_data.get("umur_17_booster1", "n.a")))
            safe_write(sheet, base_row, COL_UMUR_17_BOOSTER2, sanitize_value(year_data.get("umur_17_booster2", "n.a")))
            safe_write(sheet, base_row, COL_UMUR_18_DOS1, sanitize_value(year_data.get("umur_18_dos1", "n.a")))
            safe_write(sheet, base_row, COL_UMUR_18_DOS2, sanitize_value(year_data.get("umur_18_dos2", "n.a")))
            safe_write(sheet, base_row, COL_UMUR_18_BOOSTER1, sanitize_value(year_data.get("umur_18_booster1", "n.a")))
            safe_write(sheet, base_row, COL_UMUR_18_BOOSTER2, sanitize_value(year_data.get("umur_18_booster2", "n.a")))
            safe_write(sheet, base_row, COL_UMUR_60_DOS1, sanitize_value(year_data.get("umur_60_dos1", "n.a")))
            safe_write(sheet, base_row, COL_UMUR_60_DOS2, sanitize_value(year_data.get("umur_60_dos2", "n.a")))
            safe_write(sheet, base_row, COL_UMUR_60_BOOSTER1, sanitize_value(year_data.get("umur_60_booster1", "n.a")))
            safe_write(sheet, base_row, COL_UMUR_60_BOOSTER2, sanitize_value(year_data.get("umur_60_booster2", "n.a")))
        else:
            # XML-Level Cleanup
            safe_write(sheet, base_row, COL_DISTRICT, None)
            for c in data_cols:
                safe_write(sheet, base_row, c, None)

    # 4. XML-Level Page Deletion
    if total_districts <= 20:
        print("     [FORMAT] State only needs 1 page. Eliminating Page 2.")
        sheet.delete_rows(57, 244)