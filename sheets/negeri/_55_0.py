import pandas as pd
from src.data_provider import get_metrics_dict

# ==========================================
# 1. MAPPING & PAGINATION CONFIGURATION
# ==========================================
COL_DISTRICT = 2   # Column B
# No COL_YEAR / YEAR_OFFSETS - this table only covers a single year (2024), one row per district.

# Data columns: 3 age-group blocks, 4 dose types each
# Umur 5-17 tahun
COL_UMUR_17_DOS1 = 5        # Column E
COL_UMUR_17_DOS2 = 6         # Column F
COL_UMUR_17_BOOSTER1 = 7      # Column G
COL_UMUR_17_BOOSTER2 = 8       # Column H

# Umur 18-59 tahun
COL_UMUR_18_DOS1 = 10       # Column J
COL_UMUR_18_DOS2 = 11        # Column K
COL_UMUR_18_BOOSTER1 = 12     # Column L
COL_UMUR_18_BOOSTER2 = 13      # Column M

# Umur 60 dan lebih
COL_UMUR_60_DOS1 = 15       # Column O
COL_UMUR_60_DOS2 = 16        # Column P
COL_UMUR_60_BOOSTER1 = 17     # Column Q
COL_UMUR_60_BOOSTER2 = 18      # Column R

# The row where the State (Negeri) total is printed.
STATE_START_ROW = 10

# Block start rows for 55.0's layout:
# - Page 1: 20 district rows (rows 12-50, block height 2: 1 data row + 1 blank separator)
# - Page 2: 20 district rows (rows 66-104)
BLOCK_START_ROWS = (
    [12 + (i * 2) for i in range(20)] +   # Page 1 blocks
    [66 + (i * 2) for i in range(20)]     # Page 2 blocks
)

# Title cells: "bm"/"en" hold the description text in column C, one row apart.
TITLE_COORDINATES = [
    {"bm": "C2", "en": "C3", "is_samb": False},        # Page 1 Title
    {"bm": "C58", "en": "C59", "is_samb": True}         # Page 2 Title
]

# ==========================================
# 2. REPORT INJECTION ENGINE
# ==========================================
def populate_jadual_55_0(sheet, hierarchy, report_type):
    state_name = hierarchy.get('state_name', 'Unknown State')
    state_code = hierarchy.get('state_code', '00')
    districts = hierarchy.get('districts', [])

    # 1. Title Generation
    title_bm_base = f": Statistik penerima vaksin COVID-19 mengikut umur, {state_name}, 2024"
    title_en_base = f": Statistics of COVID-19 vaccine recipient by age, {state_name}, 2024"

    for coords in TITLE_COORDINATES:
        suffix_bm = " (samb.)" if coords["is_samb"] else ""
        suffix_en = " (cont'd)" if coords["is_samb"] else ""
        sheet.range(coords["bm"]).value = title_bm_base + suffix_bm
        sheet.range(coords["en"]).value = title_en_base + suffix_en

    # 2. State-Level Data Injection (single row, no year loop)
    sheet.range((STATE_START_ROW, COL_DISTRICT)).value = state_name.upper()
    state_metrics = get_metrics_dict("STATE_TOTAL", level='daerah', parent_code=state_code)
    year_data = state_metrics.get("2024", {})

    val_umur_17_dos1 = year_data.get("umur_17_dos1", "n.a")
    val_umur_17_dos2 = year_data.get("umur_17_dos2", "n.a")
    val_umur_17_booster1 = year_data.get("umur_17_booster1", "n.a")
    val_umur_17_booster2 = year_data.get("umur_17_booster2", "n.a")
    val_umur_18_dos1 = year_data.get("umur_18_dos1", "n.a")
    val_umur_18_dos2 = year_data.get("umur_18_dos2", "n.a")
    val_umur_18_booster1 = year_data.get("umur_18_booster1", "n.a")
    val_umur_18_booster2 = year_data.get("umur_18_booster2", "n.a")
    val_umur_60_dos1 = year_data.get("umur_60_dos1", "n.a")
    val_umur_60_dos2 = year_data.get("umur_60_dos2", "n.a")
    val_umur_60_booster1 = year_data.get("umur_60_booster1", "n.a")
    val_umur_60_booster2 = year_data.get("umur_60_booster2", "n.a")

    sheet.range((STATE_START_ROW, COL_UMUR_17_DOS1)).value = float(val_umur_17_dos1) if pd.notna(val_umur_17_dos1) and val_umur_17_dos1 not in ["n.a", ""] else "n.a"
    sheet.range((STATE_START_ROW, COL_UMUR_17_DOS2)).value = float(val_umur_17_dos2) if pd.notna(val_umur_17_dos2) and val_umur_17_dos2 not in ["n.a", ""] else "n.a"
    sheet.range((STATE_START_ROW, COL_UMUR_17_BOOSTER1)).value = float(val_umur_17_booster1) if pd.notna(val_umur_17_booster1) and val_umur_17_booster1 not in ["n.a", ""] else "n.a"
    sheet.range((STATE_START_ROW, COL_UMUR_17_BOOSTER2)).value = float(val_umur_17_booster2) if pd.notna(val_umur_17_booster2) and val_umur_17_booster2 not in ["n.a", ""] else "n.a"
    sheet.range((STATE_START_ROW, COL_UMUR_18_DOS1)).value = float(val_umur_18_dos1) if pd.notna(val_umur_18_dos1) and val_umur_18_dos1 not in ["n.a", ""] else "n.a"
    sheet.range((STATE_START_ROW, COL_UMUR_18_DOS2)).value = float(val_umur_18_dos2) if pd.notna(val_umur_18_dos2) and val_umur_18_dos2 not in ["n.a", ""] else "n.a"
    sheet.range((STATE_START_ROW, COL_UMUR_18_BOOSTER1)).value = float(val_umur_18_booster1) if pd.notna(val_umur_18_booster1) and val_umur_18_booster1 not in ["n.a", ""] else "n.a"
    sheet.range((STATE_START_ROW, COL_UMUR_18_BOOSTER2)).value = float(val_umur_18_booster2) if pd.notna(val_umur_18_booster2) and val_umur_18_booster2 not in ["n.a", ""] else "n.a"
    sheet.range((STATE_START_ROW, COL_UMUR_60_DOS1)).value = float(val_umur_60_dos1) if pd.notna(val_umur_60_dos1) and val_umur_60_dos1 not in ["n.a", ""] else "n.a"
    sheet.range((STATE_START_ROW, COL_UMUR_60_DOS2)).value = float(val_umur_60_dos2) if pd.notna(val_umur_60_dos2) and val_umur_60_dos2 not in ["n.a", ""] else "n.a"
    sheet.range((STATE_START_ROW, COL_UMUR_60_BOOSTER1)).value = float(val_umur_60_booster1) if pd.notna(val_umur_60_booster1) and val_umur_60_booster1 not in ["n.a", ""] else "n.a"
    sheet.range((STATE_START_ROW, COL_UMUR_60_BOOSTER2)).value = float(val_umur_60_booster2) if pd.notna(val_umur_60_booster2) and val_umur_60_booster2 not in ["n.a", ""] else "n.a"

    # 3. District-Level Data Injection & Cleanup
    total_districts = len(districts)

    for district_idx, base_row in enumerate(BLOCK_START_ROWS):
        if district_idx < total_districts:
            dist_code = districts[district_idx]['code']
            sheet.range((base_row, COL_DISTRICT)).value = districts[district_idx]['name']

            dist_metrics = get_metrics_dict(dist_code, level='daerah', parent_code=state_code)
            year_data = dist_metrics.get("2024", {})

            val_umur_17_dos1 = year_data.get("umur_17_dos1", "n.a")
            val_umur_17_dos2 = year_data.get("umur_17_dos2", "n.a")
            val_umur_17_booster1 = year_data.get("umur_17_booster1", "n.a")
            val_umur_17_booster2 = year_data.get("umur_17_booster2", "n.a")
            val_umur_18_dos1 = year_data.get("umur_18_dos1", "n.a")
            val_umur_18_dos2 = year_data.get("umur_18_dos2", "n.a")
            val_umur_18_booster1 = year_data.get("umur_18_booster1", "n.a")
            val_umur_18_booster2 = year_data.get("umur_18_booster2", "n.a")
            val_umur_60_dos1 = year_data.get("umur_60_dos1", "n.a")
            val_umur_60_dos2 = year_data.get("umur_60_dos2", "n.a")
            val_umur_60_booster1 = year_data.get("umur_60_booster1", "n.a")
            val_umur_60_booster2 = year_data.get("umur_60_booster2", "n.a")

            sheet.range((base_row, COL_UMUR_17_DOS1)).value = float(val_umur_17_dos1) if pd.notna(val_umur_17_dos1) and val_umur_17_dos1 not in ["n.a", ""] else "n.a"
            sheet.range((base_row, COL_UMUR_17_DOS2)).value = float(val_umur_17_dos2) if pd.notna(val_umur_17_dos2) and val_umur_17_dos2 not in ["n.a", ""] else "n.a"
            sheet.range((base_row, COL_UMUR_17_BOOSTER1)).value = float(val_umur_17_booster1) if pd.notna(val_umur_17_booster1) and val_umur_17_booster1 not in ["n.a", ""] else "n.a"
            sheet.range((base_row, COL_UMUR_17_BOOSTER2)).value = float(val_umur_17_booster2) if pd.notna(val_umur_17_booster2) and val_umur_17_booster2 not in ["n.a", ""] else "n.a"
            sheet.range((base_row, COL_UMUR_18_DOS1)).value = float(val_umur_18_dos1) if pd.notna(val_umur_18_dos1) and val_umur_18_dos1 not in ["n.a", ""] else "n.a"
            sheet.range((base_row, COL_UMUR_18_DOS2)).value = float(val_umur_18_dos2) if pd.notna(val_umur_18_dos2) and val_umur_18_dos2 not in ["n.a", ""] else "n.a"
            sheet.range((base_row, COL_UMUR_18_BOOSTER1)).value = float(val_umur_18_booster1) if pd.notna(val_umur_18_booster1) and val_umur_18_booster1 not in ["n.a", ""] else "n.a"
            sheet.range((base_row, COL_UMUR_18_BOOSTER2)).value = float(val_umur_18_booster2) if pd.notna(val_umur_18_booster2) and val_umur_18_booster2 not in ["n.a", ""] else "n.a"
            sheet.range((base_row, COL_UMUR_60_DOS1)).value = float(val_umur_60_dos1) if pd.notna(val_umur_60_dos1) and val_umur_60_dos1 not in ["n.a", ""] else "n.a"
            sheet.range((base_row, COL_UMUR_60_DOS2)).value = float(val_umur_60_dos2) if pd.notna(val_umur_60_dos2) and val_umur_60_dos2 not in ["n.a", ""] else "n.a"
            sheet.range((base_row, COL_UMUR_60_BOOSTER1)).value = float(val_umur_60_booster1) if pd.notna(val_umur_60_booster1) and val_umur_60_booster1 not in ["n.a", ""] else "n.a"
            sheet.range((base_row, COL_UMUR_60_BOOSTER2)).value = float(val_umur_60_booster2) if pd.notna(val_umur_60_booster2) and val_umur_60_booster2 not in ["n.a", ""] else "n.a"
        else:
            # Wipe the dummy row clean if no district exists for this slot (single row only, no year offset)
            sheet.range((base_row, COL_UMUR_17_DOS1), (base_row, COL_UMUR_17_BOOSTER2)).value = None
            sheet.range((base_row, COL_UMUR_18_DOS1), (base_row, COL_UMUR_18_BOOSTER2)).value = None
            sheet.range((base_row, COL_UMUR_60_DOS1), (base_row, COL_UMUR_60_BOOSTER2)).value = None
            sheet.range((base_row, COL_DISTRICT)).value = None

    # 4. Dynamic Page Elimination (Cleanup unused pages)
    # This table only has 2 pages: Page 1 holds 20 districts, Page 2 holds 20 more (40 total)
    if total_districts <= 20:
        # State only needs 1 page. Delete Page 2 entirely (page 2 title starts row 58).
        print("     [FORMAT] State only needs 1 page. Eliminating Page 2.")
        sheet.range('57:300').delete()