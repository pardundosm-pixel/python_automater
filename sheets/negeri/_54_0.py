import pandas as pd
from src.data_provider import get_metrics_dict

# ==========================================
# 1. MAPPING & PAGINATION CONFIGURATION
# ==========================================
COL_DISTRICT = 2   # Column B
# No COL_YEAR / YEAR_OFFSETS - this table only covers a single year (2024), one row per district.

# Data columns (alternating D, F, H, J, L, N, P, R - odd columns between are unused/merged)
COL_DURIAN = 4        # Column D
COL_MANGGA = 6        # Column F
COL_NANAS = 8         # Column H
COL_CILI = 10         # Column J
COL_KOBIS_BULAT = 12  # Column L
COL_TIMUN = 14        # Column N
COL_TOMATO = 16       # Column P
COL_SAWI = 18         # Column R

# The row where the State (Negeri) total is printed.
STATE_START_ROW = 13

# Block start rows for 54.0's layout:
# - Page 1: 20 district rows (rows 15-53, block height 2: 1 data row + 1 blank separator)
# - Page 2: 20 district rows (rows 82-120)
BLOCK_START_ROWS = (
    [15 + (i * 2) for i in range(20)] +   # Page 1 blocks
    [82 + (i * 2) for i in range(20)]     # Page 2 blocks
)

# Title cells: "bm"/"en" hold the description text in column C, one row apart.
TITLE_COORDINATES = [
    {"bm": "C3", "en": "C4", "is_samb": False},        # Page 1 Title
    {"bm": "C72", "en": "C73", "is_samb": True}         # Page 2 Title
]

# ==========================================
# 2. REPORT INJECTION ENGINE
# ==========================================
def populate_jadual_54_0(sheet, hierarchy, report_type):
    state_name = hierarchy.get('state_name', 'Unknown State')
    state_code = hierarchy.get('state_code', '00')
    districts = hierarchy.get('districts', [])

    # 1. Title Generation
    title_bm_base = f": Statistik terpilih Penggunaan Per Kapita item pertanian mengikut daerah pentadbiran, {state_name}, 2024"
    title_en_base = f": Selected statistics of Per Capita Consumption of agricultural item by administrative district, {state_name}, 2024"

    for coords in TITLE_COORDINATES:
        suffix_bm = " (samb.)" if coords["is_samb"] else ""
        suffix_en = " (cont'd)" if coords["is_samb"] else ""
        sheet.range(coords["bm"]).value = title_bm_base + suffix_bm
        sheet.range(coords["en"]).value = title_en_base + suffix_en

    # 2. State-Level Data Injection (single row, no year loop)
    sheet.range((STATE_START_ROW, COL_DISTRICT)).value = state_name.upper()
    state_metrics = get_metrics_dict("STATE_TOTAL", level='daerah', parent_code=state_code)
    year_data = state_metrics.get("2024", {})

    val_durian = year_data.get("durian", "n.a")
    val_mangga = year_data.get("mangga", "n.a")
    val_nanas = year_data.get("nanas", "n.a")
    val_cili = year_data.get("cili", "n.a")
    val_kobis_bulat = year_data.get("kobis_bulat", "n.a")
    val_timun = year_data.get("timun", "n.a")
    val_tomato = year_data.get("tomato", "n.a")
    val_sawi = year_data.get("sawi", "n.a")

    sheet.range((STATE_START_ROW, COL_DURIAN)).value = float(val_durian) if pd.notna(val_durian) and val_durian not in ["n.a", ""] else "n.a"
    sheet.range((STATE_START_ROW, COL_MANGGA)).value = float(val_mangga) if pd.notna(val_mangga) and val_mangga not in ["n.a", ""] else "n.a"
    sheet.range((STATE_START_ROW, COL_NANAS)).value = float(val_nanas) if pd.notna(val_nanas) and val_nanas not in ["n.a", ""] else "n.a"
    sheet.range((STATE_START_ROW, COL_CILI)).value = float(val_cili) if pd.notna(val_cili) and val_cili not in ["n.a", ""] else "n.a"
    sheet.range((STATE_START_ROW, COL_KOBIS_BULAT)).value = float(val_kobis_bulat) if pd.notna(val_kobis_bulat) and val_kobis_bulat not in ["n.a", ""] else "n.a"
    sheet.range((STATE_START_ROW, COL_TIMUN)).value = float(val_timun) if pd.notna(val_timun) and val_timun not in ["n.a", ""] else "n.a"
    sheet.range((STATE_START_ROW, COL_TOMATO)).value = float(val_tomato) if pd.notna(val_tomato) and val_tomato not in ["n.a", ""] else "n.a"
    sheet.range((STATE_START_ROW, COL_SAWI)).value = float(val_sawi) if pd.notna(val_sawi) and val_sawi not in ["n.a", ""] else "n.a"

    # 3. District-Level Data Injection & Cleanup
    total_districts = len(districts)

    for district_idx, base_row in enumerate(BLOCK_START_ROWS):
        if district_idx < total_districts:
            dist_code = districts[district_idx]['code']
            sheet.range((base_row, COL_DISTRICT)).value = districts[district_idx]['name']

            dist_metrics = get_metrics_dict(dist_code, level='daerah', parent_code=state_code)
            year_data = dist_metrics.get("2024", {})

            val_durian = year_data.get("durian", "n.a")
            val_mangga = year_data.get("mangga", "n.a")
            val_nanas = year_data.get("nanas", "n.a")
            val_cili = year_data.get("cili", "n.a")
            val_kobis_bulat = year_data.get("kobis_bulat", "n.a")
            val_timun = year_data.get("timun", "n.a")
            val_tomato = year_data.get("tomato", "n.a")
            val_sawi = year_data.get("sawi", "n.a")

            sheet.range((base_row, COL_DURIAN)).value = float(val_durian) if pd.notna(val_durian) and val_durian not in ["n.a", ""] else "n.a"
            sheet.range((base_row, COL_MANGGA)).value = float(val_mangga) if pd.notna(val_mangga) and val_mangga not in ["n.a", ""] else "n.a"
            sheet.range((base_row, COL_NANAS)).value = float(val_nanas) if pd.notna(val_nanas) and val_nanas not in ["n.a", ""] else "n.a"
            sheet.range((base_row, COL_CILI)).value = float(val_cili) if pd.notna(val_cili) and val_cili not in ["n.a", ""] else "n.a"
            sheet.range((base_row, COL_KOBIS_BULAT)).value = float(val_kobis_bulat) if pd.notna(val_kobis_bulat) and val_kobis_bulat not in ["n.a", ""] else "n.a"
            sheet.range((base_row, COL_TIMUN)).value = float(val_timun) if pd.notna(val_timun) and val_timun not in ["n.a", ""] else "n.a"
            sheet.range((base_row, COL_TOMATO)).value = float(val_tomato) if pd.notna(val_tomato) and val_tomato not in ["n.a", ""] else "n.a"
            sheet.range((base_row, COL_SAWI)).value = float(val_sawi) if pd.notna(val_sawi) and val_sawi not in ["n.a", ""] else "n.a"
        else:
            # Wipe the dummy row clean if no district exists for this slot (single row only, no year offset)
            sheet.range((base_row, COL_DISTRICT), (base_row, COL_SAWI)).value = None

    # 4. Dynamic Page Elimination (Cleanup unused pages)
    # This table only has 2 pages: Page 1 holds 20 districts, Page 2 holds 20 more (40 total)
    if total_districts <= 20:
        # State only needs 1 page. Delete Page 2 entirely (page 2 title starts row 72).
        print("     [FORMAT] State only needs 1 page. Eliminating Page 2.")
        sheet.range('71:300').delete()