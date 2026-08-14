import pandas as pd
from src.data_provider import get_metrics_dict

# ==========================================
# 1. MAPPING & PAGINATION CONFIGURATION
# ==========================================
COL_DISTRICT = 2   # Column B
COL_YEAR = 4        # Column D

# Data columns E-I
COL_PENDAPATAN_PURATA = 5     # Column E - Purata (Mean income)
COL_PENDAPATAN_PENENGAH = 6   # Column F - Penengah (Median income)
COL_PERBELANJAAN = 7           # Column G - Perbelanjaan (Mean household consumption)
COL_KEMISKINAN = 8              # Column H - Kemiskinan (Incidence of absolute poverty)
COL_GINI = 9                     # Column I - Gini coefficient

# The dictionary key (0, 1) represents how many rows down from the block's start row the year sits.
# Note: 53.0 only has 2 years per block (2022, 2024), not 3 - so block height is 3 rows (2 data + 1 blank).
YEAR_OFFSETS = {
    0: "2022",
    1: "2024"
}

# The row where the State (Negeri) total is printed.
STATE_START_ROW = 9

# Block start rows for 53.0's layout:
# - Page 1: 20 district blocks (rows 12-69, block height 3)
# - Page 2: 20 district blocks (rows 86-143)
# - Page 3: 20 district blocks (rows 158-215)
BLOCK_START_ROWS = (
    [12 + (i * 3) for i in range(20)] +   # Page 1 blocks
    [86 + (i * 3) for i in range(20)] +   # Page 2 blocks
    [158 + (i * 3) for i in range(20)]    # Page 3 blocks
)

# Title cells: "bm"/"en" hold the description text in column C, one row apart.
TITLE_COORDINATES = [
    {"bm": "C1", "en": "C2", "is_samb": False},        # Page 1 Title
    {"bm": "C78", "en": "C79", "is_samb": True},        # Page 2 Title
    {"bm": "C150", "en": "C151", "is_samb": True}       # Page 3 Title
]

# ==========================================
# 2. REPORT INJECTION ENGINE
# ==========================================
def populate_jadual_53_0(sheet, hierarchy, report_type):
    state_name = hierarchy.get('state_name', 'Unknown State')
    state_code = hierarchy.get('state_code', '00')
    districts = hierarchy.get('districts', [])

    # 1. Title Generation
    title_bm_base = f": Pendapatan, perbelanjaan dan kemiskinan, {state_name}, 2022 dan 2024"
    title_en_base = f": Income, expenditure and poverty, {state_name}, 2022 and 2024"

    for coords in TITLE_COORDINATES:
        suffix_bm = " (samb.)" if coords["is_samb"] else ""
        suffix_en = " (cont'd)" if coords["is_samb"] else ""
        sheet.range(coords["bm"]).value = title_bm_base + suffix_bm
        sheet.range(coords["en"]).value = title_en_base + suffix_en

    # 2. State-Level Data Injection
    sheet.range((STATE_START_ROW, COL_DISTRICT)).value = state_name.upper()
    state_metrics = get_metrics_dict(state_code, level='negeri')

    for offset, year in YEAR_OFFSETS.items():
        target_row = STATE_START_ROW + offset
        year_data = state_metrics.get(year, {})

        val_pendapatan_purata = year_data.get("pendapatan_purata", "n.a")
        val_pendapatan_penengah = year_data.get("pendapatan_penengah", "n.a")
        val_perbelanjaan = year_data.get("perbelanjaan", "n.a")
        val_kemiskinan = year_data.get("kemiskinan", "n.a")
        val_gini = year_data.get("gini", "n.a")

        sheet.range((target_row, COL_PENDAPATAN_PURATA)).value = float(val_pendapatan_purata) if pd.notna(val_pendapatan_purata) and val_pendapatan_purata not in ["n.a", ""] else "n.a"
        sheet.range((target_row, COL_PENDAPATAN_PENENGAH)).value = float(val_pendapatan_penengah) if pd.notna(val_pendapatan_penengah) and val_pendapatan_penengah not in ["n.a", ""] else "n.a"
        sheet.range((target_row, COL_PERBELANJAAN)).value = float(val_perbelanjaan) if pd.notna(val_perbelanjaan) and val_perbelanjaan not in ["n.a", ""] else "n.a"
        sheet.range((target_row, COL_KEMISKINAN)).value = float(val_kemiskinan) if pd.notna(val_kemiskinan) and val_kemiskinan not in ["n.a", ""] else "n.a"
        sheet.range((target_row, COL_GINI)).value = float(val_gini) if pd.notna(val_gini) and val_gini not in ["n.a", ""] else "n.a"

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

                val_pendapatan_purata = year_data.get("pendapatan_purata", "n.a")
                val_pendapatan_penengah = year_data.get("pendapatan_penengah", "n.a")
                val_perbelanjaan = year_data.get("perbelanjaan", "n.a")
                val_kemiskinan = year_data.get("kemiskinan", "n.a")
                val_gini = year_data.get("gini", "n.a")

                sheet.range((target_row, COL_PENDAPATAN_PURATA)).value = float(val_pendapatan_purata) if pd.notna(val_pendapatan_purata) and val_pendapatan_purata not in ["n.a", ""] else "n.a"
                sheet.range((target_row, COL_PENDAPATAN_PENENGAH)).value = float(val_pendapatan_penengah) if pd.notna(val_pendapatan_penengah) and val_pendapatan_penengah not in ["n.a", ""] else "n.a"
                sheet.range((target_row, COL_PERBELANJAAN)).value = float(val_perbelanjaan) if pd.notna(val_perbelanjaan) and val_perbelanjaan not in ["n.a", ""] else "n.a"
                sheet.range((target_row, COL_KEMISKINAN)).value = float(val_kemiskinan) if pd.notna(val_kemiskinan) and val_kemiskinan not in ["n.a", ""] else "n.a"
                sheet.range((target_row, COL_GINI)).value = float(val_gini) if pd.notna(val_gini) and val_gini not in ["n.a", ""] else "n.a"
        else:
            # Wipe the dummy row clean if no district exists for this slot
            sheet.range((base_row, COL_DISTRICT), (base_row + len(YEAR_OFFSETS), COL_GINI)).value = None

    # 4. Dynamic Page Elimination (Cleanup unused pages)
    # Page 1 holds 20 districts, Page 2 holds 20 more (40 total), Page 3 holds up to 20 more (60 total)
    if total_districts <= 20:
        # State only needs 1 page. Delete Page 2 and Page 3 entirely (page 2 title starts row 78).
        print("     [FORMAT] State only needs 1 page. Eliminating Page 2 and 3.")
        sheet.range('77:300').delete()

    elif total_districts <= 40:
        # State needs 2 pages. Delete Page 3 entirely (page 3 title starts row 150).
        print("     [FORMAT] State needs 2 pages. Eliminating Page 3.")
        sheet.range('149:300').delete()