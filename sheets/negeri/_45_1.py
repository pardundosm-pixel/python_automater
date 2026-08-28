import pandas as pd
from src.data_provider import get_metrics_dict
from src.excel_utils import safe_write

# ==========================================
# 1. MAPPING & PAGINATION CONFIGURATION
# ==========================================
COL_DISTRICT = 2   # Column B
COL_YEAR = 4        # Column D

# 45.1 has 6 data columns (E-J) instead of the 2 (H-I) in 42.1/45.0
COL_TENAGA_BURUH = 5             # Column E - Tenaga buruh
COL_PENDUDUK_BEKERJA = 6         # Column F - Penduduk bekerja
COL_PENGANGGUR = 7               # Column G - Penganggur
COL_LUAR_TENAGA_BURUH = 8        # Column H - Tenaga buruh luar
COL_KADAR_PENYERTAAN = 9         # Column I - Kadar penyertaan tenaga buruh (%)
COL_KADAR_PENGANGGURAN = 10      # Column J - Kadar pengangguran (%)

# The dictionary key (0, 1, 2) represents how many rows down from the block's start row the year sits.
YEAR_OFFSETS = {
    0: "2022",
    1: "2023",
    2: "2024"
}

# The row where the State (Negeri) total is printed.
STATE_START_ROW = 14

# Block start rows recalculated for 45.1's layout:
# - Page 1: 13 district blocks (rows 18-68, block height 4)
# - Page 2: 13 district blocks (rows 93-143)
# - Page 3: 14 district blocks (rows 168-222)
BLOCK_START_ROWS = (
    [18 + (i * 4) for i in range(13)] +   # Page 1 blocks
    [93 + (i * 4) for i in range(13)] +   # Page 2 blocks
    [168 + (i * 4) for i in range(14)]    # Page 3 blocks
)

# Title cells for 45.1: "bm"/"en" hold the description text in column C,
TITLE_COORDINATES = [
    {"bm": (3, 3), "en": (4, 3), "is_samb": False},      # Page 1 Title
    {"bm": (81, 3), "en": (82, 3), "is_samb": True},     # Page 2 Title
    {"bm": (157, 3), "en": (158, 3), "is_samb": True}    # Page 3 Title
]

# ==========================================
# REPORT INJECTION ENGINE
# ==========================================
def populate_jadual_45_1(sheet, hierarchy, report_type):
    state_name = hierarchy.get('state_name', 'Unknown State')
    state_code = hierarchy.get('state_code', '00')
    districts = hierarchy.get('districts', [])
    print(f"  -> Populating Jadual 45.1 (Tenaga Buruh Daerah) untuk {state_name}")

    # 1. Title Generation
    title_bm_base = f": Statistik utama tenaga buruh mengikut daerah pentadbiran, {state_name}, 2022r - 2024"
    title_en_base = f": Principal statistics of labour force by administrative district, {state_name}, 2022r - 2024"

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

        metrics = {
            COL_TENAGA_BURUH: year_data.get("tenaga_buruh", "n.a"),
            COL_PENDUDUK_BEKERJA: year_data.get("penduduk_bekerja", "n.a"),
            COL_PENGANGGUR: year_data.get("penganggur", "n.a"),
            COL_LUAR_TENAGA_BURUH: year_data.get("luar_tenaga_buruh", "n.a"),
            COL_KADAR_PENYERTAAN: year_data.get("peratus_kadar_penyertaan_tenaga_buruh", "n.a"),
            COL_KADAR_PENGANGGURAN: year_data.get("peratus_kadar_pengangguran", "n.a")
        }

        for col, val in metrics.items():
            final_val = float(val) if pd.notna(val) and val not in ["n.a", ""] else "n.a"
            safe_write(sheet, target_row, col, final_val)

    # 3. District-Level Data Injection & Safe Cleanup
    total_districts = len(districts)

    for district_idx, base_row in enumerate(BLOCK_START_ROWS):
        if district_idx < total_districts:
            dist_code = districts[district_idx]['code']
            safe_write(sheet, base_row, COL_DISTRICT, districts[district_idx]['name'])
            dist_metrics = get_metrics_dict(dist_code, level='daerah', parent_code=state_code)

            for offset, year in YEAR_OFFSETS.items():
                target_row = base_row + offset
                year_data = dist_metrics.get(year, {})

                metrics = {
                    COL_TENAGA_BURUH: year_data.get("tenaga_buruh", "n.a"),
                    COL_PENDUDUK_BEKERJA: year_data.get("penduduk_bekerja", "n.a"),
                    COL_PENGANGGUR: year_data.get("penganggur", "n.a"),
                    COL_LUAR_TENAGA_BURUH: year_data.get("luar_tenaga_buruh", "n.a"),
                    COL_KADAR_PENYERTAAN: year_data.get("peratus_kadar_penyertaan_tenaga_buruh", "n.a"),
                    COL_KADAR_PENGANGGURAN: year_data.get("peratus_kadar_pengangguran", "n.a")
                }

                for col, val in metrics.items():
                    final_val = float(val) if pd.notna(val) and val not in ["n.a", ""] else "n.a"
                    safe_write(sheet, target_row, col, final_val)
        else:
            # --- Openpyxl Safe XML Wiping ---
            for r in range(base_row, base_row + len(YEAR_OFFSETS)):
                for c in range(COL_DISTRICT, COL_KADAR_PENGANGGURAN + 1):
                    safe_write(sheet, r, c, None)

    # 4. XML-Level Page Deletion
    # Syntax: sheet.delete_rows(start_row_index, number_of_rows_to_delete)
    if total_districts <= 13:
        print("     [FORMAT] State only needs 1 page. Eliminating Page 2 and 3.")
        sheet.delete_rows(80, 220)  # Deletes 220 rows starting at row 80

    elif total_districts <= 26:
        print("     [FORMAT] State needs 2 pages. Eliminating Page 3.")
        sheet.delete_rows(156, 144) # Deletes 144 rows starting at row 156