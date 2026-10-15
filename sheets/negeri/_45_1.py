import pandas as pd
from src.data_provider import get_metrics_dict

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
# one row apart (row N = BM, row N+1 = EN).
TITLE_COORDINATES = [
    {"bm": "C3", "en": "C4", "is_samb": False},      # Page 1 Title
    {"bm": "C81", "en": "C82", "is_samb": True},      # Page 2 Title
    {"bm": "C157", "en": "C158", "is_samb": True}     # Page 3 Title
]

# ==========================================
# 2. REPORT INJECTION ENGINE
# ==========================================
def populate_jadual_45_1(sheet, hierarchy, report_type):
    state_name = hierarchy.get('state_name', 'Unknown State')
    state_code = hierarchy.get('state_code', '00')
    districts = hierarchy.get('districts', [])

    # 1. Title Generation
    title_bm_base = f": Statistik utama tenaga buruh mengikut daerah pentadbiran, {state_name}, 2022r - 2024"
    title_en_base = f": Principal statistics of labour force by administrative district, {state_name}, 2022r - 2024"

    for coords in TITLE_COORDINATES:
        suffix_bm = " (samb.)" if coords["is_samb"] else ""
        suffix_en = " (cont'd)" if coords["is_samb"] else ""
        sheet.range(coords["bm"]).value = title_bm_base + suffix_bm
        sheet.range(coords["en"]).value = title_en_base + suffix_en

    # 2. State-Level Data Injection
    sheet.range((STATE_START_ROW, COL_DISTRICT)).value = state_name.upper()
    state_metrics = get_metrics_dict("STATE_TOTAL", level='daerah', parent_code=state_code)

    for offset, year in YEAR_OFFSETS.items():
        target_row = STATE_START_ROW + offset
        year_data = state_metrics.get(year, {})

        val_tenaga_buruh = year_data.get("tenaga_buruh", "n.a")
        val_penduduk_bekerja = year_data.get("penduduk_bekerja", "n.a")
        val_penganggur = year_data.get("penganggur", "n.a")
        val_luar_tenaga_buruh = year_data.get("luar_tenaga_buruh", "n.a")
        val_kadar_penyertaan = year_data.get("peratus_kadar_penyertaan_tenaga_buruh", "n.a")
        val_kadar_pengangguran = year_data.get("peratus_kadar_pengangguran", "n.a")

        sheet.range((target_row, COL_TENAGA_BURUH)).value = float(val_tenaga_buruh) if pd.notna(val_tenaga_buruh) and val_tenaga_buruh not in ["n.a", ""] else "n.a"
        sheet.range((target_row, COL_PENDUDUK_BEKERJA)).value = float(val_penduduk_bekerja) if pd.notna(val_penduduk_bekerja) and val_penduduk_bekerja not in ["n.a", ""] else "n.a"
        sheet.range((target_row, COL_PENGANGGUR)).value = float(val_penganggur) if pd.notna(val_penganggur) and val_penganggur not in ["n.a", ""] else "n.a"
        sheet.range((target_row, COL_LUAR_TENAGA_BURUH)).value = float(val_luar_tenaga_buruh) if pd.notna(val_luar_tenaga_buruh) and val_luar_tenaga_buruh not in ["n.a", ""] else "n.a"
        sheet.range((target_row, COL_KADAR_PENYERTAAN)).value = float(val_kadar_penyertaan) if pd.notna(val_kadar_penyertaan) and val_kadar_penyertaan not in ["n.a", ""] else "n.a"
        sheet.range((target_row, COL_KADAR_PENGANGGURAN)).value = float(val_kadar_pengangguran) if pd.notna(val_kadar_pengangguran) and val_kadar_pengangguran not in ["n.a", ""] else "n.a"

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

                val_tenaga_buruh = year_data.get("tenaga_buruh", "n.a")
                val_penduduk_bekerja = year_data.get("penduduk_bekerja", "n.a")
                val_penganggur = year_data.get("penganggur", "n.a")
                val_luar_tenaga_buruh = year_data.get("luar_tenaga_buruh", "n.a")
                val_kadar_penyertaan = year_data.get("peratus_kadar_penyertaan_tenaga_buruh", "n.a")
                val_kadar_pengangguran = year_data.get("peratus_kadar_pengangguran", "n.a")

                sheet.range((target_row, COL_TENAGA_BURUH)).value = float(val_tenaga_buruh) if pd.notna(val_tenaga_buruh) and val_tenaga_buruh not in ["n.a", ""] else "n.a"
                sheet.range((target_row, COL_PENDUDUK_BEKERJA)).value = float(val_penduduk_bekerja) if pd.notna(val_penduduk_bekerja) and val_penduduk_bekerja not in ["n.a", ""] else "n.a"
                sheet.range((target_row, COL_PENGANGGUR)).value = float(val_penganggur) if pd.notna(val_penganggur) and val_penganggur not in ["n.a", ""] else "n.a"
                sheet.range((target_row, COL_LUAR_TENAGA_BURUH)).value = float(val_luar_tenaga_buruh) if pd.notna(val_luar_tenaga_buruh) and val_luar_tenaga_buruh not in ["n.a", ""] else "n.a"
                sheet.range((target_row, COL_KADAR_PENYERTAAN)).value = float(val_kadar_penyertaan) if pd.notna(val_kadar_penyertaan) and val_kadar_penyertaan not in ["n.a", ""] else "n.a"
                sheet.range((target_row, COL_KADAR_PENGANGGURAN)).value = float(val_kadar_pengangguran) if pd.notna(val_kadar_pengangguran) and val_kadar_pengangguran not in ["n.a", ""] else "n.a"
        else:
            # Wipe the dummy row clean if no district exists for this slot
            sheet.range((base_row, COL_DISTRICT), (base_row + len(YEAR_OFFSETS), COL_KADAR_PENGANGGURAN)).value = None

    # 4. Dynamic Page Elimination (Cleanup unused pages)
    # Page 1 holds 13 districts, Page 2 holds 13 more (26 total), Page 3 holds the remaining 14 (40 total)
    if total_districts <= 13:
        # State only needs 1 page. Delete Page 2 and Page 3 entirely (page 2 title starts row 81).
        print("     [FORMAT] State only needs 1 page. Eliminating Page 2 and 3.")
        sheet.range('80:300').delete()

    elif total_districts <= 26:
        # State needs 2 pages. Delete Page 3 entirely (page 3 title starts row 157).
        print("     [FORMAT] State needs 2 pages. Eliminating Page 3.")
        sheet.range('156:300').delete()