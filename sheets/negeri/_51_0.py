import pandas as pd
from src.data_provider import get_metrics_dict, get_pdrm_hierarchy

# ==========================================
# 1. MAPPING & PAGINATION CONFIGURATION
# ==========================================
# Column B holds the PDRM District Names
COL_DISTRICT = 2
COL_YEAR = 4   # Column D

# Map the EXCEL COLUMN NUMBERS to the exact Metric Names in your database
COL_METRICS_MAP = {
    5: "jumlah_kemalangan_jalan_raya",     # Column E (Kemalangan jalan raya)
    7: "jumlah_kecederaan_dan_kematian",    # Column G (Jumlah - Kecederaan dan kematian)
    8: "jumlah_kecederaan",                  # Column H (Kecederaan)
    9: "jumlah_kematian"                      # Column I (Kematian)
}
# Note: Column F is a blank spacer under the merged "Kecederaan dan kematian" header - not mapped.

# The dictionary key (0, 1, 2) represents how many rows down from the block's start row the year sits.
YEAR_OFFSETS = {
    0: "2023",
    1: "2024",
    2: "2025"
}

# The exact row where the State (Negeri) total block begins
STATE_START_ROW = 13

# Block start rows for 51.0's layout:
# - Page 1: 13 district blocks (rows 17-65, block height 4)
# - Page 2: 13 district blocks (rows 85-133)
# - Page 3: 14 district blocks (rows 157-209)
BLOCK_START_ROWS = (
    [17 + (i * 4) for i in range(13)] +
    [85 + (i * 4) for i in range(13)] +
    [157 + (i * 4) for i in range(14)]
)

# Title coordinates: the BM title is split across two cells (long prefix line, then a
# second cell holding the state name + years), while the EN title is a single cell.
TITLE_COORDINATES = [
    {"bm_prefix": "C2", "bm_suffix": "C3", "en": "C4", "is_samb": False},      # Page 1 Titles
    {"bm_prefix": "C74", "bm_suffix": "C75", "en": "C76", "is_samb": True},    # Page 2 Titles
    {"bm_prefix": "C146", "bm_suffix": "C147", "en": "C148", "is_samb": True}  # Page 3 Titles
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
def populate_jadual_51(sheet, hierarchy, report_type):
    state_name = hierarchy.get('state_name', 'Unknown State')
    state_code = hierarchy.get('state_code', '00')
    print(f"  -> Populating Jadual 51.0 (Kemalangan Jalan Raya) untuk {state_name}")

    # Fetch PDRM districts directly using the isolated domain dimension
    pdrm_districts = get_pdrm_hierarchy(state_code)

    # 1. Title Generation
    title_bm_prefix = ": Bilangan kemalangan jalan raya, kecederaan dan kematian yang dilaporkan mengikut daerah PDRM, "
    title_bm_suffix_base = f"  {state_name}, 2023 - 2025"
    title_en_base = f": Number of road accidents, injuries and deaths reported by PDRM district, {state_name}, 2023 - 2025"

    for coords in TITLE_COORDINATES:
        suffix_bm = " (samb.)" if coords["is_samb"] else ""
        suffix_en = " (cont'd)" if coords["is_samb"] else ""
        sheet.range(coords["bm_prefix"]).value = title_bm_prefix
        sheet.range(coords["bm_suffix"]).value = title_bm_suffix_base + suffix_bm
        sheet.range(coords["en"]).value = title_en_base + suffix_en

    # 2. State-Level Data Injection
    sheet.range((STATE_START_ROW, COL_DISTRICT)).value = state_name.upper()

    # Fetch State Total directly from the general Negeri fact table
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
            # --- Inject District Data ---
            dist_code = pdrm_districts[district_idx]['code']
            sheet.range((base_row, COL_DISTRICT)).value = pdrm_districts[district_idx]['name']

            branch_metrics = get_metrics_dict(dist_code, level='pdrm', parent_code=state_code)

            for offset, year in YEAR_OFFSETS.items():
                target_row = base_row + offset
                year_data = branch_metrics.get(year, {})

                for col_idx, metric_name in COL_METRICS_MAP.items():
                    raw_val = year_data.get(metric_name, "n.a")
                    sheet.range((target_row, col_idx)).value = sanitize_value(raw_val)
        else:
            # --- Cleanup Unused Row ---
            sheet.range((base_row, COL_DISTRICT)).value = None
            for offset in YEAR_OFFSETS.keys():
                sheet.range((base_row + offset, min(COL_METRICS_MAP.keys())),
                            (base_row + offset, max(COL_METRICS_MAP.keys()))).value = None

    # 4. Dynamic Page Elimination (Cleanup unused pages)
    # Page 1 holds 13 districts, Page 2 holds 13 more (26 total), Page 3 holds the remaining 14 (40 total)
    if total_districts <= 13:
        print("     [FORMAT] State only needs 1 page. Eliminating Page 2 and 3.")
        # Page 2 title starts row 74.
        sheet.range('73:300').delete()

    elif total_districts <= 26:
        print("     [FORMAT] State needs 2 pages. Eliminating Page 3.")
        # Page 3 title starts row 146.
        sheet.range('140:300').delete()