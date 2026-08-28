import pandas as pd
from src.data_provider import get_metrics_dict, get_pdrm_hierarchy
from src.excel_utils import safe_write, inject_static_table

# ==========================================
# 1. MAPPING & PAGINATION CONFIGURATION
# ==========================================
# Column B holds the PDRM District Names
COL_DISTRICT = 2  

# Map the EXCEL COLUMN NUMBERS to the exact Metric Names in your database
# TODO: Update the dictionary values to match the metric names in fact_metrics_daerah_pdrm
COL_METRICS_MAP = {
    5: "jumlah_jenayah_kekerasan",  # Column E (Jumlah)
    6: "bunuh",           # Column F (Bunuh)
    7: "rogol",           # Column G (Rogol)
    8: "samun",           # Column H (Samun)
    9: "mencederakan"     # Column I (Mencederakan)
}

# The dictionary key (0, 1, 2) represents how many rows down from the block's start row the year sits.
YEAR_OFFSETS = {
    0: "2023",
    1: "2024",
    2: "2025"
}

# The exact row where the State (Negeri) total block begins
STATE_START_ROW = 10

# Recalculated block start rows based on the Jadual 50 screenshots:
# - Page 1 holds 13 districts. Starts at Row 14, spaced every 4 rows (14, 18, 22...).
# - Page 2 starts at Row 81, spaced every 4 rows.
# - Page 3 (estimated) starts around Row 148, spaced every 4 rows.
BLOCK_START_ROWS = (
    [14 + (i * 4) for i in range(13)] +  
    [81 + (i * 4) for i in range(15)] +  
    [148 + (i * 4) for i in range(15)]   
)

# Coordinates for the bilingual titles
TITLE_COORDINATES = [
    {"bm": (2, 3), "en": (4, 3), "is_samb": False},     # Page 1 Titles
    {"bm": (74, 3), "en": (75, 3), "is_samb": True}     # Page 2 Titles
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
def populate_jadual_50_0(sheet, hierarchy, report_type):
    state_name = hierarchy.get('state_name', 'Unknown State')
    state_code = hierarchy.get('state_code', '00')
    print(f"  -> Populating Jadual 51.0 (Kemalangan Jalan Raya) untuk {state_name}")

    pdrm_districts = get_pdrm_hierarchy(state_code)

    title_bm_prefix = ": Bilangan kemalangan jalan raya, kecederaan dan kematian yang dilaporkan mengikut daerah PDRM, "
    title_bm_suffix_base = f"  {state_name}, 2023 - 2025"
    title_en_base = f": Number of road accidents, injuries and deaths reported by PDRM district, {state_name}, 2023 - 2025"

    for coords in TITLE_COORDINATES:
        suffix_bm = " (samb.)" if coords["is_samb"] else ""
        suffix_en = " (cont'd)" if coords["is_samb"] else ""
        safe_write(sheet, coords["bm"][0], coords["bm"][1], title_bm_prefix)
        safe_write(sheet, coords["bm"][0], coords["bm"][1], title_bm_suffix_base + suffix_bm)
        safe_write(sheet, coords["en"][0], coords["en"][1], title_en_base + suffix_en)

    safe_write(sheet, STATE_START_ROW, COL_DISTRICT, state_name.upper())
    state_metrics = get_metrics_dict(state_code, level='negeri')

    for offset, year in YEAR_OFFSETS.items():
        target_row = STATE_START_ROW + offset
        year_data = state_metrics.get(year, {})

        for col_idx, metric_name in COL_METRICS_MAP.items():
            raw_val = year_data.get(metric_name, "n.a")
            safe_write(sheet, target_row, col_idx, sanitize_value(raw_val))

    total_districts = len(pdrm_districts)

    for district_idx, base_row in enumerate(BLOCK_START_ROWS):
        if district_idx < total_districts:
            dist_code = pdrm_districts[district_idx]['code']
            safe_write(sheet, base_row, COL_DISTRICT, pdrm_districts[district_idx]['name'])

            branch_metrics = get_metrics_dict(dist_code, level='pdrm', parent_code=state_code)

            for offset, year in YEAR_OFFSETS.items():
                target_row = base_row + offset
                year_data = branch_metrics.get(year, {})

                for col_idx, metric_name in COL_METRICS_MAP.items():
                    raw_val = year_data.get(metric_name, "n.a")
                    safe_write(sheet, target_row, col_idx, sanitize_value(raw_val))
        else:
            # XML-Level Nested Cleanup
            safe_write(sheet, base_row, COL_DISTRICT, None)
            for offset in YEAR_OFFSETS.keys():
                for c in range(min(COL_METRICS_MAP.keys()), max(COL_METRICS_MAP.keys()) + 1):
                    safe_write(sheet, base_row + offset, c, None)

    # XML-Level Page Deletion (Start Row, Delete Amount)
    if total_districts <= 13:
        print("     [FORMAT] State only needs 1 page. Eliminating Page 2 and 3.")
        sheet.delete_rows(73, 227) # Wipes from row 73 downwards
    elif total_districts <= 26:
        print("     [FORMAT] State needs 2 pages. Eliminating Page 3.")
        sheet.delete_rows(140, 160) # Wipes from row 140 downwards