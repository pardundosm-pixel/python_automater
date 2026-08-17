import pandas as pd
from src.data_provider import get_metrics_dict

# ============================================================
# SHARED CONFIGURATION (change years here for ALL tables)
# ============================================================
# CHANGE HERE: Update this list to change the report period.
YEARS = ["2022", "2024"]


def generate_row_map(start_row, locations, years, spacing):
    """
    Generates a ROW_MAP dictionary dynamically.
    - Each location occupies len(years) rows (one per year).
    - One blank row separates each location block (spacing = len(years) + 1).
    - Row indices are calculated sequentially from start_row.
    """
    row_map = {}
    current_row = start_row
    for location_code, level in locations:
        for year in years:
            row_map[current_row] = (location_code, year, level)
            current_row += 1
        current_row += 1  # blank row between blocks
    return row_map


# ============================================================
# JADUAL 35.0 – TABLE‑SPECIFIC CONFIG
# ============================================================

# CHANGE HERE: Map Excel columns to database metric names.
# E=5 (Pendapatan Purata), F=6 (Pendapatan Penengah),
# G=7 (Perbelanjaan), H=8 (Kemiskinan), I=9 (Gini)
COL_MAP = {
    5: "pendapatan_purata",
    6: "pendapatan_penengah",
    7: "perbelanjaan",
    8: "kemiskinan",
    9: "gini"
}

# CHANGE HERE: List locations in the order they appear in the template.
# For 'negeri', use the numeric code (e.g., "01" for Johor) – NOT the name.
LOCATIONS = [
    ("Malaysia", "malaysia"),
    ("01", "negeri"),   # Johor
    ("02", "negeri"),   # Kedah
    ("03", "negeri"),   # Kelantan
    ("04", "negeri"),   # Melaka
    ("05", "negeri"),   # Negeri Sembilan
    ("06", "negeri"),   # Pahang
    ("07", "negeri"),   # Pulau Pinang
    ("08", "negeri"),   # Perak
    ("09", "negeri"),   # Perlis
    ("10", "negeri"),   # Selangor
    ("11", "negeri"),   # Terengganu
    ("12", "negeri"),   # Sabah
    ("13", "negeri"),   # Sarawak
    ("14", "negeri"),   # W.P. Kuala Lumpur
    ("15", "negeri"),   # W.P. Labuan
    ("16", "negeri"),   # W.P. Putrajaya
]

# CHANGE HERE: Adjust start_row if the title/header rows shift.
# In Jadual 35.0, data starts at row 11 (Malaysia).
START_ROW = 11

# CHANGE HERE: spacing should be len(YEARS) + 1 (2 years + 1 blank row = 3).
ROW_MAP = generate_row_map(start_row=START_ROW, locations=LOCATIONS, years=YEARS, spacing=3)


# ============================================================
# INJECTION ENGINE
# ============================================================
def populate_jadual_35(sheet, hierarchy, report_type):
    print("  -> Populating Jadual 35.0 (Pendapatan, perbelanjaan dan kemiskinan)")

    # ==========================================================
    # TITLE CONFIGURATION (CHANGE THESE STRINGS IF THE TITLE CHANGES)
    # The year range is automatically generated from the YEARS list.
    # ==========================================================

    # BM title – single row (as per template, C3)
    title_bm = f": Pendapatan, perbelanjaan dan kemiskinan, Malaysia, {YEARS[0]} dan {YEARS[-1]}"

    # EN title – single row (C4)
    title_en = f": Income, expenditure and poverty, Malaysia, {YEARS[0]} and {YEARS[-1]}"

    # Write the titles to the Excel sheet
    sheet.range("C3").value = title_bm
    sheet.range("C4").value = title_en

    # ==========================================================
    # DATA INJECTION (no changes needed here)
    # ==========================================================
    data_cache = {}

    for row_idx, (location_code, year, level) in ROW_MAP.items():
        if location_code not in data_cache:
            data_cache[location_code] = get_metrics_dict(location_code, level=level)
            if not data_cache[location_code]:
                print(f"     [Warning] No data found for {location_code}.")

        year_data = data_cache[location_code].get(year, {})

        for col_idx, metric_name in COL_MAP.items():
            raw_val = year_data.get(metric_name, "n.a")
            if pd.notna(raw_val) and str(raw_val).strip() not in ("", "n.a", "n.a.", "-"):
                try:
                    val = float(raw_val)
                except (ValueError, TypeError):
                    val = "n.a"
            else:
                val = "n.a"
            sheet.range((row_idx, col_idx)).value = val