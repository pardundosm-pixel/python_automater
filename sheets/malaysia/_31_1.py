import pandas as pd
from src.data_provider import get_metrics_dict

# ============================================================
# SHARED CONFIGURATION (change years here for ALL tables)
# ============================================================
# CHANGE HERE: Update this list to change the report period.
YEARS = ["2023", "2024", "2025"]


def generate_row_map(start_row, locations, years, spacing=4):
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
# JADUAL 31.1 – TABLE‑SPECIFIC CONFIG
# ============================================================

# CHANGE HERE: Map Excel columns to database metric names.
# E=5 (Jumlah), F=6 (Pecah rumah dan curi), G=7 (Lori/Van),
# H=8 (Kereta), I=9 (Motosikal/Skuter), J=10 (Kecurian lain)
COL_MAP = {
    5: "jumlah_jenayah_harta_benda",
    6: "pecah_rumah_curi",
    7: "kecurian_lori",
    8: "kecurian_kereta",
    9: "kecurian_moto",
    10: "kecurian_lain"
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
]

# CHANGE HERE: Adjust start_row if the title/header rows shift.
# In Jadual 31.1, data starts at row 11 (Malaysia).
START_ROW = 11

# Generate the row map – no manual row numbers needed!
ROW_MAP = generate_row_map(start_row=START_ROW, locations=LOCATIONS, years=YEARS, spacing=4)


# ============================================================
# INJECTION ENGINE
# ============================================================
def populate_jadual_31_1(sheet, hierarchy, report_type):
    print("  -> Populating Jadual 31.1 (Jenayah Harta Benda)")

    # ==========================================================
    # TITLE CONFIGURATION (CHANGE THESE STRINGS IF THE TITLE CHANGES)
    # The year range is automatically generated from the YEARS list.
    # ==========================================================

    # BM title – single line (C3)
    title_bm = f": Jenayah harta benda mengikut negeri dan jenis jenayah, Malaysia, {YEARS[0]} - {YEARS[-1]}"

    # EN title – single line (C4)
    title_en = f": Property crime by state and type of crime, Malaysia, {YEARS[0]} - {YEARS[-1]}"

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