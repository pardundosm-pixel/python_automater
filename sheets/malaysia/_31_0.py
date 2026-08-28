from src.data_provider import get_metrics_dict
from src.excel_utils import safe_write

# ============================================================
# SHARED CONFIGURATION (change years here for ALL tables)
# ============================================================
# CHANGE HERE: Update this list to change the report period.
YEARS = ["2022", "2023", "2024"]


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
# JADUAL 31.0 – TABLE‑SPECIFIC CONFIG
# ============================================================

# CHANGE HERE: Map Excel columns to database metric names.
# E=5, F=6, G=7, H=8, I=9
COL_MAP = {
    5: "jumlah_jenayah_kekerasan",
    6: "bunuh",
    7: "rogol",
    8: "samun",
    9: "mencederakan"
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
# In Jadual 31.0, data starts at row 9 (Malaysia).
START_ROW = 9

# Generate the row map – no manual row numbers needed!
ROW_MAP = generate_row_map(start_row=START_ROW, locations=LOCATIONS, years=YEARS, spacing=4)


# ============================================================
# INJECTION ENGINE
# ============================================================
def populate_jadual_31(sheet, hierarchy, report_type):
    print("  -> Populating Jadual 31.0 (Jenayah Kekerasan)")

    # Titles (Openpyxl syntax)
    sheet["C3"] = f": Jenayah kekerasan mengikut negeri dan jenis jenayah, Malaysia, {YEARS[0]} - {YEARS[-1]}"
    sheet["C4"] = f": Assault crime by state and type of crime, Malaysia, {YEARS[0]} - {YEARS[-1]}"

    data_cache = {}

    for row_idx, (location_code, year, level) in ROW_MAP.items():
        if location_code not in data_cache:
            data_cache[location_code] = get_metrics_dict(location_code, level=level)
            if not data_cache[location_code]:
                print(f"     [Warning] No data found for {location_code}.")

        year_data = data_cache[location_code].get(year, {})

        for col_idx, metric_name in COL_MAP.items():
            raw_val = year_data.get(metric_name, "n.a")
            
            if raw_val is not None and str(raw_val).strip() not in ["", "n.a", "n.a.", "nan", "NaN"]:
                try:
                    val = float(raw_val)
                except (ValueError, TypeError):
                    val = raw_val          
            else:
                val = "n.a"
            
            safe_write(sheet, row_idx, col_idx, val)