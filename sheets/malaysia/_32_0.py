from src.data_provider import get_metrics_dict
from src.excel_utils import safe_write

# ============================================================
# 1. SHARED CONFIGURATION (used by multiple tables)
# ============================================================

# CHANGE HERE: Update this list to change the report period.
# This list is used for titles, row generation, and Malaysia sum calculations.
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
# 2. JADUAL 32.0 – TABLE‑SPECIFIC CONFIGURATION
# ============================================================

# CHANGE HERE: Update this dictionary if the Excel column order changes.
# Key = column index (E=5, F=6, G=7, H=8, I=9), Value = metric name in the database.
COL_MAP = {
    5: "jumlah_kemalangan_jalan_raya",
    7: "jumlah_kecederaan_dan_kematian",
    8: "jumlah_kecederaan",
    9: "jumlah_kematian_jalan_raya"   # Malaysia uses this; states use "jumlah_kematian"
}

# CHANGE HERE: Add or remove state codes to match the template.
# Use the numeric two‑digit code (e.g., "01" for Johor).
# The order here must match the order of states in your Excel sheet.
# (Do NOT include "00" – Malaysia is handled separately)
LOCATIONS = [
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

# CHANGE HERE: Adjust this number if the table headers shift (e.g., extra blank rows added).
# This is the first row where data for Malaysia begins (row 11 in the current template).
START_ROW = 11

# Build ROW_MAP with Malaysia as the first block, then all states.
# Malaysia is special: it will be fetched from code "00" separately (not summed).
ALL_LOCATIONS = [("Malaysia", "malaysia")] + LOCATIONS
ROW_MAP = generate_row_map(start_row=START_ROW, locations=ALL_LOCATIONS, years=YEARS, spacing=4)


# ============================================================
# 3. INJECTION ENGINE 
# ============================================================
def populate_jadual_32(sheet, hierarchy, report_type):
    print("  -> Populating Jadual 32.0 (Kemalangan Jalan Raya)")

    sheet["C2"] = ": Bilangan kemalangan jalan raya, kecederaan dan kematian yang dilaporkan mengikut"
    sheet["C3"] = f"negeri, Malaysia, {YEARS[0]} - {YEARS[-1]}"
    sheet["C4"] = f": Number of road accidents, injuries and deaths reported by state, Malaysia, {YEARS[0]} - {YEARS[-1]}"

    state_cache = {}
    for state_code, _ in LOCATIONS:
        state_data = get_metrics_dict(state_code, level="negeri")
        if state_data:
            state_cache[state_code] = state_data
        else:
            print(f"     [Warning] No data found for state {state_code}.")

    malaysia_data = get_metrics_dict("00", level="negeri")
    if not malaysia_data:
        print("     [Warning] No data found for Malaysia (00).")

    for row_idx, (location_code, year, level) in ROW_MAP.items():
        if location_code == "Malaysia" and level == "malaysia":
            year_data = malaysia_data.get(year, {}) if malaysia_data else {}
        else:
            state_data = state_cache.get(location_code, {})
            year_data = state_data.get(year, {})

        for col_idx, metric_name in COL_MAP.items():
            raw_val = year_data.get(metric_name, "n.a")

            if raw_val == "n.a" and metric_name == "jumlah_kematian_jalan_raya":
                raw_val = year_data.get("jumlah_kematian", "n.a")

            if raw_val is not None and str(raw_val).strip() not in ["", "n.a", "n.a.", "nan", "NaN"]:
                try:
                    val = float(raw_val)   
                except (ValueError, TypeError):
                    val = raw_val          
            else:
                val = "n.a"

            safe_write(sheet, row_idx, col_idx, val)