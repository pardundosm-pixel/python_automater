import pandas as pd
from src.data_provider import get_metrics_dict

# ============================================================
# 1. SHARED CONFIGURATION (used by multiple tables)
# ============================================================

# CHANGE HERE: Update this list to change the report period.
# This list is used for titles, row generation, and Malaysia sum calculations.
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
# 2. JADUAL 32.0 – TABLE‑SPECIFIC CONFIGURATION
# ============================================================

# CHANGE HERE: Update this dictionary if the Excel column order changes.
# Key = column index (E=5, F=6, G=7, H=8, I=9), Value = metric name in the database.
COL_MAP = {
    5: "jumlah_kemalangan_jalan_raya",
    7: "jumlah_kecederaan_dan_kematian",
    8: "jumlah_kecederaan",
    9: "jumlah_kematian"
}

# CHANGE HERE: Add or remove state codes to match the template.
# Use the numeric two‑digit code (e.g., "01" for Johor).
# The order here must match the order of states in your Excel sheet.
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
# Malaysia is special: it will be computed by summing all states.
ALL_LOCATIONS = [("Malaysia", "malaysia")] + LOCATIONS
ROW_MAP = generate_row_map(start_row=START_ROW, locations=ALL_LOCATIONS, years=YEARS, spacing=4)


# ============================================================
# 3. INJECTION ENGINE
# ============================================================
def populate_jadual_32(sheet, hierarchy, report_type):
    print("  -> Populating Jadual 32.0 (Kemalangan Jalan Raya)")

    # ==========================================================
    # TITLE CONFIGURATION (CHANGE THESE STRINGS IF THE WORDING CHANGES)
    # The year range is automatically generated from the YEARS list.
    # ==========================================================
    # BM title is split across C2 and C3 (as per template)
    title_bm_part1 = ": Bilangan kemalangan jalan raya, kecederaan dan kematian yang dilaporkan mengikut"
    title_bm_part2 = f"negeri, Malaysia, {YEARS[0]} - {YEARS[-1]}"
    # EN title on C4 (single line)
    title_en = f": Number of road accidents, injuries and deaths reported by state, Malaysia, {YEARS[0]} - {YEARS[-1]}"

    sheet.range("C2").value = title_bm_part1
    sheet.range("C3").value = title_bm_part2
    sheet.range("C4").value = title_en

    # ==========================================================
    # DATA FETCHING & CACHING
    # ==========================================================
    # Fetch data for each state once and store in a cache.
    state_cache = {}
    for state_code, _ in LOCATIONS:
        state_data = get_metrics_dict(state_code, level="negeri")
        if state_data:
            state_cache[state_code] = state_data
        else:
            print(f"     [Warning] No data found for state {state_code}.")

    # ==========================================================
    # COMPUTE MALAYSIA TOTALS (SUM OF ALL STATES)
    # If Malaysia data is ever stored separately, you can replace this block
    # with a direct fetch from fact_metrics_malaysia.
    # CHANGE HERE: If Malaysia data becomes available in fact_metrics_malaysia,
    # comment out the summation block below and use get_metrics_dict("Malaysia", level="malaysia").
    # ==========================================================
    malaysia_totals = {}
    for year in YEARS:
        totals = {metric: 0 for metric in COL_MAP.values()}
        for state_data in state_cache.values():
            year_data = state_data.get(year, {})
            for metric in COL_MAP.values():
                val = year_data.get(metric, 0)
                if pd.notna(val):
                    try:
                        totals[metric] += float(val)
                    except (ValueError, TypeError):
                        pass
        malaysia_totals[year] = totals

    # ==========================================================
    # DATA INJECTION INTO EXCEL
    # ==========================================================
    for row_idx, (location_code, year, level) in ROW_MAP.items():
        # Choose the correct data source for this row
        if location_code == "Malaysia" and level == "malaysia":
            year_data = malaysia_totals.get(year, {})
        else:
            # For states, retrieve from the cache
            state_data = state_cache.get(location_code, {})
            year_data = state_data.get(year, {})

        # Fill each metric column for this row
        for col_idx, metric_name in COL_MAP.items():
            raw_val = year_data.get(metric_name, "n.a")
            val = "n.a"
            if pd.notna(raw_val) and str(raw_val).strip() not in ("", "n.a", "n.a.", "-"):
                try:
                    val = float(raw_val)
                except (ValueError, TypeError):
                    pass
            sheet.range((row_idx, col_idx)).value = val