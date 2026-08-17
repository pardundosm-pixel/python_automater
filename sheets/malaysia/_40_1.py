import pandas as pd
from src.data_provider import get_metrics_dict

# ============================================================
# 1. SHARED CONFIGURATION
# ============================================================
# CHANGE HERE: Update this list if the year changes (though title says 2024)
YEARS = ["2024"]


def generate_single_row_map(start_row, locations):
    """
    Generates a ROW_MAP for tables where each location occupies exactly one row.
    - No blank rows between locations.
    - Each entry maps row_index -> (location_code, level)
    """
    row_map = {}
    current_row = start_row
    for location_code, level in locations:
        row_map[current_row] = (location_code, level)
        current_row += 1
    return row_map


# ============================================================
# 2. JADUAL 40.1 – TABLE‑SPECIFIC CONFIG
# ============================================================

# CHANGE HERE: Map Excel column indices to metric names.
# Columns are: Age 5-17 (Dose1, Dose2, Booster1, Booster2),
# Age 18-59 (same four), Age 60+ (same four).
# Adjust the column numbers to match your actual template.
COL_MAP = {
    4: "umur_17_dos1",
    5: "umur_17_dos2",
    6: "umur_17_booster1",
    7: "umur_17_booster2",
    
    9: "umur_18_dos1",
    10: "umur_18_dos2",
    11: "umur_18_booster1",
    12: "umur_18_booster2",
    
    14: "umur_60_dos1",
    15: "umur_60_dos2",
    16: "umur_60_booster1",
    17: "umur_60_booster2",
}

# CHANGE HERE: List locations in the order they appear in the template.
# For Malaysia, use ("Malaysia", "malaysia") – this will fetch from fact_metrics_malaysia.
# For states, use numeric code and level="negeri".
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

# CHANGE HERE: Adjust this number to the first row where data begins (Malaysia row).
# In the template, Malaysia is at row 9, then states start at row 12, etc.
START_ROW = 10

# Generate ROW_MAP – each location occupies one row.
ROW_MAP = generate_single_row_map(start_row=START_ROW, locations=LOCATIONS)


# ============================================================
# 3. INJECTION ENGINE (Malaysia fetched directly)
# ============================================================
def populate_jadual_40_1(sheet, hierarchy, report_type):
    print("  -> Populating Jadual 40.1 (Statistik penerima vaksin COVID-19 mengikut umur)")

    # ==========================================================
    # TITLE CONFIGURATION (CHANGE THESE STRINGS IF THE WORDING CHANGES)
    # ==========================================================
    title_bm = f": Statistik penerima vaksin COVID-19 mengikut umur, Malaysia, {YEARS[0]}"
    title_en = f": Statistics of COVID-19 vaccine recipient by age, Malaysia, {YEARS[0]}"
    sheet.range("C3").value = title_bm
    sheet.range("C4").value = title_en

    # ==========================================================
    # DATA FETCHING & CACHING (Malaysia direct, states direct)
    # ==========================================================
    # Fetch Malaysia data once
    malaysia_data = get_metrics_dict("Malaysia", level="malaysia")
    if not malaysia_data:
        print("     [Warning] No data found for Malaysia.")

    # Fetch each state once and cache
    state_cache = {}
    for code, level in LOCATIONS:
        if code != "Malaysia":   # skip Malaysia since we already fetched
            state_data = get_metrics_dict(code, level=level)
            if state_data:
                state_cache[code] = state_data
            else:
                print(f"     [Warning] No data found for state {code}.")

    # ==========================================================
    # DATA INJECTION INTO EXCEL
    # ==========================================================
    for row_idx, (location_code, level) in ROW_MAP.items():
        # Choose data source
        if location_code == "Malaysia" and level == "malaysia":
            # Use the directly fetched Malaysia data
            year_data = malaysia_data.get(YEARS[0], {}) if malaysia_data else {}
        else:
            # For states, get from cache
            state_data = state_cache.get(location_code, {})
            year_data = state_data.get(YEARS[0], {})

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