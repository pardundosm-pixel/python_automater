import pandas as pd
from src.data_provider import get_metrics_dict

# ============================================================
# 1. SHARED CONFIGURATION
# ============================================================

# CHANGE HERE: Update this list to change the report period.
YEARS = ["2023", "2024", "2025"]


# ============================================================
# 2. JADUAL 41.0 – TABLE‑SPECIFIC CONFIGURATION
# ============================================================

# CHANGE HERE: Map each Excel row number to the metric name.
# Rows 9-23: Kutipan zakat (Collection)
# Rows 28-42: Agihan zakat (Distribution)
ROW_MAP = {
    # Kutipan zakat (Rows 9-23)
    11: "kutipan",       # MALAYSIA
    12: "kutipan",      # Johor
    13: "kutipan",      # Kedah
    14: "kutipan",      # Kelantan
    15: "kutipan",      # Melaka
    16: "kutipan",      # Negeri Sembilan
    17: "kutipan",      # Pahang
    18: "kutipan",      # Pulau Pinang
    19: "kutipan",      # Perak
    20: "kutipan",      # Perlis
    21: "kutipan",      # Selangor
    22: "kutipan",      # Terengganu
    23: "kutipan",      # Sabah
    24: "kutipan",      # Sarawak
    25: "kutipan",      # Wilayah Persekutuan

    # Agihan zakat (Rows 28-42)
    30: "agihan",       # MALAYSIA
    31: "agihan",       # Johor
    32: "agihan",       # Kedah
    33: "agihan",       # Kelantan
    34: "agihan",       # Melaka
    35: "agihan",       # Negeri Sembilan
    36: "agihan",       # Pahang
    37: "agihan",       # Pulau Pinang
    38: "agihan",       # Perak
    39: "agihan",       # Perlis
    40: "agihan",       # Selangor
    41: "agihan",       # Terengganu
    42: "agihan",       # Sabah
    43: "agihan",       # Sarawak
    44: "agihan",       # Wilayah Persekutuan
}

# CHANGE HERE: Map each Excel column number to the year.
# Columns: E=5, F=6, G=7 (2023, 2024, 2025)
COL_MAP = {
    6: "2023",
    7: "2024",
    8: "2025"
}

# CHANGE HERE: List locations in the order they appear in ROW_MAP.
# For states, use the numeric code (e.g., "01" for Johor) and level="negeri".
# Malaysia is included but will be handled as a special case (summed from states).
LOCATIONS = [
    # Kutipan zakat section
    ("Malaysia", "malaysia"),   # Row 9
    ("01", "negeri"),           # Row 10: Johor
    ("02", "negeri"),           # Row 11: Kedah
    ("03", "negeri"),           # Row 12: Kelantan
    ("04", "negeri"),           # Row 13: Melaka
    ("05", "negeri"),           # Row 14: Negeri Sembilan
    ("06", "negeri"),           # Row 15: Pahang
    ("07", "negeri"),           # Row 16: Pulau Pinang
    ("08", "negeri"),           # Row 17: Perak
    ("09", "negeri"),           # Row 18: Perlis
    ("10", "negeri"),           # Row 19: Selangor
    ("11", "negeri"),           # Row 20: Terengganu
    ("12", "negeri"),           # Row 21: Sabah
    ("13", "negeri"),           # Row 22: Sarawak
    ("14", "negeri"),           # Row 23: Wilayah Persekutuan

    # Agihan zakat section
    ("Malaysia", "malaysia"),   # Row 28
    ("01", "negeri"),           # Row 29: Johor
    ("02", "negeri"),           # Row 30: Kedah
    ("03", "negeri"),           # Row 31: Kelantan
    ("04", "negeri"),           # Row 32: Melaka
    ("05", "negeri"),           # Row 33: Negeri Sembilan
    ("06", "negeri"),           # Row 34: Pahang
    ("07", "negeri"),           # Row 35: Pulau Pinang
    ("08", "negeri"),           # Row 36: Perak
    ("09", "negeri"),           # Row 37: Perlis
    ("10", "negeri"),           # Row 38: Selangor
    ("11", "negeri"),           # Row 39: Terengganu
    ("12", "negeri"),           # Row 40: Sabah
    ("13", "negeri"),           # Row 41: Sarawak
    ("14", "negeri"),           # Row 42: Wilayah Persekutuan
]


# ============================================================
# 3. INJECTION ENGINE (Malaysia = sum of states)
# ============================================================
def populate_jadual_41(sheet, hierarchy, report_type):
    print("  -> Populating Jadual 41.0 (Statistik kutipan dan agihan zakat)")

    # ==========================================================
    # TITLE CONFIGURATION (CHANGE THESE STRINGS IF THE WORDING CHANGES)
    # ==========================================================
    title_bm = ": Statistik kutipan dan agihan zakat mengikut negeri, Malaysia, 2023 - 2025"
    title_en = ": Zakat collection and distribution statistics by state, Malaysia, 2023 - 2025"
    sheet.range("C3").value = title_bm
    sheet.range("C4").value = title_en

    # ==========================================================
    # 1) FETCH ALL STATE DATA
    # ==========================================================
    # Extract unique state codes (excluding "Malaysia")
    unique_states = set()
    for loc, _ in LOCATIONS:
        if loc != "Malaysia":
            unique_states.add(loc)

    state_cache = {}
    for state_code in unique_states:
        state_data = get_metrics_dict(state_code, level="negeri")
        if state_data:
            state_cache[state_code] = state_data
        else:
            print(f"     [Warning] No data found for state {state_code}.")

    # ==========================================================
    # 2) COMPUTE MALAYSIA TOTALS (SUM OF ALL STATES)
    # ==========================================================
    metrics = ["kutipan", "agihan"]
    malaysia_totals = {}
    for year in COL_MAP.values():
        totals = {metric: 0 for metric in metrics}
        for state_data in state_cache.values():
            year_data = state_data.get(year, {})
            for metric in metrics:
                val = year_data.get(metric, 0)
                if pd.notna(val):
                    try:
                        totals[metric] += float(val)
                    except (ValueError, TypeError):
                        pass
        malaysia_totals[year] = totals

    # ==========================================================
    # 3) DATA INJECTION INTO EXCEL
    # ==========================================================
    row_keys = list(ROW_MAP.keys())
    loc_values = list(LOCATIONS)

    for row_idx, (location_code, level) in zip(row_keys, loc_values):
        metric_name = ROW_MAP[row_idx]   # "kutipan" or "agihan"

        if location_code == "Malaysia":
            # Use the computed Malaysia totals
            for col_idx, year in COL_MAP.items():
                raw_val = malaysia_totals.get(year, {}).get(metric_name, "n.a")
                val = "n.a"
                if pd.notna(raw_val) and str(raw_val).strip() not in ("", "n.a", "n.a.", "-"):
                    try:
                        val = float(raw_val)
                    except (ValueError, TypeError):
                        pass
                sheet.range((row_idx, col_idx)).value = val
        else:
            # Use state data
            state_data = state_cache.get(location_code, {})
            for col_idx, year in COL_MAP.items():
                year_data = state_data.get(year, {})
                raw_val = year_data.get(metric_name, "n.a")
                val = "n.a"
                if pd.notna(raw_val) and str(raw_val).strip() not in ("", "n.a", "n.a.", "-"):
                    try:
                        val = float(raw_val)
                    except (ValueError, TypeError):
                        pass
                sheet.range((row_idx, col_idx)).value = val