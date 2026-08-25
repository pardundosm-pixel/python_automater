import pandas as pd
from src.data_provider import get_metrics_dict

# ============================================================
# 1. SHARED CONFIGURATION
# ============================================================
YEARS = ["2024"]


# ============================================================
# 2. JADUAL 40.1 – TABLE‑SPECIFIC CONFIG
# ============================================================

# UPDATE: Starts on Column D (index 4). 
# A(1), B(2), C(3) are ignored, so D(4) is the first data column.
COL_MAP = {
    4: "umur_17_dos1",   # D
    5: "umur_17_dos2",   # E
    6: "umur_17_booster1", # F
    7: "umur_17_booster2", # G
    
    9: "umur_18_dos1",   # I
    10: "umur_18_dos2",  # J
    11: "umur_18_booster1", # K
    12: "umur_18_booster2", # L
    
    14: "umur_60_dos1",  # N
    15: "umur_60_dos2",  # O
    16: "umur_60_booster1", # P
    17: "umur_60_booster2", # Q
}

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

# UPDATE: Hardcoded ROW_MAP with the specific gaps.
# Row 10 is Malaysia, Rows 11 & 12 are blank, Row 13 is Johor, 
# then a blank row (14) before Kedah (15), etc.
ROW_MAP = {
    10: ("Malaysia", "malaysia"),
    13: ("01", "negeri"),   # Johor
    15: ("02", "negeri"),   # Kedah
    17: ("03", "negeri"),   # Kelantan
    19: ("04", "negeri"),   # Melaka
    21: ("05", "negeri"),   # Negeri Sembilan
    23: ("06", "negeri"),   # Pahang
    25: ("07", "negeri"),   # Pulau Pinang
    27: ("08", "negeri"),   # Perak
    29: ("09", "negeri"),   # Perlis
    31: ("10", "negeri"),   # Selangor
    33: ("11", "negeri"),   # Terengganu
    35: ("12", "negeri"),   # Sabah
    37: ("13", "negeri"),   # Sarawak
    39: ("14", "negeri"),   # W.P. Kuala Lumpur
    41: ("15", "negeri"),   # W.P. Labuan
    43: ("16", "negeri"),   # W.P. Putrajaya
}


# ============================================================
# 3. INJECTION ENGINE
# ============================================================
def populate_jadual_40_1(sheet, hierarchy, report_type):
    print("  -> Populating Jadual 40.1 (Statistik penerima vaksin COVID-19 mengikut umur)")

    # TITLES: Matches your specific C2 and C3 layout
    title_bm = f": Statistik penerima vaksin COVID-19 mengikut umur, Malaysia, {YEARS[0]}"
    title_en = f": Statistics of COVID-19 vaccine recipient by age, Malaysia, {YEARS[0]}"
    sheet.range("C2").value = title_bm
    sheet.range("C3").value = title_en

    # DATA FETCHING & CACHING
    malaysia_data = get_metrics_dict("Malaysia", level="malaysia")
    if not malaysia_data:
        print("     [Warning] No data found for Malaysia.")

    state_cache = {}
    for code, level in LOCATIONS:
        if code != "Malaysia":   
            state_data = get_metrics_dict(code, level=level)
            if state_data:
                state_cache[code] = state_data
            else:
                print(f"     [Warning] No data found for state {code}.")

    # DATA INJECTION INTO EXCEL
    for row_idx, (location_code, level) in ROW_MAP.items():
        # Choose data source
        if location_code == "Malaysia" and level == "malaysia":
            year_data = malaysia_data.get(YEARS[0], {}) if malaysia_data else {}
        else:
            state_data = state_cache.get(location_code, {})
            year_data = state_data.get(YEARS[0], {})

        # Fill each metric column starting at column D (4)
        for col_idx, metric_name in COL_MAP.items():
            raw_val = year_data.get(metric_name, "n.a")
            val = "n.a"
            if pd.notna(raw_val) and str(raw_val).strip() not in ("", "n.a", "n.a.", "-"):
                try:
                    val = float(raw_val)
                except (ValueError, TypeError):
                    pass
            sheet.range((row_idx, col_idx)).value = val