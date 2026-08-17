import pandas as pd
from src.data_provider import get_metrics_dict

# ============================================================
# 1. SHARED CONFIGURATION
# ============================================================
YEARS = ["2021", "2022", "2023"]


def generate_row_map(start_row, locations, years, spacing=4):
    row_map = {}
    current_row = start_row
    for location_code, level in locations:
        for year in years:
            row_map[current_row] = (location_code, year, level)
            current_row += 1
        current_row += 1  # blank row
    return row_map


# ============================================================
# 2. JADUAL 40.0 – TABLE‑SPECIFIC CONFIG
# ============================================================
COL_MAP = {
    5: "positif_covid",
    6: "kematian_covid",
}

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

START_ROW = 9   # Malaysia block starts here
ALL_LOCATIONS = [("Malaysia", "malaysia")] + LOCATIONS
ROW_MAP = generate_row_map(start_row=START_ROW, locations=ALL_LOCATIONS, years=YEARS, spacing=4)


# ============================================================
# 3. INJECTION ENGINE (Malaysia = sum of states)
# ============================================================
def populate_jadual_40(sheet, hierarchy, report_type):
    print("  -> Populating Jadual 40.0 (Kes COVID 19)")

    # Titles
    title_bm = f": Bilangan kes positif dan kematian disebabkan COVID-19 mengikut negeri, Malaysia, {YEARS[0]} - {YEARS[-1]}"
    title_en = f": Number of positive cases and deaths due to COVID-19 by state, Malaysia, {YEARS[0]} - {YEARS[-1]}"
    sheet.range("C3").value = title_bm
    sheet.range("C4").value = title_en

    # 1) Fetch all state data once
    state_cache = {}
    for state_code, _ in LOCATIONS:
        state_data = get_metrics_dict(state_code, level="negeri")
        if state_data:
            state_cache[state_code] = state_data
        else:
            print(f"     [Warning] No data found for state {state_code}.")

    # 2) Compute Malaysia totals by summing states
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

    # 3) Inject data
    for row_idx, (location_code, year, level) in ROW_MAP.items():
        if location_code == "Malaysia" and level == "malaysia":
            year_data = malaysia_totals.get(year, {})
        else:
            state_data = state_cache.get(location_code, {})
            year_data = state_data.get(year, {})

        for col_idx, metric_name in COL_MAP.items():
            raw_val = year_data.get(metric_name, "n.a")
            val = "n.a"
            if pd.notna(raw_val) and str(raw_val).strip() not in ("", "n.a", "n.a.", "-"):
                try:
                    val = float(raw_val)
                except (ValueError, TypeError):
                    pass
            sheet.range((row_idx, col_idx)).value = val