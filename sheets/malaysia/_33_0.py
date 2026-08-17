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
# 2. JADUAL 33.0 – TABLE‑SPECIFIC CONFIGURATION
# ============================================================

# CHANGE HERE: Update this dictionary if the Excel column order changes.
# Columns: E=5 (Jumlah), F=6 (Penglihatan), G=7 (Pendengaran),
# H=8 (Fizikal), I=9 (Masalah pembelajaran), J=10 (Pertuturan),
# K=11 (Mental), L=12 (Ketidakupayaan Pelbagai).
COL_MAP = {
    5 : "jumlah_oku_berdaftar",      # Jumlah / Total
    6 : "oku_penglihatan",           # Penglihatan / Visually impaired
    7 : "oku_pendengaran",           # Pendengaran / Hearing impaired
    8 : "oku_fizikal",               # Fizikal / Physical
    9 : "oku_masalah_pembelajaran",  # Masalah pembelajaran / Learning disability
    10: "oku_pertuturan",           # Pertuturan / Speech
    11: "oku_mental",               # Mental
    12: "oku_pelbagai"              # Ketidakupayaan Pelbagai / Multiple disabilities
}

# CHANGE HERE: List states (codes) in the order they appear in the template.
# Note: W.P. Putrajaya is NOT a separate state – it's included in W.P. Kuala Lumpur (code 14).
# So we do NOT include code "16".
# Codes: 01=Johor, 02=Kedah, 03=Kelantan, 04=Melaka, 05=Negeri Sembilan,
# 06=Pahang, 07=Pulau Pinang, 08=Perak, 09=Perlis, 10=Selangor,
# 11=Terengganu, 12=Sabah, 13=Sarawak, 14=W.P. Kuala Lumpur, 15=W.P. Labuan.
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
]

# CHANGE HERE: Adjust this number if the table headers shift.
# This is the first row where data for Malaysia begins.
# In the current template, Malaysia starts at row 11.
START_ROW = 16

# Build ROW_MAP with Malaysia as the first block, then all states.
# Malaysia is special: it will be computed by summing all states.
ALL_LOCATIONS = [("Malaysia", "malaysia")] + LOCATIONS
ROW_MAP = generate_row_map(start_row=START_ROW, locations=ALL_LOCATIONS, years=YEARS, spacing=4)


# ============================================================
# 3. INJECTION ENGINE (with Malaysia sum calculation)
# ============================================================
def populate_jadual_33(sheet, hierarchy, report_type):
    print("  -> Populating Jadual 33.0 (Bilangan OKU Berdaftar)")

    # ==========================================================
    # TITLE CONFIGURATION (CHANGE THESE STRINGS IF THE WORDING CHANGES)
    # The year range is automatically generated from the YEARS list.
    # ==========================================================
    # BM title is split across C2 and C3 (as per template)
    title_bm_part1 = ": Bilangan kumulatif Orang Kurang Upaya (OKU) yang berdaftar mengikut negeri dan kategori"
    title_bm_part2 = f"ketidakupayaan, Malaysia, {YEARS[0]} - {YEARS[-1]}"
    # EN title on C4 (single line)
    title_en = f": Cumulative number of registered Persons with Disabilities (PWD) cases by state and category of disabilities, Malaysia, {YEARS[0]} - {YEARS[-1]}"

    sheet.range("C2").value = title_bm_part1
    sheet.range("C3").value = title_bm_part2
    sheet.range("C4").value = title_en

    # ==========================================================
    # DATA FETCHING & CACHING (each state once)
    # ==========================================================
    state_cache = {}
    for state_code, _ in LOCATIONS:
        state_data = get_metrics_dict(state_code, level="negeri")
        if state_data:
            state_cache[state_code] = state_data
        else:
            print(f"     [Warning] No data found for state {state_code}.")

    # ==========================================================
    # COMPUTE MALAYSIA TOTALS (SUM OF ALL STATES)
    # If Malaysia data ever becomes available in fact_metrics_malaysia,
    # you can replace this block with a direct fetch.
    # CHANGE HERE: To use fact_metrics_malaysia instead, comment out the
    # summation block and use:
    # malaysia_data = get_metrics_dict("Malaysia", level="malaysia")
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