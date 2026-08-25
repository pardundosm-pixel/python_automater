import pandas as pd
from src.data_provider import get_metrics_dict

# ============================================================
# CONFIGURATION
# ============================================================

VERBOSE = True
AUTO_CHECK = True

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


COL_MAP = {
    5: "positif_covid",
    6: "kematian_covid",
}

LOCATIONS = [
    ("01", "negeri"), ("02", "negeri"), ("03", "negeri"),
    ("04", "negeri"), ("05", "negeri"), ("06", "negeri"),
    ("07", "negeri"), ("08", "negeri"), ("09", "negeri"),
    ("10", "negeri"), ("11", "negeri"), ("12", "negeri"),
    ("13", "negeri"), ("14", "negeri"), ("15", "negeri"),
    ("16", "negeri"),
]

START_ROW = 9
ALL_LOCATIONS = [("Malaysia", "malaysia")] + LOCATIONS
ROW_MAP = generate_row_map(start_row=START_ROW, locations=ALL_LOCATIONS, years=YEARS, spacing=4)

# For state order in summary
STATE_ORDER = ["Malaysia"] + [code for code, _ in LOCATIONS]
STATE_NAMES = {
    "Malaysia": "Malaysia",
    "01": "Johor",
    "02": "Kedah",
    "03": "Kelantan",
    "04": "Melaka",
    "05": "Negeri Sembilan",
    "06": "Pahang",
    "07": "Pulau Pinang",
    "08": "Perak",
    "09": "Perlis",
    "10": "Selangor",
    "11": "Terengganu",
    "12": "Sabah",
    "13": "Sarawak",
    "14": "W.P. Kuala Lumpur",
    "15": "W.P. Labuan",
    "16": "W.P. Putrajaya",
}

METRIC_LABELS = {
    "positif_covid": "Positif",
    "kematian_covid": "Kematian"
}


# ============================================================
# HELPER FUNCTIONS
# ============================================================

def format_value(v):
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    return str(v)


def normalize(v):
    if v is None or v == "":
        return "n.a"
    try:
        return float(v)
    except (ValueError, TypeError):
        return str(v)


def auto_check_excel(sheet, row_map, col_map, expected_values):
    errors = []
    for row_idx, (location_code, year, level) in row_map.items():
        for col_idx, metric_name in col_map.items():
            actual = sheet.range((row_idx, col_idx)).value
            expected = expected_values.get((location_code, year, metric_name), "n.a")

            norm_actual = normalize(actual)
            norm_expected = normalize(expected)

            if isinstance(norm_actual, float) and isinstance(norm_expected, float):
                if abs(norm_actual - norm_expected) > 1e-9:
                    errors.append(
                        f"{location_code} | {year} | {metric_name} | expected: {norm_expected} | actual: {norm_actual}"
                    )
            else:
                if str(norm_actual) != str(norm_expected):
                    errors.append(
                        f"{location_code} | {year} | {metric_name} | expected: {norm_expected} | actual: {norm_actual}"
                    )

    if errors:
        print("\n" + "="*80)
        print("❌ MISMATCHES FOUND (Excel vs Database)")
        print("="*80)
        for e in errors:
            print(e)
        print("="*80 + "\n")
        return False
    else:
        print("\n" + "="*80)
        print("✅ ALL VALUES MATCH! (Excel is correct)")
        print("="*80 + "\n")
        return True


# ============================================================
# MAIN FUNCTION
# ============================================================

def populate_jadual_40(sheet, hierarchy, report_type):
    print("  -> Populating Jadual 40.0 (Kes COVID 19)")

    # Titles
    title_bm = f": Bilangan kes positif dan kematian disebabkan COVID-19 mengikut negeri, Malaysia, {YEARS[0]} - {YEARS[-1]}"
    title_en = f": Number of positive cases and deaths due to COVID-19 by state, Malaysia, {YEARS[0]} - {YEARS[-1]}"
    sheet.range("C3").value = title_bm
    sheet.range("C4").value = title_en

    # Data fetch
    state_cache = {}
    for state_code, _ in LOCATIONS:
        state_data = get_metrics_dict(state_code, level="negeri")
        if state_data:
            state_cache[state_code] = state_data
        else:
            print(f"     [Warning] No data found for state {state_code}.")

    # Compute Malaysia totals (sum of states)
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

    # Injection + store expected values
    expected_values = {}
    for row_idx, (location_code, year, level) in ROW_MAP.items():
        if location_code == "Malaysia" and level == "malaysia":
            year_data = malaysia_totals.get(year, {})
        else:
            state_data = state_cache.get(location_code, {})
            year_data = state_data.get(year, {})

        for col_idx, metric_name in COL_MAP.items():
            raw_val = year_data.get(metric_name, "n.a")
            expected_values[(location_code, year, metric_name)] = raw_val

            # Clean and write
            if pd.notna(raw_val) and str(raw_val).strip() not in ("", "n.a", "n.a.", "-"):
                try:
                    val = float(raw_val)
                except (ValueError, TypeError):
                    val = raw_val
            else:
                val = "n.a"
            sheet.range((row_idx, col_idx)).value = val

    # ---- Summary if VERBOSE ----
    if VERBOSE:
        # Build dataframe from expected values
        df = pd.DataFrame([
            {
                "Location": loc,
                "Year": year,
                "Metric": metric,
                "Value": expected_values.get((loc, year, metric), "n.a")
            }
            for (loc, year, _) in ROW_MAP.values()
            for metric in COL_MAP.values()
        ])
        # Pivot: Location, Year as rows, Metric as columns
        pivot = df.pivot_table(
            index=["Location", "Year"],
            columns="Metric",
            values="Value",
            aggfunc="first"
        ).fillna("n.a")
        # Rename metrics for display
        pivot = pivot.rename(columns=METRIC_LABELS)
        # Reorder locations to match Excel
        pivot = pivot.reset_index()
        pivot["Location"] = pd.Categorical(pivot["Location"], categories=STATE_ORDER, ordered=True)
        pivot = pivot.sort_values(["Location", "Year"]).set_index(["Location", "Year"])

        print("\n" + "="*80)
        print("SUMMARY TABLE (ordered by Excel layout)")
        print("="*80)
        current_loc = None
        for (loc, year), row in pivot.iterrows():
            if loc != current_loc:
                current_loc = loc
                print(f"\n--- {STATE_NAMES.get(loc, loc)} ---")
            # Format values
            pos = format_value(row.get("Positif", "n.a"))
            death = format_value(row.get("Kematian", "n.a"))
            print(f"{year}  {pos:>12}  {death:>12}")
        print("\n" + "="*80)

    # ---- Auto-check ----
    if AUTO_CHECK:
        auto_check_excel(sheet, ROW_MAP, COL_MAP, expected_values)