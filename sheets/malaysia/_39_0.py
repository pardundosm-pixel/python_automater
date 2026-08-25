import pandas as pd
from src.data_provider import get_metrics_dict

# ==========================================
# CONFIGURATION
# ==========================================

VERBOSE = True
AUTO_CHECK = True

# Column mapping: F=6, G=7, H=8
COL_MAP = {
    6: "2022",
    7: "2023",
    8: "2024"
}

# State order as they appear in this sheet
STATE_ORDER = ["01", "02", "03", "04", "05", "06", "07", "08", "09", "10"]
STATE_NAMES = {
    "01": "Johor",
    "02": "Kedah",
    "03": "Kelantan",
    "04": "Melaka",
    "05": "Negeri Sembilan",
    "06": "Pahang",
    "07": "Pulau Pinang",
    "08": "Perak",
    "09": "Perlis",
    "10": "Selangor"
}

# Metric label (only one metric for humidity)
METRIC_LABELS = {
    "purata_kelembapan_relatif": "Kelembapan (%)"
}

ROW_MAP = {
    # JOHOR (01)
    8:  ("Batu Pahat", "01", "purata_kelembapan_relatif"),
    9:  ("Kluang", "01", "purata_kelembapan_relatif"),
    10: ("Mersing", "01", "purata_kelembapan_relatif"),
    11: ("Senai", "01", "purata_kelembapan_relatif"),
    # KEDAH (02)
    13: ("Alor Setar", "02", "purata_kelembapan_relatif"),
    14: ("Pulau Langkawi", "02", "purata_kelembapan_relatif"),
    # KELANTAN (03)
    16: ("Kota Bharu", "03", "purata_kelembapan_relatif"),
    17: ("Kuala Krai", "03", "purata_kelembapan_relatif"),
    18: ("Gong Kedak", "03", "purata_kelembapan_relatif"),
    # MELAKA (04)
    20: ("Melaka", "04", "purata_kelembapan_relatif"),
    # NEGERI SEMBILAN (05)
    22: ("Kuala Pilah", "05", "purata_kelembapan_relatif"),
    # PAHANG (06)
    24: ("Cameron Highlands", "06", "purata_kelembapan_relatif"),
    25: ("Batu Embun, Jerantut", "06", "purata_kelembapan_relatif"),
    26: ("Kuantan", "06", "purata_kelembapan_relatif"),
    27: ("Muadzam Shah", "06", "purata_kelembapan_relatif"),
    28: ("Temerloh", "06", "purata_kelembapan_relatif"),
    # PULAU PINANG (07)
    30: ("Bayan Lepas", "07", "purata_kelembapan_relatif"),
    31: ("Butterworth", "07", "purata_kelembapan_relatif"),
    # PERAK (08)
    33: ("Ipoh", "08", "purata_kelembapan_relatif"),
    34: ("Lubok Merbau, Kuala Kangsar", "08", "purata_kelembapan_relatif"),
    35: ("Sitiawan", "08", "purata_kelembapan_relatif"),
    # PERLIS (09)
    37: ("Chuping", "09", "purata_kelembapan_relatif"),
    # SELANGOR (10)
    39: ("Petaling Jaya", "10", "purata_kelembapan_relatif"),
    40: ("Subang", "10", "purata_kelembapan_relatif"),
    41: ("KLIA Sepang", "10", "purata_kelembapan_relatif"),
}

# ==========================================
# HELPER FUNCTIONS
# ==========================================

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
    for row_idx, (station_name, state_code, metric_name) in row_map.items():
        for col_idx, year in col_map.items():
            actual = sheet.range((row_idx, col_idx)).value
            expected = expected_values.get((station_name, metric_name, year), "n.a")

            norm_actual = normalize(actual)
            norm_expected = normalize(expected)

            if isinstance(norm_actual, float) and isinstance(norm_expected, float):
                if abs(norm_actual - norm_expected) > 1e-9:
                    errors.append(
                        f"{station_name} | {metric_name} | {year} | expected: {norm_expected} | actual: {norm_actual}"
                    )
            else:
                if str(norm_actual) != str(norm_expected):
                    errors.append(
                        f"{station_name} | {metric_name} | {year} | expected: {norm_expected} | actual: {norm_actual}"
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

# ==========================================
# MAIN FUNCTION
# ==========================================

def populate_jadual_39(sheet, hierarchy, report_type):
    print("  -> Populating Jadual 39.0 (Purata Kelembapan Relatif) untuk Malaysia")

    sheet.range("C2").value = ": Purata kelembapan relatif, Malaysia, 2022 - 2024"
    sheet.range("C3").value = ": Mean relative humidity, Malaysia, 2022 - 2024"

    data_cache = {}
    expected_values = {}

    for row_idx, (station_name, state_code, metric_name) in ROW_MAP.items():
        cache_key = f"{station_name}_{state_code}"
        if cache_key not in data_cache:
            data_cache[cache_key] = get_metrics_dict(
                location_code=station_name,
                level='meteorologi',
                parent_code=state_code
            )
            if not data_cache[cache_key]:
                print(f"     [Warning] No data found for station: {station_name} in state {state_code}.")

        station_data = data_cache[cache_key]

        for col_idx, year in COL_MAP.items():
            year_data = station_data.get(str(year), {})
            raw_val = year_data.get(metric_name, "n.a")

            expected_values[(station_name, metric_name, year)] = raw_val

            if pd.notna(raw_val) and raw_val != "n.a" and raw_val != "":
                try:
                    val = float(raw_val)
                except (ValueError, TypeError):
                    val = raw_val
            else:
                val = "n.a"

            sheet.range((row_idx, col_idx)).value = val

    # ---- Summary ----
    if VERBOSE:
        # Since there's only one metric per station, we can simplify
        df = pd.DataFrame([
            {
                "State": state,
                "Station": station,
                "Year": year,
                "Value": expected_values.get((station, metric, year), "n.a")
            }
            for (station, state, metric) in ROW_MAP.values()
            for year in COL_MAP.values()
        ])
        pivot = df.pivot_table(
            index=["State", "Station"],
            columns="Year",
            values="Value",
            aggfunc="first"
        ).fillna("n.a")
        pivot = pivot.reset_index()
        pivot["State"] = pd.Categorical(pivot["State"], categories=STATE_ORDER, ordered=True)
        pivot = pivot.sort_values(["State", "Station"]).set_index(["State", "Station"])
        pivot = pivot[["2022", "2023", "2024"]]

        print("\n" + "="*80)
        print("SUMMARY TABLE (ordered by Excel layout)")
        print("="*80)
        current_state = None
        for (state, station), row in pivot.iterrows():
            if state != current_state:
                current_state = state
                print(f"\n--- {STATE_NAMES.get(state, state)} ---")
            print(f"{station:30}  {format_value(row['2022']):>10}  {format_value(row['2023']):>10}  {format_value(row['2024']):>10}")
        print("\n" + "="*80)

    # ---- Auto-check ----
    if AUTO_CHECK:
        auto_check_excel(sheet, ROW_MAP, COL_MAP, expected_values)