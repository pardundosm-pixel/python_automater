import pandas as pd
from src.data_provider import get_metrics_dict

# ==========================================
# CONFIGURATION
# ==========================================

VERBOSE = True
AUTO_CHECK = True

COL_MAP = {
    7: "2022",
    8: "2023",
    9: "2024"
}

# State order as they appear in this sheet
STATE_ORDER = ["06", "07", "08", "09", "10"]
STATE_NAMES = {
    "06": "Pahang",
    "07": "Pulau Pinang",
    "08": "Perak",
    "09": "Perlis",
    "10": "Selangor"
}

METRIC_LABELS = {
    "jumlah_volum_hujan": "Jumlah (mm)",
    "bilangan_hari_hujan": "Bil. Hari"
}

ROW_MAP = {
    # PAHANG (06)
    7:  ("Batu Embun, Jerantut", "06", "jumlah_volum_hujan"),
    8:  ("Batu Embun, Jerantut", "06", "bilangan_hari_hujan"),
    9:  ("Kuantan", "06", "jumlah_volum_hujan"),
    10: ("Kuantan", "06", "bilangan_hari_hujan"),
    11: ("Muadzam Shah", "06", "jumlah_volum_hujan"),
    12: ("Muadzam Shah", "06", "bilangan_hari_hujan"),
    13: ("Temerloh", "06", "jumlah_volum_hujan"),
    14: ("Temerloh", "06", "bilangan_hari_hujan"),
    # PULAU PINANG (07)
    16: ("Bayan Lepas", "07", "jumlah_volum_hujan"),
    17: ("Bayan Lepas", "07", "bilangan_hari_hujan"),
    18: ("Butterworth", "07", "jumlah_volum_hujan"),
    19: ("Butterworth", "07", "bilangan_hari_hujan"),
    # PERAK (08)
    21: ("Ipoh", "08", "jumlah_volum_hujan"),
    22: ("Ipoh", "08", "bilangan_hari_hujan"),
    23: ("Lubok Merbau, Kuala Kangsar", "08", "jumlah_volum_hujan"),
    24: ("Lubok Merbau, Kuala Kangsar", "08", "bilangan_hari_hujan"),
    25: ("Sitiawan", "08", "jumlah_volum_hujan"),
    26: ("Sitiawan", "08", "bilangan_hari_hujan"),
    # PERLIS (09)
    28: ("Chuping", "09", "jumlah_volum_hujan"),
    29: ("Chuping", "09", "bilangan_hari_hujan"),
    # SELANGOR (10)
    31: ("Petaling Jaya", "10", "jumlah_volum_hujan"),
    32: ("Petaling Jaya", "10", "bilangan_hari_hujan"),
    33: ("Subang", "10", "jumlah_volum_hujan"),
    34: ("Subang", "10", "bilangan_hari_hujan"),
    35: ("KLIA Sepang", "10", "jumlah_volum_hujan"),
    36: ("KLIA Sepang", "10", "bilangan_hari_hujan"),
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

def populate_jadual_38_0_1(sheet, hierarchy, report_type):
    print("  -> Populating Jadual 38 (Hujan) untuk Malaysia")

    sheet.range("C2").value = ": Volum hujan, Malaysia, 2022 - 2024 (samb.)"
    sheet.range("C3").value = ": Rainfall volume, Malaysia, 2022 - 2024 (cont'd)"

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
        df = pd.DataFrame([
            {
                "State": state,
                "Station": station,
                "Metric": METRIC_LABELS.get(metric, metric),
                "Year": year,
                "Value": expected_values.get((station, metric, year), "n.a")
            }
            for (station, state, metric) in ROW_MAP.values()
            for year in COL_MAP.values()
        ])
        pivot = df.pivot_table(
            index=["State", "Station", "Metric"],
            columns="Year",
            values="Value",
            aggfunc="first"
        ).fillna("n.a")
        pivot = pivot.reset_index()
        pivot["State"] = pd.Categorical(pivot["State"], categories=STATE_ORDER, ordered=True)
        pivot = pivot.sort_values(["State", "Station"]).set_index(["State", "Station", "Metric"])
        pivot = pivot[["2022", "2023", "2024"]]

        print("\n" + "="*80)
        print("SUMMARY TABLE (ordered by Excel layout)")
        print("="*80)
        current_state = None
        for (state, station, metric), row in pivot.iterrows():
            if state != current_state:
                current_state = state
                print(f"\n--- {STATE_NAMES.get(state, state)} ---")
            print(f"{station:20} {metric:12}  {format_value(row['2022']):>10}  {format_value(row['2023']):>10}  {format_value(row['2024']):>10}")
        print("\n" + "="*80)

    # ---- Auto-check ----
    if AUTO_CHECK:
        auto_check_excel(sheet, ROW_MAP, COL_MAP, expected_values)