import pandas as pd
from src.data_provider import get_metrics_dict
from src.excel_utils import safe_write

# ==========================================
# CONFIGURATION
# ==========================================

VERBOSE = True
AUTO_CHECK = True

COL_MAP = {
    6: "2022",
    7: "2023",
    8: "2024"
}

STATE_ORDER = ["11", "12", "13", "15"]
STATE_NAMES = {
    "11": "Terengganu",
    "12": "Sabah",
    "13": "Sarawak",
    "15": "W.P. Labuan"
}

METRIC_LABELS = {
    "purata_kelembapan_relatif": "Kelembapan (%)"
}

ROW_MAP = {
    # TERENGGANU (11)
    8:  ("Kerteh", "11", "purata_kelembapan_relatif"),
    9:  ("Kuala Terengganu", "11", "purata_kelembapan_relatif"),
    # SABAH (12)
    11: ("Keningau", "12", "purata_kelembapan_relatif"),
    12: ("Kota Kinabalu", "12", "purata_kelembapan_relatif"),
    13: ("Kudat", "12", "purata_kelembapan_relatif"),
    14: ("Ranau", "12", "purata_kelembapan_relatif"),
    15: ("Sandakan", "12", "purata_kelembapan_relatif"),
    16: ("Tawau", "12", "purata_kelembapan_relatif"),
    # SARAWAK (13)
    18: ("Bintulu", "13", "purata_kelembapan_relatif"),
    19: ("Kapit", "13", "purata_kelembapan_relatif"),
    20: ("Kuching", "13", "purata_kelembapan_relatif"),
    21: ("Limbang", "13", "purata_kelembapan_relatif"),
    22: ("Miri", "13", "purata_kelembapan_relatif"),
    23: ("Mulu", "13", "purata_kelembapan_relatif"),
    24: ("Sibu", "13", "purata_kelembapan_relatif"),
    25: ("Sri Aman", "13", "purata_kelembapan_relatif"),
    # LABUAN (15)
    27: ("Labuan", "15", "purata_kelembapan_relatif"),
}

# ==========================================
# HELPER FUNCTIONS (same as above)
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
def populate_jadual_39_0_1(sheet, hierarchy, report_type):
    print("  -> Populating Jadual 39.1 (Purata Kelembapan Relatif) untuk Malaysia")

    sheet["C2"] = ": Purata kelembapan relatif Malaysia, 2022 - 2024 (samb.)"
    sheet["C3"] = ": Mean relative humidity, Malaysia, 2022 - 2024 (cont'd)"

    data_cache = {}

    for row_idx, (station_name, state_code, metric_name) in ROW_MAP.items():
        cache_key = f"{station_name}_{state_code}"
        if cache_key not in data_cache:
            data_cache[cache_key] = get_metrics_dict(location_code=station_name, level='meteorologi', parent_code=state_code)

        station_data = data_cache[cache_key]

        for col_idx, year in COL_MAP.items():
            raw_val = station_data.get(str(year), {}).get(metric_name, "n.a")

            if pd.notna(raw_val) and str(raw_val).strip() not in ["", "n.a", "n.a.", "nan", "NaN"]:
                try:
                    val = float(raw_val)
                except (ValueError, TypeError):
                    val = raw_val
            else:
                val = "n.a"

            safe_write(sheet, row_idx, col_idx, val)