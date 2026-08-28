import pandas as pd
from src.data_provider import get_metrics_dict
from src.excel_utils import safe_write

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

STATE_ORDER = ["11", "12", "13"]
STATE_NAMES = {
    "11": "Terengganu",
    "12": "Sabah",
    "13": "Sarawak"
}

METRIC_LABELS = {
    "jumlah_volum_hujan": "Jumlah (mm)",
    "bilangan_hari_hujan": "Bil. Hari"
}

ROW_MAP = {
    # TERENGGANU (11)
    7:  ("Kerteh", "11", "jumlah_volum_hujan"),
    8:  ("Kerteh", "11", "bilangan_hari_hujan"),
    9:  ("Kuala Terengganu", "11", "jumlah_volum_hujan"),
    10: ("Kuala Terengganu", "11", "bilangan_hari_hujan"),
    # SABAH (12)
    12: ("Keningau", "12", "jumlah_volum_hujan"),
    13: ("Keningau", "12", "bilangan_hari_hujan"),
    14: ("Kota Kinabalu", "12", "jumlah_volum_hujan"),
    15: ("Kota Kinabalu", "12", "bilangan_hari_hujan"),
    16: ("Kudat", "12", "jumlah_volum_hujan"),
    17: ("Kudat", "12", "bilangan_hari_hujan"),
    18: ("Ranau", "12", "jumlah_volum_hujan"),
    19: ("Ranau", "12", "bilangan_hari_hujan"),
    20: ("Sandakan", "12", "jumlah_volum_hujan"),
    21: ("Sandakan", "12", "bilangan_hari_hujan"),
    22: ("Tawau", "12", "jumlah_volum_hujan"),
    23: ("Tawau", "12", "bilangan_hari_hujan"),
    # SARAWAK (13) – first half
    25: ("Bintulu", "13", "jumlah_volum_hujan"),
    26: ("Bintulu", "13", "bilangan_hari_hujan"),
    27: ("Kapit", "13", "jumlah_volum_hujan"),
    28: ("Kapit", "13", "bilangan_hari_hujan"),
    29: ("Kuching", "13", "jumlah_volum_hujan"),
    30: ("Kuching", "13", "bilangan_hari_hujan"),
    31: ("Limbang", "13", "jumlah_volum_hujan"),
    32: ("Limbang", "13", "bilangan_hari_hujan"),
    33: ("Miri", "13", "jumlah_volum_hujan"),
    34: ("Miri", "13", "bilangan_hari_hujan"),
    35: ("Mulu", "13", "jumlah_volum_hujan"),
    36: ("Mulu", "13", "bilangan_hari_hujan"),
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
def populate_jadual_38_0_2(sheet, hierarchy, report_type):
    print("  -> Populating Jadual 38.2 (Hujan) untuk Malaysia")

    sheet["C2"] = ": Volum hujan, Malaysia, 2022 - 2024 (samb.)"
    sheet["C3"] = ": Rainfall volume, Malaysia, 2022 - 2024 (cont'd)"

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