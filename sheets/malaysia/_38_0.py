import pandas as pd
from src.data_provider import get_metrics_dict
from src.excel_utils import safe_write

# ==========================================
# CONFIGURATION
# ==========================================

# Set to True to print summary and auto-check; False to run silently
VERBOSE = True
AUTO_CHECK = True

# Map Excel column index to year
COL_MAP = {
    8: "2022",
    9: "2023",
    10: "2024"
}

# Order of states as they appear in Excel (used for summary ordering)
STATE_ORDER = ["01", "02", "03", "04", "05", "06"]
STATE_NAMES = {
    "01": "Johor",
    "02": "Kedah",
    "03": "Kelantan",
    "04": "Melaka",
    "05": "Negeri Sembilan",
    "06": "Pahang"
}

# Map metric codes to display labels
METRIC_LABELS = {
    "jumlah_volum_hujan": "Jumlah (mm)",
    "bilangan_hari_hujan": "Bil. Hari"
}

# Row mapping: (station_name, state_code, metric_name)
ROW_MAP = {
    # JOHOR
    7:  ("Batu Pahat", "01", "jumlah_volum_hujan"),
    8:  ("Batu Pahat", "01", "bilangan_hari_hujan"),
    9:  ("Kluang", "01", "jumlah_volum_hujan"),
    10: ("Kluang", "01", "bilangan_hari_hujan"),
    11: ("Mersing", "01", "jumlah_volum_hujan"),
    12: ("Mersing", "01", "bilangan_hari_hujan"),
    13: ("Senai", "01", "jumlah_volum_hujan"),
    14: ("Senai", "01", "bilangan_hari_hujan"),
    # KEDAH
    16: ("Alor Setar", "02", "jumlah_volum_hujan"),
    17: ("Alor Setar", "02", "bilangan_hari_hujan"),
    18: ("Pulau Langkawi", "02", "jumlah_volum_hujan"),
    19: ("Pulau Langkawi", "02", "bilangan_hari_hujan"),
    # KELANTAN
    21: ("Kota Bharu", "03", "jumlah_volum_hujan"),
    22: ("Kota Bharu", "03", "bilangan_hari_hujan"),
    23: ("Kuala Krai", "03", "jumlah_volum_hujan"),
    24: ("Kuala Krai", "03", "bilangan_hari_hujan"),
    25: ("Gong Kedak", "03", "jumlah_volum_hujan"),
    26: ("Gong Kedak", "03", "bilangan_hari_hujan"),
    # MELAKA
    28: ("Melaka", "04", "jumlah_volum_hujan"),
    29: ("Melaka", "04", "bilangan_hari_hujan"),
    # NEGERI SEMBILAN
    31: ("Kuala Pilah", "05", "jumlah_volum_hujan"),
    32: ("Kuala Pilah", "05", "bilangan_hari_hujan"),
    # PAHANG
    34: ("Cameron Highlands", "06", "jumlah_volum_hujan"),
    35: ("Cameron Highlands", "06", "bilangan_hari_hujan"),
}

# ==========================================
# HELPER FUNCTIONS
# ==========================================

def format_value(v):
    """Format a value for display: integers as ints, floats with .0 removed if whole."""
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    return str(v)

def normalize(v):
    """Normalise a value for comparison: convert to float if possible, else string."""
    if v is None or v == "":
        return "n.a"
    try:
        return float(v)
    except (ValueError, TypeError):
        return str(v)

def auto_check_excel(sheet, row_map, col_map, expected_values):
    """
    Reads back all written cells and compares with the expected values.
    Uses numeric comparison for numbers, string comparison for text.
    """
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
# MAIN INJECTION FUNCTION
# ==========================================
def populate_jadual_38(sheet, hierarchy, report_type):
    print("  -> Populating Jadual 38.0 (Hujan) untuk Malaysia")

    sheet["C2"] = ": Volum hujan, Malaysia, 2022 - 2024"
    sheet["C3"] = ": Rainfall volume, Malaysia, 2022 - 2024"

    data_cache = {}

    for row_idx, (station_name, state_code, metric_name) in ROW_MAP.items():
        cache_key = f"{station_name}_{state_code}"
        if cache_key not in data_cache:
            data_cache[cache_key] = get_metrics_dict(location_code=station_name, level='meteorologi', parent_code=state_code)
            if not data_cache[cache_key]:
                print(f"     [Warning] No data found for station: {station_name} in state {state_code}.")

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