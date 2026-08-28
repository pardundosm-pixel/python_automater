import pandas as pd
from src.data_provider import get_metrics_dict
from src.excel_utils import safe_write

COL_MAP = {
    8: "2022",
    9: "2023",
    10: "2024"
}

ROW_MAP = {
    # SABAH
    7:  ("Keningau", "12", "minimum_suhu"),
    8:  ("Keningau", "12", "maksimum_suhu"),
    9:  ("Kota Kinabalu", "12", "minimum_suhu"),
    10: ("Kota Kinabalu", "12", "maksimum_suhu"),
    11: ("Kudat", "12", "minimum_suhu"),
    12: ("Kudat", "12", "maksimum_suhu"),
    13: ("Ranau", "12", "minimum_suhu"),
    14: ("Ranau", "12", "maksimum_suhu"),
    15: ("Sandakan", "12", "minimum_suhu"),
    16: ("Sandakan", "12", "maksimum_suhu"),
    17: ("Tawau", "12", "minimum_suhu"),
    18: ("Tawau", "12", "maksimum_suhu"),
    # SARAWAK
    20: ("Bintulu", "13", "minimum_suhu"),
    21: ("Bintulu", "13", "maksimum_suhu"),
    22: ("Kapit", "13", "minimum_suhu"),
    23: ("Kapit", "13", "maksimum_suhu"),
    24: ("Kuching", "13", "minimum_suhu"),
    25: ("Kuching", "13", "maksimum_suhu"),
    26: ("Limbang", "13", "minimum_suhu"),
    27: ("Limbang", "13", "maksimum_suhu"),
    28: ("Miri", "13", "minimum_suhu"),
    29: ("Miri", "13", "maksimum_suhu"),
    30: ("Mulu", "13", "minimum_suhu"),
    31: ("Mulu", "13", "maksimum_suhu"),
    32: ("Sibu", "13", "minimum_suhu"),
    33: ("Sibu", "13", "maksimum_suhu"),
    34: ("Sri Aman", "13", "minimum_suhu"),
    35: ("Sri Aman", "13", "maksimum_suhu"),
}

def populate_jadual_37_0_2(sheet, hierarchy, report_type):
    print("  -> Populating Jadual 37.0 (Purata Suhu) untuk Malaysia")

    sheet["C2"] = ": Purata suhu, Malaysia, 2022 - 2024 (samb.)"
    sheet["C3"] = ": Mean temperature, Malaysia, 2022 - 2024 (cont'd)"

    data_cache = {}

    for row_idx, (station_name, state_code, metric_name) in ROW_MAP.items():
        cache_key = f"{station_name}_{state_code}"

        if cache_key not in data_cache:
            data_cache[cache_key] = get_metrics_dict(location_code=station_name, level='meteorologi', parent_code=state_code)
            if not data_cache[cache_key]:
                print(f"     [⚠️] No data at all for station: '{station_name}' in state {state_code}.")
            else:
                years = list(data_cache[cache_key].keys())
                print(f"     [✅] Station '{station_name}' has years: {years}")

        station_data = data_cache[cache_key]

        for col_idx, year in COL_MAP.items():
            year_data = station_data.get(str(year), {})

            if station_name == "Keningau" and year == "2022":
                print(f"   [DEBUG] Keys in year_data for Keningau 2022: {list(year_data.keys())}")

            raw_val = "n.a"
            possible_names = [
                metric_name,                                 
                metric_name.replace("minimum", "min").replace("maksimum", "max"),
                metric_name.replace("minimum", "rendah").replace("maksimum", "tinggi"),
                f"suhu_{metric_name.split('_')[-1]}",        
                f"{metric_name.split('_')[-1]}_suhu",        
                metric_name.replace("_suhu", ""),            
                "min_temp" if "minimum" in metric_name else "max_temp",
                "suhu_minimum" if "minimum" in metric_name else "suhu_maksimum",
            ]

            for alt in possible_names:
                if alt in year_data:
                    raw_val = year_data[alt]
                    if alt != metric_name:
                        print(f"     [INFO] Using alternative metric '{alt}' for {station_name} in {year}")
                    break

            if pd.notna(raw_val) and str(raw_val).strip() not in ["", "n.a", "n.a.", "nan", "NaN"]:
                try:
                    val = float(raw_val)
                except (ValueError, TypeError):
                    val = raw_val
            else:
                val = "n.a"

            safe_write(sheet, row_idx, col_idx, val)