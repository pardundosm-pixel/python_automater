import pandas as pd
from src.data_provider import get_metrics_dict
from src.excel_utils import safe_write

# ==========================================
# 1. MAPPING CONFIGURATION FOR JADUAL 37.0
# ==========================================

# Map Excel Column Index to the Year
COL_MAP = {
    8: "2022", # Column H
    9: "2023", # Column I
    10: "2024" # Column J
}

# Map Excel Row Index to a tuple of ("Stesen Name", "State Code", "Metric Name")
# The State Code prevents collisions (e.g., Johor is "01", Kedah is "02")
ROW_MAP = {
    # JOHOR (State Code "01")
    7:  ("Batu Pahat", "01", "minimum_suhu"),
    8:  ("Batu Pahat", "01", "maksimum_suhu"),
    9:  ("Kluang", "01", "minimum_suhu"),
    10: ("Kluang", "01", "maksimum_suhu"),
    11: ("Mersing", "01", "minimum_suhu"),
    12: ("Mersing", "01", "maksimum_suhu"),
    13: ("Senai", "01", "minimum_suhu"),
    14: ("Senai", "01", "maksimum_suhu"),
    
    # KEDAH (State Code "02")
    16: ("Alor Setar", "02", "minimum_suhu"),
    17: ("Alor Setar", "02", "maksimum_suhu"),
    18: ("Pulau Langkawi", "02", "minimum_suhu"),
    19: ("Pulau Langkawi", "02", "maksimum_suhu"),
    
    # KELANTAN (State Code "03")
    21: ("Kota Bharu", "03", "minimum_suhu"),
    22: ("Kota Bharu", "03", "maksimum_suhu"),
    23: ("Kuala Krai", "03", "minimum_suhu"),
    24: ("Kuala Krai", "03", "maksimum_suhu"),
    25: ("Gong Kedak", "03", "minimum_suhu"),
    26: ("Gong Kedak", "03", "maksimum_suhu"),

    # MELAKA (State Code "04")
    28: ("Melaka", "04", "minimum_suhu"),
    29: ("Melaka", "04", "maksimum_suhu"),

    # NEGERI SEMBILAN (State Code "05")
    31: ("Kuala Pilah", "05", "minimum_suhu"),
    32: ("Kuala Pilah", "05", "maksimum_suhu"),

    # PAHANG (State Code "06")
    34: ("Cameron Highlands", "06", "minimum_suhu"),
    35: ("Cameron Highlands", "06", "maksimum_suhu"),
    36: ("Batu Embun, Jerantut", "06", "minimum_suhu"),
    37: ("Batu Embun, Jerantut", "06", "maksimum_suhu"),
}

# ==========================================
# 2. REPORT INJECTION ENGINE
# ==========================================
def populate_jadual_37(sheet, hierarchy, report_type):
    print("  -> Populating Jadual 37.0 (Purata Suhu) untuk Malaysia")

    sheet["C2"] = ": Purata suhu, Malaysia, 2022 - 2024"
    sheet["C3"] = ": Mean temperature, Malaysia, 2022 - 2024"

    data_cache = {}

    for row_idx, (station_name, state_code, metric_name) in ROW_MAP.items():
        cache_key = f"{station_name}_{state_code}"
        
        if cache_key not in data_cache:
            data_cache[cache_key] = get_metrics_dict(location_code=station_name, level='meteorologi', parent_code=state_code)
            if not data_cache[cache_key]:
                print(f"     [Warning] No data found for station: {station_name} in state {state_code}.")
        
        station_data = data_cache[cache_key]
        
        for col_idx, year in COL_MAP.items():
            year_data = station_data.get(str(year), {})
            raw_val = year_data.get(metric_name, "n.a")
            
            if pd.notna(raw_val) and str(raw_val).strip() not in ["", "n.a", "n.a.", "nan", "NaN"]:
                try: 
                    val = float(raw_val)
                except (ValueError, TypeError): 
                    val = raw_val
            else:
                val = "n.a"
                
            safe_write(sheet, row_idx, col_idx, val)