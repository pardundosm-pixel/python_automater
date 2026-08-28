import pandas as pd
from src.data_provider import get_metrics_dict
from src.excel_utils import safe_write

# ==========================================
# 1. MAPPING CONFIGURATION FOR JADUAL 37.0
# ==========================================

# Map Excel Column Index to the Year
COL_MAP = {
    7: "2022", # Column G
    8: "2023", # Column H
    9: "2024" # Column I
}

# Map Excel Row Index to a tuple of ("Stesen Name", "State Code", "Metric Name")
# The State Code prevents collisions (e.g., Johor is "01", Kedah is "02")
ROW_MAP = {
    # LABUAN (State Code "15")
    7:  ("Labuan", "15", "minimum_suhu"),
    8:  ("Labuan", "15", "maksimum_suhu"),
}

# ==========================================
# 2. REPORT INJECTION ENGINE
# ==========================================
def populate_jadual_37_0_3(sheet, hierarchy, report_type):
    print("  -> Populating Jadual 37.0 (Purata Suhu) untuk Malaysia")

    sheet["C2"] = ": Purata suhu, Malaysia, 2022 - 2024 (samb.)"
    sheet["C3"] = ": Mean temperature, Malaysia, 2022 - 2024 (cont'd)"

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