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
    # PAHANG (State Code "06")
    7:  ("Kuantan", "06", "minimum_suhu"),
    8:  ("Kuantan", "06", "maksimum_suhu"),
    9:  ("Muadzam Shah", "06", "minimum_suhu"),
    10: ("Muadzam Shah", "06", "maksimum_suhu"),
    11: ("Temerloh", "06", "minimum_suhu"),
    12: ("Temerloh", "06", "maksimum_suhu"),
    
    # P.PINANG (State Code "07")
    14: ("Bayan Lepas", "07", "minimum_suhu"),
    15: ("Bayan Lepas", "07", "maksimum_suhu"),
    16: ("Butterworth", "07", "minimum_suhu"),
    17: ("Butterworth", "07", "maksimum_suhu"),
    
    # PERAK (State Code "08")
    19: ("Ipoh", "08", "minimum_suhu"),
    20: ("Ipoh", "08", "maksimum_suhu"),
    21: ("Lubok Merbau, Kuala Kangsar", "08", "minimum_suhu"),
    22: ("Lubok Merbau, Kuala Kangsar", "08", "maksimum_suhu"),
    23: ("Sitiawan", "08", "minimum_suhu"),
    24: ("Sitiawan", "08", "maksimum_suhu"),

    # PERLIS (State Code "09")
    26: ("Chuping", "09", "minimum_suhu"),
    27: ("Chuping", "09", "maksimum_suhu"),

    # SELANGOR (State Code "10")
    29: ("Petaling Jaya", "10", "minimum_suhu"),
    30: ("Petaling Jaya", "10", "maksimum_suhu"),
    31: ("Subang", "10", "minimum_suhu"),
    32: ("Subang", "10", "maksimum_suhu"),
    33: ("KLIA Sepang", "10", "minimum_suhu"),
    34: ("KLIA Sepang", "10", "maksimum_suhu"),

    # TERENGGANU (State Code "11")
    36: ("Kerteh", "11", "minimum_suhu"),
    37: ("Kerteh", "11", "maksimum_suhu"),
    38: ("Kuala Terengganu", "11", "minimum_suhu"),
    39: ("Kuala Terengganu", "11", "maksimum_suhu"),
}

# ==========================================
# 2. REPORT INJECTION ENGINE
# ==========================================
def populate_jadual_37_0_1(sheet, hierarchy, report_type):
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