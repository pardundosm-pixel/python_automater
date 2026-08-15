import pandas as pd
from src.data_provider import get_metrics_dict

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
    # SABAH (State Code "12")
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

    # SARAWAK (State Code "13")
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

# ==========================================
# 2. REPORT INJECTION ENGINE
# ==========================================
def populate_jadual_37_0_2(sheet, hierarchy, report_type):
    print("  -> Populating Jadual 37.0 (Purata Suhu) untuk Malaysia")

    # 1. Static Title Injection
    title_bm = ": Purata suhu, Malaysia, 2022 - 2024"
    title_en = ": Mean temperature, Malaysia, 2022 - 2024"
    sheet.range("C2").value = title_bm
    sheet.range("C3").value = title_en

    # 2. Cache dictionary to prevent querying the database repeatedly
    data_cache = {}

    # 3. Inject Data
    for row_idx, (station_name, state_code, metric_name) in ROW_MAP.items():
        
        # Create a unique cache key for each station + state combo
        cache_key = f"{station_name}_{state_code}"
        
        # Fetch data only if we haven't fetched this station yet
        if cache_key not in data_cache:
            # We pass the station name as the location_code, and state_code as the parent_code
            data_cache[cache_key] = get_metrics_dict(location_code=station_name, level='meteorologi', parent_code=state_code)
            
            if not data_cache[cache_key]:
                print(f"     [Warning] No data found for station: {station_name} in state {state_code}.")
        
        # Extract the dictionary for the specific station
        station_data = data_cache[cache_key]
        
        # Loop through columns (Years)
        for col_idx, year in COL_MAP.items():
            year_data = station_data.get(str(year), {})
            val = year_data.get(metric_name, "n.a")
            
            # Clean and parse missing values
            if pd.notna(val) and val != "n.a" and val != "":
                try: 
                    val = float(val)
                except (ValueError, TypeError): 
                    pass
            else:
                val = "n.a"
                
            sheet.range((row_idx, col_idx)).value = val