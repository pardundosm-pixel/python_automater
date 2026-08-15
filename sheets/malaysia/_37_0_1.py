import pandas as pd
from src.data_provider import get_metrics_dict

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