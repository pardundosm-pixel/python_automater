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
    # TERENGGANU (State Code "11")
    7: ("Kerteh", "11", "jumlah_volum_hujan"),
    8: ("Kerteh", "11", "bilangan_hari_hujan"),
    9: ("Kuala Terengganu", "11", "jumlah_volum_hujan"),
    10: ("Kuala Terengganu", "11", "bilangan_hari_hujan"),


    # SABAH (State Code "12")
    12:  ("Keningau", "12", "jumlah_volum_hujan"),
    13:  ("Keningau", "12", "bilangan_hari_hujan"),
    14:  ("Kota Kinabalu", "12", "jumlah_volum_hujan"),
    15: ("Kota Kinabalu", "12", "bilangan_hari_hujan"),
    16: ("Kudat", "12", "jumlah_volum_hujan"),
    17: ("Kudat", "12", "bilangan_hari_hujan"),
    
    18: ("Ranau", "12", "jumlah_volum_hujan"),
    19: ("Ranau", "12", "bilangan_hari_hujan"),
    20: ("Sandakan", "12", "jumlah_volum_hujan"),
    21: ("Sandakan", "12", "bilangan_hari_hujan"),
    
    22: ("Tawau", "12", "jumlah_volum_hujan"),
    23: ("Tawau", "12", "bilangan_hari_hujan"),

    # SARAWAK (State Code "13")
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
    36: ("Mulu", "13", "bilangan_hari_hujan")
}

# ==========================================
# 2. REPORT INJECTION ENGINE
# ==========================================
def populate_jadual_38_0_2(sheet, hierarchy, report_type):
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