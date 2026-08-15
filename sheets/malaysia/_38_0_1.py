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
    7: ("Batu Embun, Jerantut", "06", "jumlah_volum_hujan"),
    8: ("Batu Embun, Jerantut", "06", "bilangan_hari_hujan"),
    9:  ("Kuantan", "06", "jumlah_volum_hujan"),
    10:  ("Kuantan", "06", "bilangan_hari_hujan"),
    11:  ("Muadzam Shah", "06", "jumlah_volum_hujan"),
    12: ("Muadzam Shah", "06", "bilangan_hari_hujan"),
    13: ("Temerloh", "06", "jumlah_volum_hujan"),
    14: ("Temerloh", "06", "bilangan_hari_hujan"),
    
    # P.PINANG (State Code "07")
    16: ("Bayan Lepas", "07", "jumlah_volum_hujan"),
    17: ("Bayan Lepas", "07", "bilangan_hari_hujan"),
    18: ("Butterworth", "07", "jumlah_volum_hujan"),
    19: ("Butterworth", "07", "bilangan_hari_hujan"),
    
    # PERAK (State Code "08")
    21: ("Ipoh", "08", "jumlah_volum_hujan"),
    22: ("Ipoh", "08", "bilangan_hari_hujan"),
    23: ("Lubok Merbau, Kuala Kangsar", "08", "jumlah_volum_hujan"),
    24: ("Lubok Merbau, Kuala Kangsar", "08", "bilangan_hari_hujan"),
    25: ("Sitiawan", "08", "jumlah_volum_hujan"),
    26: ("Sitiawan", "08", "bilangan_hari_hujan"),

    # PERLIS (State Code "09")
    28: ("Chuping", "09", "jumlah_volum_hujan"),
    29: ("Chuping", "09", "bilangan_hari_hujan"),

    # SELANGOR (State Code "10")
    31: ("Petaling Jaya", "10", "jumlah_volum_hujan"),
    32: ("Petaling Jaya", "10", "bilangan_hari_hujan"),
    33: ("Subang", "10", "jumlah_volum_hujan"),
    34: ("Subang", "10", "bilangan_hari_hujan"),
    35: ("KLIA Sepang", "10", "jumlah_volum_hujan"),
    36: ("KLIA Sepang", "10", "bilangan_hari_hujan")
}

# ==========================================
# 2. REPORT INJECTION ENGINE
# ==========================================
def populate_jadual_38_0_1(sheet, hierarchy, report_type):
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