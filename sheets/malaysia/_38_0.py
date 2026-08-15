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
    # JOHOR (State Code "01")
    7:  ("Batu Pahat", "01", "jumlah_volum_hujan"),
    8:  ("Batu Pahat", "01", "bilangan_hari_hujan"),
    9:  ("Kluang", "01", "jumlah_volum_hujan"),
    10: ("Kluang", "01", "bilangan_hari_hujan"),
    11: ("Mersing", "01", "jumlah_volum_hujan"),
    12: ("Mersing", "01", "bilangan_hari_hujan"),
    13: ("Senai", "01", "jumlah_volum_hujan"),
    14: ("Senai", "01", "bilangan_hari_hujan"),
    
    # KEDAH (State Code "02")
    16: ("Alor Setar", "02", "jumlah_volum_hujan"),
    17: ("Alor Setar", "02", "bilangan_hari_hujan"),
    18: ("Pulau Langkawi", "02", "jumlah_volum_hujan"),
    19: ("Pulau Langkawi", "02", "bilangan_hari_hujan"),
    
    # KELANTAN (State Code "03")
    21: ("Kota Bharu", "03", "jumlah_volum_hujan"),
    22: ("Kota Bharu", "03", "bilangan_hari_hujan"),
    23: ("Kuala Krai", "03", "jumlah_volum_hujan"),
    24: ("Kuala Krai", "03", "bilangan_hari_hujan"),
    25: ("Gong Kedak", "03", "jumlah_volum_hujan"),
    26: ("Gong Kedak", "03", "bilangan_hari_hujan"),

    # MELAKA (State Code "04")
    28: ("Melaka", "04", "jumlah_volum_hujan"),
    29: ("Melaka", "04", "bilangan_hari_hujan"),

    # NEGERI SEMBILAN (State Code "05")
    31: ("Kuala Pilah", "05", "jumlah_volum_hujan"),
    32: ("Kuala Pilah", "05", "bilangan_hari_hujan"),

    # PAHANG (State Code "06")
    34: ("Cameron Highlands", "06", "jumlah_volum_hujan"),
    35: ("Cameron Highlands", "06", "bilangan_hari_hujan")
}

# ==========================================
# 2. REPORT INJECTION ENGINE
# ==========================================
def populate_jadual_38(sheet, hierarchy, report_type):
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