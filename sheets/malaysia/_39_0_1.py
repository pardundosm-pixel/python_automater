import pandas as pd
from src.data_provider import get_metrics_dict

# ==========================================
# 1. MAPPING CONFIGURATION FOR JADUAL 39.0
# ==========================================

# Map Excel Column Index to the Year
# F=6, G=7, H=8
COL_MAP = {
    6: "2022", 
    7: "2023", 
    8: "2024" 
}

# Map Excel Row Index to a tuple of ("Stesen Name", "State Code", "Metric Name")
# The State Code prevents collisions (e.g., Johor is "01", Kedah is "02")
# NOTE: Ensure the metric name "purata_kelembapan_relatif" matches your fact_metrics_meteorologi table.
ROW_MAP = {
    # TERENGGANU (State Code "11")
        8: ("Kerteh", "11", "purata_kelembapan_relatif"),
        9: ("Kuala Terengganu", "11", "purata_kelembapan_relatif"),
    
        # SABAH (State Code "12")
        11:  ("Keningau", "12", "purata_kelembapan_relatif"),
        12:  ("Kota Kinabalu", "12", "purata_kelembapan_relatif"),
        13: ("Kudat", "12", "purata_kelembapan_relatif"),
        
        14: ("Ranau", "12", "purata_kelembapan_relatif"),
        15: ("Sandakan", "12", "purata_kelembapan_relatif"),
        
        16: ("Tawau", "12", "purata_kelembapan_relatif"),
    
        # SARAWAK (State Code "13")
        18: ("Bintulu", "13", "purata_kelembapan_relatif"),
        19: ("Kapit", "13", "purata_kelembapan_relatif"),
    
        20: ("Kuching", "13", "purata_kelembapan_relatif"),
    
        21: ("Limbang", "13", "purata_kelembapan_relatif"),
        22: ("Miri", "13", "purata_kelembapan_relatif"),
        23: ("Mulu", "13", "purata_kelembapan_relatif"),

        24: ("Sibu", "13", "purata_kelembapan_relatif"),
        25: ("Sri Aman", "13", "purata_kelembapan_relatif"),
    
        # LABUAN (State Code "15")
        27:  ("Labuan", "15", "purata_kelembapan_relatif"),
}

# ==========================================
# 2. REPORT INJECTION ENGINE
# ==========================================
def populate_jadual_39_0_1(sheet, hierarchy, report_type):
    print("  -> Populating Jadual 39.0 (Purata Kelembapan Relatif) untuk Malaysia")

    # 1. Static Title Injection
    title_bm = ": Purata kelembapan relatif, Malaysia, 2022 - 2024"
    title_en = ": Mean relative humidity, Malaysia, 2022 - 2024"
    
    # Check your template to ensure these are the correct title cells
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