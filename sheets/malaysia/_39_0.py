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
    # JOHOR (State Code "01")
    8:  ("Batu Pahat", "01", "purata_kelembapan_relatif"),
    9:  ("Kluang", "01", "purata_kelembapan_relatif"),
    10: ("Mersing", "01", "purata_kelembapan_relatif"),
    11: ("Senai", "01", "purata_kelembapan_relatif"),
    
    # KEDAH (State Code "02")
    13: ("Alor Setar", "02", "purata_kelembapan_relatif"),
    14: ("Pulau Langkawi", "02", "purata_kelembapan_relatif"),
    
    # KELANTAN (State Code "03")
    16: ("Kota Bharu", "03", "purata_kelembapan_relatif"),
    17: ("Kuala Krai", "03", "purata_kelembapan_relatif"),
    18: ("Gong Kedak", "03", "purata_kelembapan_relatif"),
    
    # MELAKA (State Code "04")
    20: ("Melaka", "04", "purata_kelembapan_relatif"),
    
    # NEGERI SEMBILAN (State Code "05")
    22: ("Kuala Pilah", "05", "purata_kelembapan_relatif"),
    
    # PAHANG (State Code "06")
    24: ("Cameron Highlands", "06", "purata_kelembapan_relatif"),
    25: ("Batu Embun, Jerantut", "06", "purata_kelembapan_relatif"),
    26: ("Kuantan", "06", "purata_kelembapan_relatif"),
    27: ("Muadzam Shah", "06", "purata_kelembapan_relatif"),
    28: ("Temerloh", "06", "purata_kelembapan_relatif"),
    
    # P.PINANG (State Code "07")
    30: ("Bayan Lepas", "07", "purata_kelembapan_relatif"),
    31: ("Butterworth", "07", "purata_kelembapan_relatif"),

    
    # PERAK (State Code "08")
    33: ("Ipoh", "08", "purata_kelembapan_relatif"),
    34: ("Lubok Merbau, Kuala Kangsar", "08", "purata_kelembapan_relatif"),
    35: ("Sitiawan", "08", "purata_kelembapan_relatif"),

    # PERLIS (State Code "09")
    37: ("Chuping", "09", "purata_kelembapan_relatif"),

    # SELANGOR (State Code "10")
    39: ("Petaling Jaya", "10", "purata_kelembapan_relatif"),
    40: ("Subang", "10", "purata_kelembapan_relatif"),
    41: ("KLIA Sepang", "10", "purata_kelembapan_relatif")
}

# ==========================================
# 2. REPORT INJECTION ENGINE
# ==========================================
def populate_jadual_39(sheet, hierarchy, report_type):
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