import pandas as pd
from src.data_provider import get_metrics_dict

# ==========================================
# 1. MAPPING CONFIGURATION FOR JADUAL 31.1
# ==========================================

# Map Excel Column Index to the Metric Name in the database
# E=5, F=6, G=7, H=8, I=9, J=10
COL_MAP = {
    5: "jumlah_jenayah_harta_benda",
    6: "pecah_rumah_dan_curi",
    7: "kecurian_lori_van",
    8: "kecurian_kereta",
    9: "kecurian_motosikal_skuter",
    10: "kecurian_lain"
}

# Map Excel Row Index to a tuple of ("Location_Name", "Year", "Level")
# Based on the screenshot rows
ROW_MAP = {
    # MALAYSIA
    11: ("Malaysia", "2022", "malaysia"),
    12: ("Malaysia", "2023", "malaysia"),
    13: ("Malaysia", "2024", "malaysia"),
    
    # JOHOR
    15: ("Johor", "2022", "negeri"),
    16: ("Johor", "2023", "negeri"),
    17: ("Johor", "2024", "negeri"),
    
    # KEDAH
    19: ("Kedah", "2022", "negeri"),
    20: ("Kedah", "2023", "negeri"),
    21: ("Kedah", "2024", "negeri"),
    
    # KELANTAN
    23: ("Kelantan", "2022", "negeri"),
    24: ("Kelantan", "2023", "negeri"),
    25: ("Kelantan", "2024", "negeri"),
    
    # MELAKA
    27: ("Melaka", "2022", "negeri"),
    28: ("Melaka", "2023", "negeri"),
    29: ("Melaka", "2024", "negeri"), # Assuming sequence continues
    
    # ... continue mapping the remaining states down the sheet
}

# ==========================================
# 2. REPORT INJECTION ENGINE
# ==========================================
def populate_jadual_31_1(sheet, hierarchy, report_type):
    print("  -> Populating Jadual 31.1 (Jenayah Harta Benda) untuk Malaysia & Negeri")

    # 1. Static Title Injection
    title_bm = ": Jenayah harta benda mengikut negeri dan jenis jenayah, Malaysia, 2022 - 2024"
    title_en = ": Property crime by state and type of crime, Malaysia, 2022 - 2024"
    sheet.range("C3").value = title_bm
    sheet.range("C4").value = title_en

    # 2. Cache dictionary to prevent querying the database multiple times for the same location
    data_cache = {}

    # 3. Inject Data
    for row_idx, (location, year, level) in ROW_MAP.items():
        
        # Fetch data only if we haven't fetched this location yet
        if location not in data_cache:
            data_cache[location] = get_metrics_dict(location, level=level)
            
            if not data_cache[location]:
                print(f"     [Warning] No data found for {location}.")
        
        # Extract the dictionary for the specific year
        location_data = data_cache[location]
        year_data = location_data.get(year, {})
        
        # Loop through columns and inject the appropriate metric
        for col_idx, metric_name in COL_MAP.items():
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