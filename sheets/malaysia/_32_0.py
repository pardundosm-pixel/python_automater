import pandas as pd
from src.data_provider import get_metrics_dict

# ==========================================
# 1. MAPPING CONFIGURATION FOR JADUAL 32.0
# ==========================================

# Map Excel Column Index to the Metric Name in the database
# E=5, G=7, H=8, I=9 (Assuming F is skipped based on the visual layout)
COL_MAP = {
    5: "kemalangan_jalan_raya",
    7: "jumlah_kecederaan_kematian",
    8: "kecederaan",
    9: "kematian"
}

# Map Excel Row Index to a tuple of ("Location_Name", "Year", "Level")
# Based on the screenshot rows (Note the years are 2021, 2022, 2023)
ROW_MAP = {
    # MALAYSIA
    11: ("Malaysia", "2021", "malaysia"),
    12: ("Malaysia", "2022", "malaysia"),
    13: ("Malaysia", "2023", "malaysia"),
    
    # JOHOR
    15: ("Johor", "2021", "negeri"),
    16: ("Johor", "2022", "negeri"),
    17: ("Johor", "2023", "negeri"),
    
    # KEDAH
    19: ("Kedah", "2021", "negeri"),
    20: ("Kedah", "2022", "negeri"),
    21: ("Kedah", "2023", "negeri"),
    
    # KELANTAN
    23: ("Kelantan", "2021", "negeri"),
    24: ("Kelantan", "2022", "negeri"),
    25: ("Kelantan", "2023", "negeri"),
    
    # MELAKA
    27: ("Melaka", "2021", "negeri"),
    28: ("Melaka", "2022", "negeri"),
    29: ("Melaka", "2023", "negeri"), 
    
    # NEGERI SEMBILAN
    31: ("Negeri Sembilan", "2021", "negeri"),
    32: ("Negeri Sembilan", "2022", "negeri"),
    33: ("Negeri Sembilan", "2023", "negeri"), 
    
    # ... continue mapping the remaining states down the sheet
}

# ==========================================
# 2. REPORT INJECTION ENGINE
# ==========================================
def populate_jadual_32(sheet, hierarchy, report_type):
    print("  -> Populating Jadual 32.0 (Kemalangan Jalan Raya) untuk Malaysia & Negeri")

    # 1. Static Title Injection
    title_bm = ": Bilangan kemalangan jalan raya, kecederaan dan kematian yang dilaporkan mengikut negeri, Malaysia, 2021 - 2023"
    title_en = ": Number of road accidents, injuries and deaths reported by state, Malaysia, 2021 - 2023"
    
    # Ensure these cell coordinates match where the title is in your actual file
    sheet.range("C2").value = title_bm 
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