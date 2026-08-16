import pandas as pd
from src.data_provider import get_metrics_dict

# ==========================================
# 1. MAPPING CONFIGURATION FOR JADUAL 36.1
# ==========================================

# Map Excel Column Index to the Metric Name in the database
# D=4, F=6, H=8, J=10, L=12, N=14, P=16, R=18
# IMPORTANT: Adjust these metric strings if they differ in your fact_metrics_negeri table.
COL_MAP = {
    4: "durian",
    6: "mangga",
    8: "nanas",
    10: "cili",
    12: "kobis_bulat",
    14: "timun",
    16: "tomato",
    18: "sawi"
}

# Map Excel Row Index to a tuple of ("Location_Code", "Year", "Level")
# Note: level='negeri' requires the zero-padded state code, not the string name.
ROW_MAP = {
    # MALAYSIA (Root Level)
    14: ("Malaysia", "2023", "malaysia"),
    
    # NEGERI (State Codes 01 to 14)
    16: ("01", "2023", "negeri"),  # Johor
    18: ("02", "2023", "negeri"),  # Kedah
    20: ("03", "2023", "negeri"),  # Kelantan
    22: ("04", "2023", "negeri"),  # Melaka
    24: ("05", "2023", "negeri"),  # Negeri Sembilan
    26: ("06", "2023", "negeri"),  # Pahang
    28: ("07", "2023", "negeri"),  # Pulau Pinang
    30: ("08", "2023", "negeri"),  # Perak
    32: ("09", "2023", "negeri"),  # Perlis
    34: ("10", "2023", "negeri"),  # Selangor
    36: ("11", "2023", "negeri"),  # Terengganu
    38: ("12", "2023", "negeri"),  # Sabah
    40: ("13", "2023", "negeri"),  # Sarawak
    42: ("14", "2023", "negeri"),  # W.P. Kuala Lumpur
    44: ("15", "2023", "negeri"),  # W.P. Labuan
    46: ("16", "2023", "negeri"),  # W.P. Putrajaya
}

# ==========================================
# 2. REPORT INJECTION ENGINE
# ==========================================
def populate_jadual_36_1(sheet, hierarchy, report_type):
    print("  -> Populating Jadual 36.1 (Statistik Terpilih PCC) untuk Malaysia & Negeri")

    # 1. Static Title Injection
    title_bm = ": Statistik terpilih Penggunaan Per Kapita item pertanian mengikut negeri, Malaysia, 2024"
    title_en = ": Selected statistics of Per Capita Consumption of agricultural item by state, Malaysia, 2024"
    
    # Target the exact cells where your title sits
    sheet.range("C3").value = title_bm
    sheet.range("C5").value = title_en

    # 2. Cache dictionary to prevent querying the database multiple times for the same location
    data_cache = {}

    # 3. Inject Data
    for row_idx, (location, year, level) in ROW_MAP.items():
        
        # Fetch data only if we haven't fetched this location yet
        if location not in data_cache:
            data_cache[location] = get_metrics_dict(location, level=level)
            
            if not data_cache[location]:
                print(f"     [Warning] No data found for location code: {location} (Level: {level}).")
        
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