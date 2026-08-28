import pandas as pd
from src.data_provider import get_metrics_dict
from src.excel_utils import safe_write

# ==========================================
# 1. MAPPING CONFIGURATION FOR JADUAL 36.1
# ==========================================

# CHANGE THIS to the year you want to report (currently 2023)
REPORT_YEAR = "2024"

# Map Excel Column Index to the Metric Name in the database
# D=4, F=6, H=8, J=10, L=12, N=14, P=16, R=18
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
# Malaysia is stored as "Malaysia" with level "malaysia"
# States use numeric codes with level "negeri"
ROW_MAP = {
    # MALAYSIA (directly from Malaysia table)
    14: ("Malaysia", REPORT_YEAR, "malaysia"),
    
    # NEGERI (State Codes 01 to 16)
    16: ("01", REPORT_YEAR, "negeri"),  # Johor
    18: ("02", REPORT_YEAR, "negeri"),  # Kedah
    20: ("03", REPORT_YEAR, "negeri"),  # Kelantan
    22: ("04", REPORT_YEAR, "negeri"),  # Melaka
    24: ("05", REPORT_YEAR, "negeri"),  # Negeri Sembilan
    26: ("06", REPORT_YEAR, "negeri"),  # Pahang
    28: ("07", REPORT_YEAR, "negeri"),  # Pulau Pinang
    30: ("08", REPORT_YEAR, "negeri"),  # Perak
    32: ("09", REPORT_YEAR, "negeri"),  # Perlis
    34: ("10", REPORT_YEAR, "negeri"),  # Selangor
    36: ("11", REPORT_YEAR, "negeri"),  # Terengganu
    38: ("12", REPORT_YEAR, "negeri"),  # Sabah
    40: ("13", REPORT_YEAR, "negeri"),  # Sarawak
    42: ("14", REPORT_YEAR, "negeri"),  # W.P. Kuala Lumpur
    44: ("15", REPORT_YEAR, "negeri"),  # W.P. Labuan
    46: ("16", REPORT_YEAR, "negeri"),  # W.P. Putrajaya
}

# ==========================================
# 2. REPORT INJECTION ENGINE
# ==========================================
def populate_jadual_36_1(sheet, hierarchy, report_type):
    print("  -> Populating Jadual 36.1 (Statistik Terpilih PCC) untuk Malaysia & Negeri")

    sheet["C3"] = f": Statistik terpilih Penggunaan Per Kapita item pertanian mengikut negeri, Malaysia, {REPORT_YEAR}"
    sheet["C5"] = f": Selected statistics of Per Capita Consumption of agricultural item by state, Malaysia, {REPORT_YEAR}"

    data_cache = {}

    for row_idx, (location, year, level) in ROW_MAP.items():
        if location not in data_cache:
            data_cache[location] = get_metrics_dict(location, level=level)
            if not data_cache[location]:
                print(f"     [Warning] No data found for location: {location} (Level: {level}).")
        
        location_data = data_cache[location]
        year_data = location_data.get(year, {})
        
        for col_idx, metric_name in COL_MAP.items():
            raw_val = year_data.get(metric_name, "n.a")
            
            if pd.notna(raw_val) and str(raw_val).strip() not in ["", "n.a", "n.a.", "nan", "NaN"]:
                try: 
                    val = float(raw_val)
                except (ValueError, TypeError): 
                    val = raw_val
            else:
                val = "n.a"
                
            safe_write(sheet, row_idx, col_idx, val)