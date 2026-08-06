import pandas as pd
from src.data_provider import get_metrics_dict

# ==========================================
# 1. MAPPING CONFIGURATION 
# ==========================================
# Map the EXCEL ROW NUMBER to the METRIC NAME in the database
ROW_MAP = {
    8:  "jumlah_penduduk",
    9:  "lelaki",
    10: "perempuan",
    12: "warganegara",
    13: "bukan_warganegara",
    15: "bumiputera",
    16: "cina",
    17: "india",
    18: "lain_lain",
    20: "umur_0_14",
    21: "umur_15_64",
    22: "umur_65_ke_atas"
}

# Map the EXCEL COLUMN INDEX to the YEAR STRING
COL_MAP = {
    6: "2023",
    7: "2024",
}

# ==========================================
# 2. REPORT INJECTION ENGINE
# ==========================================
def populate_jadual_2_2(sheet, hierarchy, report_type):
    # Safely get the location name for the debug print
    loc_name_debug = hierarchy.get('parl_code') or hierarchy.get('dun_code')
    print(f"  -> Populating Jadual 2.2 (Parlimen) for {loc_name_debug}")

    # 1. Fetch the Data Payload
    # We use `.get() or .get()` so it works flawlessly for BOTH Parlimen and DUN reports
    target_code = hierarchy.get('parl_code') or hierarchy.get('parent_parl_code')
    metrics_data = get_metrics_dict(target_code, level='parlimen')
    
    if not metrics_data:
        print(f"     [Warning] No data found for Parliament {target_code}.")
        return

    # ==========================================
    # DYNAMIC TABLE TITLE MODIFICATION
    # ==========================================
    # Safely grab the parliament name for the title
    parl_name = hierarchy.get('parl_name') or hierarchy.get('parent_parl_name')
    
    # Because this is Jadual 2.2, the title ALWAYS says "Parlimen", even in a DUN report
    title_bm = f": Keluasan dan penduduk bagi Parlimen {parl_name}, {hierarchy.get('state_name')}, 2022 - 2024"
    title_en = f": Area and population for Parliament {parl_name}, {hierarchy.get('state_name')}, 2022 - 2024"

    sheet.range("C3").value = title_bm
    sheet.range("C4").value = title_en
    # ==========================================

    # 2. Inject Data
    for col_idx, year in COL_MAP.items():
        year_data = metrics_data.get(str(year), {})
        
        for row_idx, metric_name in ROW_MAP.items():
            val = year_data.get(metric_name, "n.a")
            
            if pd.notna(val) and val != "n.a" and val != "":
                try: 
                    val = float(val)
                except (ValueError, TypeError): 
                    pass
            else:
                val = "n.a"
                
            sheet.range((row_idx, col_idx)).value = val