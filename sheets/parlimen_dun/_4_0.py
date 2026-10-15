import pandas as pd
from src.data_provider import get_metrics_dict

# ==========================================
# MAPPING CONFIGURATION 
# ==========================================
# 1. Map Excel Row to Metric Name
ROW_MAP = {
    11: "tenaga_buruh",
    15: "penduduk_bekerja",
    19: "penganggur",
    23: "luar_tenaga_buruh",
    27: "kadar_penyertaan_tenaga_buruh",
    31: "kadar_pengangguran",

}

# 2. Map Excel Column to Year
COL_MAP = {
    11: "2024",  # E.g., Column D
}

# ==========================================
# 2. REPORT INJECTION ENGINE
# ==========================================
def populate_jadual_4(sheet, hierarchy, report_type):
    # Safely get the location name for the debug print
    loc_name_debug = hierarchy.get('parl_code') or hierarchy.get('dun_code')
    print(f"  -> Populating Jadual 4.0 (Guna Tenaga) for {loc_name_debug}")

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
    title_bm = f": Statistik guna tenaga, Parlimen {parl_name}, {hierarchy.get('state_name')}, 2024"
    title_en = f": Statistics of employment, Parliament {parl_name}, {hierarchy.get('state_name')}, 2024"

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