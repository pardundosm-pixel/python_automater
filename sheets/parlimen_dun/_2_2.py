import pandas as pd
from src.data_provider import get_metrics_dict

# ==========================================
# 1. MAPPING CONFIGURATION 
# ==========================================
# Map the EXCEL ROW NUMBER to the METRIC NAME in the database
ROW_MAP = {
    8:  "jumlah_penduduk",
    12: "penduduk_warganegara",
    13: "penduduk_bukan_warganegara",
    15: "penduduk_lelaki",
    16: "penduduk_perempuan",
    19: "peratus_penduduk_warganegara",
    20: "peratus_penduduk_bukan_warganegara",
    24: "peratus_penduduk_bumiputera",
    25: "peratus_penduduk_cina",
    26: "peratus_penduduk_india",
    27: "peratus_penduduk_lain_lain",
    31: "penduduk_umur_0_14",
    33: "penduduk_umur_15_30",
    35: "penduduk_umur_15_64",
    37: "penduduk_umur_65_lebih",
    39: "penduduk_umur_18_lebih",
    44: "jumlah_nisbah_tanggungan",
    45: "umur_muda",
    46: "umur_tua",
    48: "nisbah_jantina",
    51: "kepadatan_penduduk"
}

# Map the EXCEL COLUMN INDEX to the YEAR STRING
COL_MAP = {
    6: "2024",
    7: "2025",
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
    title_bm = f": Anggaran penduduk pertengahan tahun, Parlimen {parl_name}, {hierarchy.get('state_name')}, 2024 - 2025"
    title_en = f": Mid-year population estimates, Parliament of {parl_name}, {hierarchy.get('state_name')}, 2024 - 2025"

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