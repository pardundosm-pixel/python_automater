import pandas as pd
from src.data_provider import get_metrics_dict

# ==========================================
# MAPPING CONFIGURATION FOR JADUAL 42.0 (NEGERI)
# ==========================================
# Map the EXCEL ROW NUMBER to the METRIC NAME in the database
ROW_MAP = {
    # KDNK pada harga malar 2015 (RM juta)
    8:  "kdnk_harga_malar_nilai",      # Total KDNK
    9:  "pertanian_nilai",
    10: "perlombongan_pengkuarian_nilai",
    11: "pembuatan_nilai",
    12: "pembinaan_nilai",
    13: "perkhidmatan_nilai",
    14: "duti_import_nilai",
    
    # Perubahan peratusan tahunan (%)
    17: "kdnk_harga_malar_peratus",    # Total KDNK Growth
    18: "pertanian_peratus",
    19: "perlombongan_pengkuarian_peratus",
    20: "pembuatan_peratus",           # Deduced based on standard pattern
    21: "pembinaan_peratus",           # Deduced based on standard pattern
    22: "perkhidmatan_peratus",        # Deduced based on standard pattern
    23: "duti_import_peratus"          # Deduced based on standard pattern
}

# Map the EXCEL COLUMN NUMBER to the YEAR STRING
# Based on screenshot: Column J = 10, Column K = 11, Column L = 12
COL_MAP = {
    10: "2022",  
    11: "2023",  
    12: "2024"   
}

# ==========================================
# 2. REPORT INJECTION ENGINE
# ==========================================
def populate_jadual_42(sheet, hierarchy, report_type):
    state_name = hierarchy.get('state_name', 'Unknown State')
    state_code = hierarchy.get('state_code')
    print(f"  -> Populating Jadual 42.0 (KDNK) untuk {state_name}")

    # 1. Fetch the Data Payload for the specific Negeri
    metrics_data = get_metrics_dict(state_code, level='negeri')
    
    if not metrics_data:
        print(f"     [Warning] No data found for {state_name}.")
        return

    # ==========================================
    # DYNAMIC TABLE TITLE MODIFICATION
    # ==========================================
    title_bm = f": Keluaran Dalam Negeri Kasar (KDNK), {state_name}, 2022 - 2024"
    title_en = f": Gross Domestic Product (GDP), {state_name}, 2022 - 2024"

    # Set the exact cells where your title sits in the template
    sheet.range("C3").value = title_bm
    sheet.range("C4").value = title_en
    # ==========================================

    # 2. Inject Data Flush to the Grid
    for col_idx, year in COL_MAP.items():
        year_data = metrics_data.get(str(year), {})
        
        for row_idx, metric_name in ROW_MAP.items():
            val = year_data.get(metric_name, "n.a")
            
            # Sanitization and missing value fallback
            if pd.notna(val) and val != "n.a" and val != "":
                try: 
                    val = float(val)
                except (ValueError, TypeError): 
                    pass
            else:
                val = "n.a"
                
            sheet.range((row_idx, col_idx)).value = val