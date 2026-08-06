import pandas as pd
from src.data_provider import get_metrics_dict

# ==========================================
# MAPPING CONFIGURATION FOR JADUAL 14.0 (MALAYSIA)
# ==========================================
# Map the EXCEL ROW NUMBER to the METRIC NAME in the database
ROW_MAP = {
    # KDNK mengikut Jenis Aktiviti Ekonomi pada Harga Semasa (RM juta)
    10: "kdnk_harga_pembeli_nilai",
    11: "pertanian_nilai",
    12: "perlombongan_pengkuarian_nilai",
    13: "pembuatan_nilai",
    14: "pembinaan_nilai",
    15: "perkhidmatan_nilai",
    16: "duti_import_nilai",
    
    # KDNK mengikut Jenis Aktiviti Ekonomi pada Harga Semasa - Perubahan Peratusan Tahunan (%)
    21: "kdnk_harga_pembeli_peratus",
    22: "pertanian_peratus",
    23: "perlombongan_pengkuarian_peratus",
    24: "pembuatan_peratus",
    25: "pembinaan_peratus",       # Assuming row sequence continues
    26: "perkhidmatan_peratus",    # Assuming row sequence continues
    27: "duti_import_peratus"      # Assuming row sequence continues
}

# Map the EXCEL COLUMN NUMBER to the YEAR STRING
# Based on screenshot: Column E (2022) = 5, Column F (2023) = 6, Column G (2024) = 7
COL_MAP = {
    5: "2022",  
    6: "2023",  
    7: "2024"   
}

# ==========================================
# 2. REPORT INJECTION ENGINE
# ==========================================
def populate_jadual_14(sheet, hierarchy, report_type):
    print(f"  -> Populating Jadual 14.0 (KDNK) untuk Malaysia")

    # 1. Fetch the Data Payload strictly for Malaysia
    metrics_data = get_metrics_dict("Malaysia", level='malaysia')
    
    if not metrics_data:
        print(f"     [Warning] No data found for Malaysia.")
        return

    # ==========================================
    # DYNAMIC TABLE TITLE MODIFICATION
    # ==========================================
    title_bm = ": Keluaran Dalam Negeri Kasar (KDNK), Malaysia, 2022 - 2024"
    title_en = ": Gross Domestic Product (GDP), Malaysia, 2022 - 2024"

    # Set the exact cells where your title sits in the template
    # Targeting Column C based on standard template behavior
    sheet.range("C3").value = title_bm
    sheet.range("C4").value = title_en
    # ==========================================

    # 2. Inject Data
    # Loop over the columns (Years) first
    for col_idx, year in COL_MAP.items():
        year_data = metrics_data.get(str(year), {})
        
        for row_idx, metric_name in ROW_MAP.items():
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