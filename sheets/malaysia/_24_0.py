import pandas as pd
from src.data_provider import get_metrics_dict

# ==========================================
# MAPPING CONFIGURATION FOR JADUAL 24.0 (IHPR)
# ==========================================
# TODO: Map the EXCEL ROW NUMBER to the METRIC NAME in the database
# Example: 7: "jumlah_penduduk", 11: "warganegara"
ROW_MAP = {
    # Indeks Harga Pengeluar (IHPR)
    12: "indeks_harga_pengeluar_jumlah",                                 # Jumlah / Total
    15: "indeks_harga_pengeluar_pertanian_perhutanan_dan_perikanan",    # Pertanian, perhutanan dan perikanan
    18: "indeks_harga_pengeluar_perlombongan",                          # Perlombongan
    21: "indeks_harga_pengeluar_pembuatan",                             # Pembuatan
    24: "indeks_harga_pengeluar_bekalan_elektrik_dan_gas",              # Bekalan elektrik dan gas
    27: "indeks_harga_pengeluar_bekalan_air",                           # Bekalan air

    # Perubahan Peratus Tahunan (Annual % Change)
    34: "perubahan_peratus_tahunan_jumlah",                             # Jumlah / Total
    37: "perubahan_peratus_tahunan_pertanian_perhutanan_dan_perikanan", # Pertanian...
    40: "perubahan_peratus_tahunan_perlombongan",                       # Perlombongan
    43: "perubahan_peratus_tahunan_pembuatan",                          # Pembuatan
    46: "perubahan_peratus_tahunan_bekalan_elektrik_dan_gas",           # Bekalan elektrik dan gas
    49: "perubahan_peratus_tahunan_bekalan_air"                         # Bekalan air
}

# TODO: Map the EXCEL COLUMN NUMBER to the YEAR STRING
COL_MAP = {
    5  : "2023",  
    6  : "2024",  
    7  : "2025"   
}

def populate_jadual_24_0(sheet, hierarchy, report_type):
    print(f"  -> Populating Jadual 24.0 (IHPR) untuk Malaysia_24_0")
    
    # 1. Fetch the Data Payload strictly for Malaysia
    metrics_data = get_metrics_dict("Malaysia", level='malaysia')
    
    if not metrics_data:
            print(f"     [Warning] No data found for Malaysia.")
            return

    # ==========================================
        # DYNAMIC TABLE TITLE MODIFICATION
    # ==========================================
    title_bm  = ": Indeks Harga Pengeluar (IHPR) Pengeluaran Tempatan dan Perubahan Peratus Tahunan mengikut "
    title_bm2 = "  Sektor, Malaysia, 2023 - 2025 "
    title_en  = ": Average price for selected items, Malaysia, 2023 - 2025"
    title_en2 = "  2023 - 2025"
        
    # Set the exact cells where your title sits in the template
    # Targeting Column C based on standard template behavior
    sheet.range("C1").value = title_bm
    sheet.range("C2").value = title_bm2
    sheet.range("C3").value = title_en
    sheet.range("C4").value = title_en2

    # Standard Injection Loop
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