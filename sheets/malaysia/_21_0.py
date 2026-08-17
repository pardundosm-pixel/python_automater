import pandas as pd
from src.data_provider import get_metrics_dict

# ==========================================
# MAPPING CONFIGURATION FOR JADUAL 21.0 (Pelancongan Domestik)
# ==========================================
# TODO: Map the EXCEL ROW NUMBER to the METRIC NAME in the database
# Example: 7: "jumlah_penduduk", 11: "warganegara"
ROW_MAP = {
    # Jumlah terimaan (RM juta)
    8 :  "jumlah_terimaan",
    10:  "jumlah_pelawat_domestik",
    11:  "isi_rumah_yang_dilawati",
    
    # Jumlah Pelawat ('000)
    13:  "jumlah_pelawat",
    15:  "pelawat_harian",
    16:  "pelancong",
    
    # Jumlah perjalanan Pelancongan ('000)
    18:  "jumlah_perjalanan_pelancongan"
    
}

# TODO: Map the EXCEL COLUMN NUMBER to the YEAR STRING
COL_MAP = {
    5 : "2023",  
    6 : "2024",  
    7 : "2025"   
}

def populate_jadual_21_0(sheet, hierarchy, report_type):
    print(f"  -> Populating Jadual 21.0 (Pelancongan Domestik) untuk Malaysia_21_0")
    
    # 1. Fetch the Data Payload strictly for Malaysia
    metrics_data = get_metrics_dict("Malaysia", level='malaysia')
    
    if not metrics_data:
            print(f"     [Warning] No data found for Malaysia.")
            return

    # ==========================================
        # DYNAMIC TABLE TITLE MODIFICATION
    # ==========================================
    title_bm = ": Statistik utama pelancongan domestik, Malaysia, 2023 - 2025"
    title_en = ": Principal statistics of domestic tourism, Malaysia, 2023 - 2025"
        
    # Set the exact cells where your title sits in the template
    # Targeting Column C based on standard template behavior
    sheet.range("C3").value = title_bm
    sheet.range("C4").value = title_en

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