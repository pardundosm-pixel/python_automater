import pandas as pd
from src.data_provider import get_metrics_dict

# ==========================================
# 1. MAPPING CONFIGURATION FOR JADUAL 43.0 (NEGERI)
# ==========================================
# Map the EXCEL ROW NUMBER to the METRIC NAME in the database
ROW_MAP = {

    9:  "jumlah_terimaan",
    11: "jumlah_pelawat_domestik",
    12: "isi_rumah_yang_dilawati",
    
    14: "jumlah_pelawat",
    16: "pelawat_harian",
    17: "pelancong",

    19: "jumlah_perjalanan_pelancongan",

    22: "destinasi_1",
    23: "destinasi_2",
    24: "destinasi_3",
    25: "destinasi_4",
    26: "destinasi_5"
}

# Map the EXCEL COLUMN NUMBER to the YEAR STRING
# Based on screenshot: Column G = 7, Column H = 8, Column I = 9
COL_MAP = {
    5: "2023",  
    6: "2024",  
    7: "2025"   
}

# ==========================================
# 2. REPORT INJECTION ENGINE
# ==========================================
def populate_jadual_44_0(sheet, hierarchy, report_type):
    state_name = hierarchy.get('state_name', 'Unknown State')
    state_code = hierarchy.get('state_code')
    print(f"  -> Populating Jadual 44.0 (Pelancongan) untuk {state_name}")

    # 1. Fetch the Data Payload for the specific Negeri
    metrics_data = get_metrics_dict(state_code, level='negeri')
    
    if not metrics_data:
        print(f"     [Warning] No data found for {state_name}.")
        return

    # ==========================================
    # DYNAMIC TABLE TITLE MODIFICATION
    # ==========================================
    title_bm = f": Statistik utama pelancongan domestik, {state_name} (RM juta), 2023 - 2025"
    title_en = f": Principal statistics of domestic tourism,, {state_name} (RM million), 2023 - 2025"

    # Set the exact cells where your title sits in the template (Column C)
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