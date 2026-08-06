import pandas as pd
from src.data_provider import get_metrics_dict

# ==========================================
# 1. MAPPING CONFIGURATION FOR JADUAL 43.0 (NEGERI)
# ==========================================
# Map the EXCEL ROW NUMBER to the METRIC NAME in the database
ROW_MAP = {
    # Nilai (Value) - RM juta
    8:  "eksport_nilai",
    11: "import_nilai",
    14: "jumlah_dagangan_nilai",
    17: "imbangan_dagangan_nilai",
    
    # Pertumbuhan Tahunan (Annual Growth) - %
    20: "eksport_peratus",
    23: "import_peratus",
    26: "jumlah_dagangan_peratus"
}

# Map the EXCEL COLUMN NUMBER to the YEAR STRING
# Based on screenshot: Column G = 7, Column H = 8, Column I = 9
COL_MAP = {
    7: "2022",  
    8: "2023",  
    9: "2024"   
}

# ==========================================
# 2. REPORT INJECTION ENGINE
# ==========================================
def populate_jadual_43(sheet, hierarchy, report_type):
    state_name = hierarchy.get('state_name', 'Unknown State')
    state_code = hierarchy.get('state_code')
    print(f"  -> Populating Jadual 43.0 (Perdagangan) untuk {state_name}")

    # 1. Fetch the Data Payload for the specific Negeri
    metrics_data = get_metrics_dict(state_code, level='negeri')
    
    if not metrics_data:
        print(f"     [Warning] No data found for {state_name}.")
        return

    # ==========================================
    # DYNAMIC TABLE TITLE MODIFICATION
    # ==========================================
    title_bm = f": Eksport, import, jumlah dagangan dan imbangan dagangan, {state_name} (RM juta), 2022 - 2024p"
    title_en = f": Exports, imports, total trade and balance of trade, {state_name} (RM million), 2022 - 2024p"

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