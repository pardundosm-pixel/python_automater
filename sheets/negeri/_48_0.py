import pandas as pd
from src.data_provider import get_metrics_dict

# ==========================================
# MAPPING CONFIGURATION FOR JADUAL 48.0 (NEGERI)
# ==========================================
# Map the EXCEL ROW NUMBER to the METRIC NAME in the database
ROW_MAP = {
    # Sekolah kerajaan & bantuan kerajaan / Government school & government assistance
    10: "bilangan_murid_prasekolah_kerajaan",
    13: "bilangan_murid_rendah_kerajaan",
    16: "bilangan_murid_menengah_rendah_kerajaan",
    19: "bilangan_murid_menengah_atas_kerajaan",
    22: "bilangan_murid_khas_rendah_kerajaan",
    25: "bilangan_murid_khas_menengah_kerajaan",

    # Tadika, sekolah rendah dan menengah swasta / Kindergarten, private primary and secondary schools
    30: "bilangan_murid_swasta_tadika",
    33: "bilangan_murid_swasta_rendah",
    36: "bilangan_murid_swasta_menengah",
    39: "bilangan_murid_swasta_khas",
    42: "bilangan_murid_swasta_antarabangsa",
    45: "bilangan_murid_swasta_ekspatriat"
}

# Map the EXCEL COLUMN NUMBER to the YEAR STRING
# Based on screenshot: Column J = 10, Column K = 11, Column L = 12
COL_MAP = {
    6: "2024",  
    7: "2025"   
}

# ==========================================
# 2. REPORT INJECTION ENGINE
# ==========================================
def populate_jadual_48_0(sheet, hierarchy, report_type):
    state_name = hierarchy.get('state_name', 'Unknown State')
    state_code = hierarchy.get('state_code')
    print(f"  -> Populating Jadual 48.0 (Bilangan Murid) untuk {state_name}")

    # 1. Fetch the Data Payload for the specific Negeri
    metrics_data = get_metrics_dict(state_code, level='negeri')
    
    if not metrics_data:
        print(f"     [Warning] No data found for {state_name}.")
        return

    # ==========================================
    # DYNAMIC TABLE TITLE MODIFICATION
    # ==========================================
    title_bm = f": Bilangan murid pelbagai peringkat dan jenis sekolah, {state_name}, 2024 - 2025"
    title_en = f":  Number of pupils of various level and types of school, {state_name}, 2024 - 2025"

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