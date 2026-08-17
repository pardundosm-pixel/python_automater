import pandas as pd
from src.data_provider import get_metrics_dict

# ==========================================
# MAPPING CONFIGURATION FOR JADUAL 18.0 (Stok Modal)
# ==========================================
# TODO: Map the EXCEL ROW NUMBER to the METRIC NAME in the database
# Example: 7: "jumlah_penduduk", 11: "warganegara"
ROW_MAP = {
    # Sekolah kerajaan & bantuan kerajaan (Government school & government aid)
    10: "bilangan_murid_prasekolah_kerajaan",        # Prasekolah / Pre-school
    13: "bilangan_murid_rendah_kerajaan",            # Rendah / Primary
    16: "bilangan_murid_menengah_rendah_kerajaan",   # Menengah rendah / Lower secondary
    19: "bilangan_murid_menengah_atas_kerajaan",     # Menengah atas / Upper secondary
    22: "bilangan_murid_khas_rendah_kerajaan",       # Pendidikan khas, peringkat rendah
    25: "bilangan_murid_khas_menengah_kerajaan",     # Pendidikan khas, peringkat menengah

    # Tadika, sekolah rendah dan menengah swasta (Private kindergarten, primary and secondary school)
    30: "bilangan_murid_swasta_tadika",              # Tadika / Kindergarten
    33: "bilangan_murid_swasta_rendah",              # Rendah / Primary
    36: "bilangan_murid_swasta_menengah",            # Menengah / Secondary
    39: "bilangan_murid_swasta_khas",                # Sekolah Pendidikan Khas
    42: "bilangan_murid_swasta_antarabangsa",        # Sekolah Antarabangsa
    45: "bilangan_murid_swasta_ekspatriat"           # Sekolah Ekspatriat
}

# TODO: Map the EXCEL COLUMN NUMBER to the YEAR STRING
COL_MAP = {
    6 : "2024",  
    7 : "2025" 
}

def populate_jadual_30_0(sheet, hierarchy, report_type):
    print(f"  -> Populating Jadual 30.0 (Bil Murid) untuk Malaysia_30_0")
    
    # 1. Fetch the Data Payload strictly for Malaysia
    metrics_data = get_metrics_dict("Malaysia", level='negeri')
    
    if not metrics_data:
            print(f"     [Warning] No data found for Malaysia.")
            return

    # ==========================================
        # DYNAMIC TABLE TITLE MODIFICATION
    # ==========================================
    title_bm = ": Bilangan murid pelbagai peringkat dan jenis sekolah, Malaysia, 2024 - 2025"
    title_en = ": Number of pupils of various level and types of school, Malaysia, 2024 - 2025"
        
    # Set the exact cells where your title sits in the template
    # Targeting Column C based on standard template behavior
    sheet.range("C2").value = title_bm
    sheet.range("C3").value = title_en

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