import pandas as pd
from src.data_provider import get_metrics_dict
from src.excel_utils import safe_write, inject_static_table

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
    print(f"  -> Populating Jadual 47.0 (Kemalangan Pekerjaan) untuk {state_name}")

    metrics_data = get_metrics_dict(state_code, level='negeri')
    if not metrics_data:
        print(f"     [Warning] No data found for {state_name}.")
        return

    # 1. Title Generation
    title_bm = f": Statistik utama kecederaan pekerjaan, {state_name}, 2022 - 2024"
    title_en = f": Principal statistics of occupational injury, {state_name}, 2022 - 2024"
    
    # 2. Inject Titles Safely
    safe_write(sheet, 3, 3, title_bm) # C3
    safe_write(sheet, 4, 3, title_en) # C4

    # 3. Trigger Centralized Injection
    inject_static_table(sheet, metrics_data, ROW_MAP, COL_MAP)