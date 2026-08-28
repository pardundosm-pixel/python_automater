from src.data_provider import get_metrics_dict
from src.excel_utils import inject_static_table

# ==========================================
# MAPPING CONFIGURATION FOR JADUAL 18.0 (Stok Modal)
# ==========================================
# !!! No data for 2024 and 2025
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
    print("  -> Populating Jadual 30.0 (Bil Murid) untuk Malaysia")
    
    metrics_data = get_metrics_dict("00", level='negeri')
    if not metrics_data:
        print("     [Warning] No data found for Malaysia.")
        return

    # Titles (Openpyxl syntax)
    sheet["D2"] = ": Bilangan murid pelbagai peringkat dan jenis sekolah, Malaysia, 2024 - 2025"
    sheet["D3"] = ": Number of pupils of various level and types of school, Malaysia, 2024 - 2025"
        
    inject_static_table(sheet, metrics_data, ROW_MAP, COL_MAP)