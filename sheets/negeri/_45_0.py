import pandas as pd
from src.data_provider import get_metrics_dict
from src.excel_utils import safe_write

# ==========================================
# MAPPING CONFIGURATION FOR JADUAL 45.0 (NEGERI)
# ==========================================
# Map the EXCEL ROW NUMBER to the METRIC NAME in the database
ROW_MAP = {
    # Statistik Utama Tenaga Buruh
    7: "stastistik_tenaga_buruh_tenaga_buruh",
    8:  "statistik_tenaga_buruh_penduduk_bekerja",      
    9:  "stastistik_tenaga _buruh_penganggur",
    10: "statistik_tenaga_buruh_luar_tenaga_buruh",
    11: "statistik_tenaga_buruh_kadar_penyertaan_tenaga_buruh",
    12: "statistik_tenaga_buruh_kadar_pengangguran",

    # Penduduk Bekerja Mengikut Kemahiran
    15: "penduduk_bekerja_mahir",
    16: "penduduk_bekerja_separuhmahir",
    17: "penduduk_bekerja_berkemahiran_rendah",
    
    # Penduduk Bekerja Mengikut Pencapaian Pendidikan
    20: "penduduk_bekerja_tertiari",           
    21: "penduduk_bekerja_menengah",           
    22: "penduduk_bekerja_rendah",        
    23: "penduduk_bekerja_tiada_pendidikan_rasmi",          

    # Statistik Gaji dan Upah
    26: "statistik_gaji_upah_bilangan_penerima",
    27: "statistik_gaji_upah_penengah_gaji",
    28: "statistik_gaji_upah_purata_gaji"
}

# Map the EXCEL COLUMN NUMBER to the YEAR STRING
# Based on screenshot: Column J = 10, Column K = 11, Column L = 12
COL_MAP = {
    5: "2023",  
    6: "2024",  
    7: "2025"   
}

# ==========================================
# 2. REPORT INJECTION ENGINE
# ==========================================
def populate_jadual_45_0(sheet, hierarchy, report_type):
    state_name = hierarchy.get('state_name', 'Unknown State')
    state_code = hierarchy.get('state_code')
    print(f"  -> Populating Jadual 42.0 (KDNK) untuk {state_name}")

    metrics_data = get_metrics_dict(state_code, level='negeri')
    if not metrics_data:
        print(f"     [Warning] No data found for {state_name}.")
        return

    # 1. Inject Titles Safely (Row, Column)
    title_bm = f": Keluaran Dalam Negeri Kasar (KDNK), {state_name}, 2023 - 2025p"
    title_en = f": Gross Domestic Product (GDP), {state_name}, 2023 - 2025p"
    safe_write(sheet, 3, 3, title_bm) # C3
    safe_write(sheet, 4, 3, title_en) # C4