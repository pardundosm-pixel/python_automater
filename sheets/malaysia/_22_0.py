import pandas as pd
from src.data_provider import get_metrics_dict

# ==========================================
# MAPPING CONFIGURATION FOR JADUAL 22.0 (Pasaran Buruh)
# ==========================================
# TODO: Map the EXCEL ROW NUMBER to the METRIC NAME in the database
# Example: 7: "jumlah_penduduk", 11: "warganegara"
ROW_MAP = {
    # Statistik Utama Tenaga Buruh
    7 :  "stastistik_tenaga_buruh_tenaga_buruh",
    8 :  "statistik_tenaga_buruh_penduduk_bekerja",
    9 :  "stastistik_tenaga _buruh_penganggur",
    10:  "statistik_tenaga_buruh_luar_tenaga_buruh",
    11:  "statistik_tenaga_buruh_kadar_penyertaan_tenaga_buruh",
    12:  "statistik_tenaga_buruh_kadar_pengangguran",
    
    # Penduduk Bekerja Mengikut Kemahiran 
    15:  "penduduk_bekerja_mahir",
    16:  "penduduk_bekerja_separuhmahir",
    17:  "penduduk_bekerja_berkemahiran_rendah",
    
    # Penduduk Bekerja Mengikut Pencapaian Pendidikan
    20:  "penduduk_bekerja_tertiari",
    21:  "penduduk_bekerja_menengah",
    22:  "penduduk_bekerja_rendah",
    23:  "penduduk_bekerja_tiada_pendidikan_rasmi",
    
    # Statistik Gaji dan Upah
    26:  "statistik_gaji_upah_bilangan_penerima",
    27:  "statistik_gaji_upah_penengah_gaji",
    28:  "statistik_gaji_upah_purata_gaji",
}

# TODO: Map the EXCEL COLUMN NUMBER to the YEAR STRING
COL_MAP = {
    5 : "2023",  
    6 : "2024",  
    7 : "2025"   
}

def populate_jadual_22_0(sheet, hierarchy, report_type):
    print(f"  -> Populating Jadual 22.0 (Pasaran Buruh) untuk Malaysia_22_0")
    
    # 1. Fetch the Data Payload strictly for Malaysia
    metrics_data = get_metrics_dict("00", level='negeri')
    
    if not metrics_data:
            print(f"     [Warning] No data found for Malaysia.")
            return

    # ==========================================
        # DYNAMIC TABLE TITLE MODIFICATION
    # ==========================================
    title_bm = ": Statistik Pasaran Buruh, Malaysia, 2023 - 2025"
    title_en = ": Labour Market Statistics, Malaysia, 2023 - 2025"
        
    # Set the exact cells where your title sits in the template
    # Targeting Column C based on standard template behavior
    sheet.range("C1").value = title_bm
    sheet.range("C2").value = title_en

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