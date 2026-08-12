import pandas as pd
from src.data_provider import get_metrics_dict

# ==========================================
# MAPPING CONFIGURATION FOR JADUAL 14.0 (MALAYSIA)
# ==========================================
# Map the EXCEL ROW NUMBER to the METRIC NAME in the database
ROW_MAP = {
    # KDNK mengikut Jenis Aktiviti Ekonomi pada Harga Semasa (RM juta)
    10: "kdnk_perbelanjaan_harga_semasa_jumlah",
    11: "kdnk_perbelanjaan_harga_semasa_pertanian",
    12: "kdnk_perbelanjaan_harga_semasa_perlombongan_dan_pengkuarian",
    13: "kdnk_perbelanjaan_harga_semasa_pembuatan",
    14: "kdnk_perbelanjaan_harga_semasa_pembinaan",
    15: "kdnk_perbelanjaan_harga_semasa_perkhidmatan",
    16: "kdnk_perbelanjaan_harga_semasa_duti_import",

    # KDNK mengikut Jenis Aktiviti Ekonomi pada Harga Semasa - Perubahan Peratusan Tahunan (%)
    21: "kdnk_perubahan_peratusan_tahunan_jumlah",
    22: "kdnk_perubahan_peratusan_tahunan_pertanian",
    23: "kdnk_perubahan_peratusan_tahunan_perlombongan_dan_pengkuarian",
    24: "kdnk_perubahan_peratusan_tahunan_pembuatan",
    25: "kdnk_perubahan_peratusan_tahunan_pembinaan",
    26: "kdnk_perubahan_peratusan_tahunan_perkhidmatan",
    27: "kdnk_perubahan_peratusan_tahunan_duti_import",      # Assuming row sequence continues
    
    # KDNK mengikut Jenis Perbelanjaan pada Harga Semasa (RM juta)
    32: "perbelanjaan_kdnk_jumlah",
    33: "perbelanjaan_kdnk_perbelanjaan_pengunaan_akhir_swasta",
    34: "perbelanjaan_kdnk_perbelanjaan_pengunaan_akhir_kerajaan",
    35: "perbelanjaan_kdnk_pembentukan_modal_tetap_kasar",
    36: "perbelanjaan_kdnk_perubahan_inventori_dan_barangan_berharga",
    37: "perbelanjaan_kdnk_eksport_barangan_dan_perkhidmatan",
    38: "perbelanjaan_kdnk_tolak_import_barangan_dan_perkhidmatan",
    
    # KDNK mengikut Jenis Perbelanjaan pada Harga Semasa - Perubahan Peratusan Tahunan (%)
    43: "perbelanjaan_kdnk_perubahan_peratusan_tahunan_jumlah",
    44: "perbelanjaan_kdnk_perubahan_peratusan_tahunan_perbelanjaan_penggunaan_akhir_swasta",
    45: "perbelanjaan_kdnk_perubahan_peratusan_tahunan_perbelanjaan_penggunaan_akhir_kerajaan",
    46: "perbelanjaan_kdnk_perubahan_peratusan_tahunan_pembentukan_modal_tetap_kasar",
    47: "perbelanjaan_kdnk_perubahan_peratusan_tahunan_perubahan_inventori_dan_barangan_berharga",
    48: "perbelanjaan_kdnk_perubahan_peratusan_tahunan_eksport_barangan_dan_perkhidmatan",
    49: "perbelanjaan_kdnk_perubahan_peratusan_tahunan_tolak_import_barangan_dan_perkhidmatan"
}

# Map the EXCEL COLUMN NUMBER to the YEAR STRING
# Based on screenshot: Column E (2022) = 5, Column F (2023) = 6, Column G (2024) = 7
COL_MAP = {
    5: "2023",  
    6: "2024",  
    7: "2025"   
}

# ==========================================
# 2. REPORT INJECTION ENGINE
# ==========================================
def populate_jadual_14(sheet, hierarchy, report_type):
    print(f"  -> Populating Jadual 14.0 (KDNK) untuk Malaysia")

    # 1. Fetch the Data Payload strictly for Malaysia
    metrics_data = get_metrics_dict("Malaysia", level='malaysia')
    
    if not metrics_data:
        print(f"     [Warning] No data found for Malaysia.")
        return

    # ==========================================
    # DYNAMIC TABLE TITLE MODIFICATION
    # ==========================================
    title_bm = ": Keluaran Dalam Negeri Kasar (KDNK), Malaysia, 2023 - 2025"
    title_en = ": Gross Domestic Product (GDP), Malaysia, 2023 - 2025"

    # Set the exact cells where your title sits in the template
    # Targeting Column C based on standard template behavior
    sheet.range("C3").value = title_bm
    sheet.range("C4").value = title_en
    # ==========================================

    # 2. Inject Data
    # Loop over the columns (Years) first
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