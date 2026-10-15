import pandas as pd
from src.data_provider import get_metrics_dict

# ==========================================
# MAPPING CONFIGURATION FOR JADUAL 14.1 (MALAYSIA)
# ==========================================
# TODO: Map the EXCEL ROW NUMBER to the METRIC NAME in the database
# Example: 7: "jumlah_penduduk", 11: "warganegara"
ROW_MAP = {
    # KDNK mengikut Jenis Aktiviti Ekonomi pada Harga Malar 2015(RM juta)
    10:  "kdnk_pada_harga_malar_2015_jumlah",
    11:  "kdnk_pada_harga_malar_2015_pertanian",
    12:  "kdnk_pada_harga_malar_2015_perlombongan_dan_pengkuarian",
    13:  "kdnk_pada_harga_malar_2015_pembuatan",
    14:  "kdnk_pada_harga_malar_2015_pembinaan",
    15:  "kdnk_pada_harga_malar_2015_perkhidmatan",
    16:  "kdnk_pada_harga_malar_2015_duti_import",
    
    # KDNK mengikut Jenis Aktiviti Ekonomi pada Harga Malar 2015 - Perubahan Peratusan Tahunan (%)
    21:  "kdnk_perubahan_peratusan_tahunan_jumlah", # Not in database
    22:  "kdnk_pada_harga_malar_2015_perubahan_peratusan_tahunan_pertanian",
    23:  "kdnk_pada_harga_malar_2015_perubahan_peratusan_tahunan_perlombongan_dan_pengkuarian",
    24:  "kdnk_pada_harga_malar_2015_perubahan_peratusan_tahunan_pembuatan",
    25:  "kdnk_pada_harga_malar_2015_perubahan_peratusan_tahunan_pembinaan",
    26:  "kdnk_pada_harga_malar_2015_perubahan_peratusan_tahunan_perkhidmatan",
    27:  "kdnk_pada_harga_malar_2015_perubahan_peratusan_tahunan_duti_import",
    
    # KDNK mengikut Jenis Perbelanjaan pada Harga Malar 2015 (RM juta)
    32:  "perbelanjaan_kdnk_pada_harga_malar_2015_jumlah",
    33:  "perbelanjaan_kdnk_pada_harga_malar_2015_perbelanjaan_pengunaan_akhir_swasta",
    34:  "perbelanjaan_kdnk_pada_harga_malar_2015_perbelanjaan_pengunaan_akhir_kerajaan",
    35:  "perbelanjaan_kdnk_pada_harga_malar_2015_pembentukan_modal_tetap_kasar",
    36:  "perbelanjaan_kdnk_pada_harga_malar_2015_perubahan_inventori_dan_barangan_berharga",
    37:  "perbelanjaan_kdnk_pada_harga_malar_2015_eksport_barangan_dan_perkhidmatan",
    38:  "perbelanjaan_kdnk_pada_harga_malar_2015_tolak_import_barangan_dan_perkhidmatan",
    
    # KDNK mengikut Jenis Perbelanjaan pada Harga Malar 2015 - Perubahan Peratusan Tahunan (%)
    44:  "perbelanjaan_kdnk_perubahan_peratusan_tahunan_pada_harga_malar_2015_jumlah",
    45:  "perbelanjaan_kdnk_perubahan_peratusan_tahunan_pada_harga_malar_2015_perbelanjaan_pengunaan_akhir_swasta",
    46:  "perbelanjaan_kdnk_perubahan_peratusan_tahunan_pada_harga_malar_2015_perbelanjaan_pengunaan_akhir_kerajaan",
    47:  "perbelanjaan_kdnk_perubahan_peratusan_tahunan_pada_harga_malar_2015_pembentukan_modal_tetap_kasar",
    48:  "perbelanjaan_kdnk_perubahan_peratusan_tahunan_pada_harga_malar_2015_perubahan_inventori_dan_barangan_berharga",
    49:  "perbelanjaan_kdnk_perubahan_peratusan_tahunan_pada_harga_malar_2015_eksport_barangan_dan_perkhidmatan",
    50:  "perbelanjaan_kdnk_perubahan_peratusan_tahunan_pada_harga_malar_2015_tolak_import_barangan_dan_perkhidmatan"
    

}

# TODO: Map the EXCEL COLUMN NUMBER to the YEAR STRING
COL_MAP = {
    6: "2023",  
    7: "2024",  
    8: "2025"   
}

def populate_jadual_14_1(sheet, hierarchy, report_type):
    print(f"  -> Populating Jadual 14.0 (KDNK) untuk Malaysia_14_1")
    
    # 1. Fetch the Data Payload strictly for Malaysia
    metrics_data = get_metrics_dict("Malaysia", level='malaysia')
    
    if not metrics_data:
            print(f"     [Warning] No data found for Malaysia.")
            return

    # ==========================================
        # DYNAMIC TABLE TITLE MODIFICATION
    # ==========================================
    title_bm = ": Keluaran Dalam Negeri Kasar (KDNK), Malaysia, 2023 - 2025 (samb.)"
    title_en = ": Gross Domestic Product (GDP), Malaysia, 2023 - 2025 (cont'd)"
        
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