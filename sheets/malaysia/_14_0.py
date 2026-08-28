
from src.data_provider import get_metrics_dict
from src.excel_utils import inject_static_table

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
    16: "kdnk_perbelanjaan_harga_semasa_tambah_duit_import", #n.a

    # KDNK mengikut Jenis Aktiviti Ekonomi pada Harga Semasa - Perubahan Peratusan Tahunan (%)
    21: "kdnk_harga_semasa_perubahan_peratusan_tahunan_jumlah",
    22: "kdnk_perubahan_peratusan_tahunan_pertanian",
    23: "kdnk_perubahan_peratusan_tahunan_perlombongan_dan_pengkuarian",
    24: "kdnk_perubahan_peratusan_tahunan_pembuatan",
    25: "kdnk_perubahan_peratusan_tahunan_pembinaan",
    26: "kdnk_perubahan_peratusan_tahunan_perkhidmatan",
    27: "kdnk_perubahan_peratusan_tahunan_tambah_duti_import", #n.a
    
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
# REPORT INJECTION ENGINE
# ==========================================
def populate_jadual_14_0(sheet, hierarchy, report_type):
    print("  -> Populating Jadual 14.0 (KDNK) untuk Malaysia")

    metrics_data = get_metrics_dict("Malaysia", level='malaysia')
    if not metrics_data:
        print("     [Warning] No data found for Malaysia.")
        return

    # Titles (Updated to Openpyxl syntax)
    sheet["C3"] = ": Keluaran Dalam Negeri Kasar (KDNK), Malaysia, 2023 - 2025"
    sheet["C4"] = ": Gross Domestic Product (GDP), Malaysia, 2023 - 2025"

    # Single-line data injection
    inject_static_table(sheet, metrics_data, ROW_MAP, COL_MAP)