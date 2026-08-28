
from src.data_provider import get_metrics_dict
from src.excel_utils import inject_static_table

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
    16:  "kdnk_pada_harga_malar_2015_tambah_duti_import", #n.a
    
    # KDNK mengikut Jenis Aktiviti Ekonomi pada Harga Malar 2015 - Perubahan Peratusan Tahunan (%)
    21:  "kdnk_pada_harga_malar_2015_perubahan_peratusan_tahunan_jumlah",
    22:  "kdnk_pada_harga_malar_2015_perubahan_peratusan_tahunan_pertanian",
    23:  "kdnk_pada_harga_malar_2015_perubahan_peratusan_tahunan_perlombongan_dan_pengkuarian",
    24:  "kdnk_pada_harga_malar_2015_perubahan_peratusan_tahunan_pembuatan",
    25:  "kdnk_pada_harga_malar_2015_perubahan_peratusan_tahunan_pembinaan",
    26:  "kdnk_pada_harga_malar_2015_perubahan_peratusan_tahunan_perkhidmatan",
    27:  "kdnk_pada_harga_malar_2015_perubahan_peratusan_tahunan_tambah_duti_import", #n.a
    
    # KDNK mengikut Jenis Perbelanjaan pada Harga Malar 2015 (RM juta)
    32:  "perbelanjaan_kdnk_pada_harga_malar_2015_jumlah",
    33:  "perbelanjaan_kdnk_pada_harga_malar_2015_perbelanjaan_penggunaan_akhir_swasta", #n.a
    34:  "perbelanjaan_kdnk_pada_harga_malar_2015_perbelanjaan_penggunaan_akhir_kerajaan", #n.a
    35:  "perbelanjaan_kdnk_pada_harga_malar_2015_pembentukan_modal_tetap_kasar",
    36:  "perbelanjaan_kdnk_pada_harga_malar_2015_perubahan_inventori_dan_barang_berharga", #n.a //barangan -> barang
    37:  "perbelanjaan_kdnk_pada_harga_malar_2015_eksport_barangan_dan_perkhidmatan",
    38:  "perbelanjaan_kdnk_pada_harga_malar_2015_tolak_import_barangan_dan_perkhidmatan",
    
    # KDNK mengikut Jenis Perbelanjaan pada Harga Malar 2015 - Perubahan Peratusan Tahunan (%)
    44:  "perbelanjaan_kdnk_perubahan_peratusan_tahunan_pada_harga_malar_2015_jumlah",
    45:  "perbelanjaan_kdnk_perubahan_peratusan_tahunan_pada_harga_malar_2015_perbelanjaan_penggunaan_akhir_swasta", #n.a
    46:  "perbelanjaan_kdnk_perubahan_peratusan_tahunan_pada_harga_malar_2015_perbelanjaan_penggunaan_akhir_kerajaan", #n.a
    47:  "perbelanjaan_kdnk_perubahan_peratusan_tahunan_pada_harga_malar_2015_pembentukan_modal_tetap_kasar",
    48:  "perbelanjaan_kdnk_perubahan_peratusan_tahunan_pada_harga_malar_2015_perubahan_inventori_dan_barang_berharga", #n.a //barangan -> barang
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