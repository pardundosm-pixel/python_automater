
from src.data_provider import get_metrics_dict
from src.excel_utils import inject_static_table

# ==========================================
# MAPPING CONFIGURATION FOR JADUAL 25.0 (Kemalangan Pekerja 1)
# ==========================================
# TODO: Map the EXCEL ROW NUMBER to the METRIC NAME in the database
# Example: 7: "jumlah_penduduk", 11: "warganegara"
ROW_MAP = {
    # Jumlah Kecederaan Pekerjaan
    11: "jumlah_kadar_kemalangan",
    
    # Kewarganegaraan
    14: "kadar_kemalangan_warganegara",
    15: "kadar_kemalangan_bukan_warganegara",
    
    # Jantina
    18: "kadar_kemalangan_lelaki",
    19: "kadar_kemalangan_perempuan",
    
    # Sektor OSHA (Kecederaan)
    23: "kadar_kemalangan_pertanian_perhutanan_dan_perikanan",
    24: "kadar_kemalangan_perlombongan_dan_pengkuarian",
    25: "kadar_kemalangan_pembuatan",
    26: "kadar_kemalangan_pembinaan",
    27: "kadar_kemalangan_utiliti",
    28: "kadar_kemalangan_perdagangan_borong_dan_runcit",
    29: "kadar_kemalangan_pengangkutan_penyimpanan_dan_komunikasi",
    31: "kadar_kemalangan_hotel_dan_restoran",
    32: "kadar_kemalangan_kewangan_insurans_hartanah_dan_perkhidmatan_perniagaan",
    34: "kadar_kemalangan_perkhidmatan",
    
    # Jumlah Kematian Pekerjaan
    41: "jumlah_kadar_kematian",
    
    # Kewarganegaraan (Kematian)
    44: "kadar_kematian_warganegara",
    45: "kadar_kematian_bukan_warganegara",
    
    # Jantina (Kematian)
    48: "kadar_kematian_lelaki",
    49: "kadar_kematian_perempuan",
    
    # Sektor OSHA (Kematian)
    53: "kadar_kematian_pertanian_perhutanan_dan_perikanan",
    54: "kadar_kematian_perlombongan_dan_pengkuarian",
    55: "kadar_kematian_pembuatan",
    56: "kadar_kematian_pembinaan",
    57: "kadar_kematian_utiliti",
    58: "kadar_kematian_perdagangan_borong_dan_runcit",
    59: "kadar_kematian_pengangkutan_penyimpanan_dan_komunikasi",
    61: "kadar_kematian_hotel_dan_restoran",
    62: "kadar_kematian_kewangan_insurans_hartanah_dan_perkhidmatan_perniagaan",
    64: "kadar_kematian_perkhidmatan"
}

# TODO: Map the EXCEL COLUMN NUMBER to the YEAR STRING
COL_MAP = {
    5  : "2022",  
    6  : "2023",  
    7  : "2024"   
}

# ==========================================
# REPORT INJECTION ENGINE
# ==========================================
def populate_jadual_25_0_1(sheet, hierarchy, report_type):
    print("  -> Populating Jadual 25.0 (Kemalangan Pekerja 1) untuk Malaysia")
    
    metrics_data = get_metrics_dict("00", level='negeri')
    if not metrics_data:
        print("     [Warning] No data found for Malaysia.")
        return

    # Titles (Openpyxl syntax)
    sheet["C3"] = ": Statistik utama kecederaan pekerjaan, Malaysia, 2022 - 2024 (samb.)"
    sheet["C4"] = ": Principal statistics of occupational injury, Malaysia, 2022 - 2024 (cont'd)"
        
    inject_static_table(sheet, metrics_data, ROW_MAP, COL_MAP)