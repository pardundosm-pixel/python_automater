
from src.data_provider import get_metrics_dict
from src.excel_utils import inject_static_table

# ==========================================
# MAPPING CONFIGURATION FOR JADUAL 16.0 (Akaun SP)
# ==========================================
# TODO: Map the EXCEL ROW NUMBER to the METRIC NAME in the database
# Example: 7: "jumlah_penduduk", 11: "warganegara"
ROW_MAP = {
    # Akaun Satelit Pelancongan (Nilai)
    11:  "akaun_satelit_pelancongan_nilai_nilai_ditambah_kasar_industri_pelancongan",
    14:  "akaun_satelit_pelancongan_nilai_keluaran_dalam_negeri_kasar_pelancongan_langsung",
    17:  "akaun_satelit_pelancongan_nilai_perbelanjaan_pelancongan_inbound_bagi_pelawat",
    20:  "akaun_satelit_pelancongan_nilai_perbelanjaan_pelancongan_domestik_bagi_pelawat",
    23:  "akaun_satelit_pelancongan_nilai_perbelanjaan_pelancongan_outbound_bagi_pelawat",
    26:  "akaun_satelit_pelancongan_nilai_guna_tenaga_dalam_industri_pelancongan",
    
    # Akaun Satelit Pelancongan - Perubahan Peratusan Tahunan (%)
    32:  "akaun_satelit_pelancongan_perubahan_peratusan_tahunan__nilai_ditambah_kasar_industri_pelancongan",
    35:  "akaun_satelit_pelancongan_perubahan_peratusan_tahunan_keluaran_dalam_negeri_kasar_pelancongan_langsung",
    38:  "akaun_satelit_pelancongan_perubahan_peratusan_tahunan_perbelanjaan_pelancongan_inbound_bagi_pelawat",
    41:  "akaun_satelit_pelancongan_perubahan_peratusan_tahunan_perbelanjaan_pelancongan_domestik_bagi_pelawat",
    44:  "akaun_satelit_pelancongan_perubahan_peratusan_tahunan_perbelanjaan_pelancongan_outbound_bagi_pelawat",
    47:  "akaun_satelit_pelancongan_perubahan_peratusan_tahunan_guna_tenaga_dalam_industri_pelancongan",
    
    #Akaun Satelit Pelancongan - Peratus Sumbangan kepada KDNK (%)
    54:  "peratus_sumbangan_kepada_keluaran_dalam_negeri_kasar_nilai_ditambah_kasar_industri_pelancongan",
    57:  "peratus_sumbangan_kepada_keluaran_dalam_negeri_kasar_keluaran_dalam_negeri_kasar_pelancongan_langsung"
}

# TODO: Map the EXCEL COLUMN NUMBER to the YEAR STRING
COL_MAP = {
    6 : "2023",  
    7 : "2024",  
    8 : "2025"   
}

def populate_jadual_16_0(sheet, hierarchy, report_type):
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