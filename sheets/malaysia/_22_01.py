
from src.data_provider import get_metrics_dict
from src.excel_utils import inject_static_table

# ==========================================
# MAPPING CONFIGURATION FOR JADUAL 22.01 (Pasaran Buruh 1)
# ==========================================
# TODO: Map the EXCEL ROW NUMBER to the METRIC NAME in the database
# Example: 7: "jumlah_penduduk", 11: "warganegara"
ROW_MAP = {
    # Statistik Siswazah
    7 :  "siswazah_tenaga_buruh",
    8 :  "siswazah_penduduk_bekerja",
    9 :  "siswazah_penganggur",
    10:  "siswazah_luar_tenaga_buruh",
    11:  "siswazah_kadar_penyertaan_tenaga_buruh",
    12:  "siswazah_kadar_pengangguran",
    13:  "siswazah_penengah_gaji",
    14:  "siswazah_purata_gaji",
    
    # Produktiviti Buruh
    17:  "produktiviti_nilai_tambah_jam_bekerja_rm",
    18:  "peratus_produktiviti_nilai_tambah_jam_bekerja",
    19:  "jumlah_produktiviti_nilai_tambah_pekerja",
    20:  "peratus_produktiviti_nilai_tambah_pekerja",
    
    # Prestasi Malaysia dalam Buku Saing Dunia 2025
    22:  "prestasi_malaysia_buku_daya_saing_dunia",
    23:  "prestasi_malaysia_prestasi_ekonomi",
    24:  "prestasi_malaysia_kecekapan_kerajaan",
    25:  "prestasi_malaysia_kecekapan_perniagaan",
    26:  "prestasi_malaysia_infrastruktur"
    
}

# TODO: Map the EXCEL COLUMN NUMBER to the YEAR STRING
COL_MAP = {
    5 : "2023",  
    6 : "2024",  
    7 : "2025"   
}

# ==========================================
# REPORT INJECTION ENGINE
# ==========================================
def populate_jadual_22_01(sheet, hierarchy, report_type):
    print("  -> Populating Jadual 22.01 (Pasaran Buruh) untuk Malaysia")
    
    metrics_data = get_metrics_dict("00", level='negeri')
    if not metrics_data:
        print("     [Warning] No data found for Malaysia.")
        return

    # Titles (Openpyxl syntax)
    sheet["C1"] = ": Statistik Pasaran Buruh, Malaysia, 2023 - 2025 (samb.)"
    sheet["C2"] = ": Labour Market Statistics, Malaysia, 2023 - 2025 (cont'd)"

    inject_static_table(sheet, metrics_data, ROW_MAP, COL_MAP)