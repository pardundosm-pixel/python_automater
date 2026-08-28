
from src.data_provider import get_metrics_dict
from src.excel_utils import inject_static_table

# ==========================================
# MAPPING CONFIGURATION FOR JADUAL 29.0 (Bil KSU)
# ==========================================
# TODO: Map the EXCEL ROW NUMBER to the METRIC NAME in the database
# Example: 7: "jumlah_penduduk", 11: "warganegara"
ROW_MAP = {
    # Jumlah (Total) – Overall
    8:  "jumlah_ketua_setiausaha_timbalan_ketua_setiausaha_ketua_pengarah",  # Jumlah/Total
    9:  "ketua_setiausaha_timbalan_ketua_setiausaha_ketua_pengarah_lelaki",  # Lelaki/Male
    10: "ketua_setiausaha_timbalan_ketua_setiausaha_ketua_pengarah_perempuan", # Perempuan/Female

    # Ketua Setiausaha Negara (Chief Secretary to the Government)
    13: "ketua_setiausaha_negara",           # Jumlah/Total
    14: "ketua_setiausaha_negara_lelaki",    # Lelaki/Male
    15: "ketua_setiausaha_negara_perempuan", # Perempuan/Female

    # Ketua Setiausaha (Secretary General)
    18: "jumlah_ketua_setiausaha",           # Jumlah/Total
    19: "ketua_setiausaha_lelaki",           # Lelaki/Male
    20: "ketua_setiausaha_perempuan",        # Perempuan/Female

    # Timbalan Ketua Setiausaha (Deputy Secretary General)
    23: "jumlah_timbalan_setiausaha",        # Jumlah/Total
    24: "timbalan_ketua_setiausaha_lelaki",  # Lelaki/Male
    25: "timbalan_ketua_setiausaha_perempuan", # Perempuan/Female

    # Ketua-Ketua Pengarah, Pengarah dan Pengurus Besar Badan-Badan Berkanun
    29: "jumlah_ketua_ketua_pengarah_pengarah_pengurus_besar_badan_berkanun", # Jumlah/Total
    30: "ketua_ketua_pengarah_pengarah_pengurus_besar_badan_berkanun_lelaki", # Lelaki/Male
    31: "ketua_ketua_pengarah_pengarah_pengurus_besar_badan_berkanun_perempuan", # Perempuan/Female

    # Ketua-Ketua Pengarah Jabatan Persekutuan (Director General of Federal Departments)
    34: "jumlah_ketua_ketua_pengarah_jabatan_persekutuan", # Jumlah/Total
    35: "ketua_ketua_pengarah_jabatan_persekutuan_lelaki", # Lelaki/Male
    36: "ketua_ketua_pengarah_jabatan_persekutuan_perempuan" # Perempuan/Female
}

# TODO: Map the EXCEL COLUMN NUMBER to the YEAR STRING
COL_MAP = {
    7  : "2022",  
    8  : "2023",  
    9  : "2024"   
}

def populate_jadual_29_0(sheet, hierarchy, report_type):
    print("  -> Populating Jadual 29.0 (Bil KSU) untuk Malaysia")
    
    metrics_data = get_metrics_dict("Malaysia", level='malaysia')
    if not metrics_data:
        print("     [Warning] No data found for Malaysia.")
        return

    # Titles (Openpyxl syntax)
    sheet["C3"] = ": Bilangan Ketua Setiausaha, Timbalan Ketua Setiausaha dan Ketua Pengarah mengikut jawatan dan jantina, Malaysia, 2022 - 2024"
    sheet["C4"] = ": Number of Secretary General, Deputy Secretary General and Director General by position and sex, Malaysia, 2022 - 2024"
        
    inject_static_table(sheet, metrics_data, ROW_MAP, COL_MAP)