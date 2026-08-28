
from src.data_provider import get_metrics_dict
from src.excel_utils import inject_static_table

# ==========================================
# MAPPING CONFIGURATION FOR JADUAL 28.0 (Bilangan Hakim Syariah)
# ==========================================
# TODO: Map the EXCEL ROW NUMBER to the METRIC NAME in the database
# Example: 7: "jumlah_penduduk", 11: "warganegara"
ROW_MAP = {
    # Jumlah (Total)
    8:  "jumlah_kehakiman_syariah",          # Jumlah/Total
    9:  "kehakiman_syariah_lelaki",          # Lelaki/Male
    10: "kehakiman_syariah_perempuan",       # Perempuan/Female

    # Ketua Pengarah/ Ketua Hakim Syarie (Director General/ Syarie Chief Justice)
    13: "jumlah_ketua_pengarah",             # Jumlah/Total
    14: "ketua_pengarah_lelaki",             # Lelaki/Male
    15: "ketua_pengarah_perempuan",          # Perempuan/Female

    # Hakim Mahkamah Rayuan Syariah (Judges of the Court of Appeal Syarie)
    18: "jumlah_hakim_mahkamah_rayuan_syariah",         # Jumlah/Total
    19: "hakim_mahkamah_rayuan_syariah_lelaki",         # Lelaki/Male
    20: "hakim_mahkamah_rayuan_syariah_perempuan",      # Perempuan/Female

    # Ketua Hakim Syarie Negeri (State Syarie Chief Judge)
    23: "jumlah_ketua_hakim_syarie_negeri",             # Jumlah/Total
    24: "ketua_hakim_syarie_negeri_lelaki",             # Lelaki/Male
    25: "ketua_hakim_syarie_negeri_perempuan",          # Perempuan/Female

    # Ketua Pendaftar Mahkamah Syariah Negeri (Chief Register of the State Syariah Court)
    28: "jumlah_ketua_pendaftar_mahkamah_syariah_negeri",    # Jumlah/Total
    29: "ketua_pendaftar_mahkamah_syariah_negeri_lelaki",    # Lelaki/Male
    30: "ketua_pendaftar_mahkamah_syariah_negeri_perempuan", # Perempuan/Female

    # Hakim Syarie (Syarie Judge)
    33: "jumlah_hakim_syarie",                  # Jumlah/Total
    34: "hakim_syarie_lelaki",                  # Lelaki/Male
    35: "hakim_syarie_perempuan"                # Perempuan/Female
}

# TODO: Map the EXCEL COLUMN NUMBER to the YEAR STRING
COL_MAP = {
    8  : "2022",  
    9  : "2023",  
    10 : "2024"   
}

def populate_jadual_28_0(sheet, hierarchy, report_type):
    print("  -> Populating Jadual 28.0 (Bilangan Hakim Syariah) untuk Malaysia")
    
    metrics_data = get_metrics_dict("Malaysia", level='malaysia')
    if not metrics_data:
        print("     [Warning] No data found for Malaysia.")
        return

    # Titles (Openpyxl syntax)
    sheet["C3"] = ": Bilangan hakim di Kehakiman Syariah mengikut jawatan dan jantina, Malaysia, 2022 - 2024"
    sheet["C4"] = ": Number of judges in the Syariah Judiciary by position and sex, Malaysia, 2022 - 2024"
        
    inject_static_table(sheet, metrics_data, ROW_MAP, COL_MAP)