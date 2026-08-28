
from src.data_provider import get_metrics_dict
from src.excel_utils import inject_static_table

# ==========================================
# MAPPING CONFIGURATION FOR JADUAL 27.0 (Bilangan Hakim)
# ==========================================
# TODO: Map the EXCEL ROW NUMBER to the METRIC NAME in the database
# Example: 7: "jumlah_penduduk", 11: "warganegara"
ROW_MAP = {
    # Jumlah (Total)
    8:  "jumlah_bilangan_kehakiman",          # Jumlah/Total
    9:  "bilangan_kehakiman_lelaki",          # Lelaki/Male
    10: "bilangan_kehakiman_perempuan",       # Perempuan/Female

    # Ketua Hakim Negara (Chief Justice)
    13: "jumlah_ketua_hakim_negara",          # Jumlah/Total
    14: "ketua_hakim_negara_lelaki",          # Lelaki/Male
    15: "ketua_hakim_negara_perempuan",       # Perempuan/Female

    # Presiden Mahkamah Rayuan Malaysia (President Court of Appeal)
    18: "jumlah_presiden_mahkamah_rayuan_malaysia",          # Jumlah/Total
    19: "presiden_mahkamah_rayuan_malaysia_lelaki",          # Lelaki/Male
    20: "presiden_mahkamah_rayuan_malaysia_perempuan",       # Perempuan/Female

    # Hakim Besar Malaya (Chief Judge of Malaya)
    23: "jumlah_hakim_besar_malaysia",        # Jumlah/Total
    24: "hakim_besar_malaysia_lelaki",        # Lelaki/Male
    25: "hakim_besar_malaysia_perempuan",     # Perempuan/Female

    # Hakim Besar Sabah dan Sarawak (Chief Judge of Sabah & Sarawak)
    28: "jumlah_hakim_besar_sabah_dan_sarawak",          # Jumlah/Total
    29: "hakim_besar_sabah_dan_sarawak_lelaki",          # Lelaki/Male
    30: "hakim_besar_sabah_dan_sarawak_perempuan",       # Perempuan/Female

    # Hakim Mahkamah Persekutuan Malaysia (Federal Court Judges)
    33: "jumlah_hakim_mahkamah_persekutuan_malaysia",    # Jumlah/Total
    34: "hakim_mahkamah_persekutuan_malaysia_lelaki",    # Lelaki/Male
    35: "hakim_mahkamah_persekutuan_malaysia_perempuan", # Perempuan/Female

    # Hakim Mahkamah Rayuan Malaysia (Court of Appeal Judges)
    38: "jumlah_hakim_mahkamah_rayuan_malaysia",         # Jumlah/Total
    39: "hakim_mahkamah_rayuan_malaysia_lelaki",         # Lelaki/Male
    40: "hakim_mahkamah_rayuan_malaysia_perempuan",      # Perempuan/Female

    # Hakim Mahkamah Tinggi (High Court Judges)
    43: "jumlah_hakim_mahkamah_tinggi",          # Jumlah/Total
    44: "hakim_mahkamah_tinggi_lelaki",          # Lelaki/Male
    45: "hakim_mahkamah_tinggi_perempuan",       # Perempuan/Female

    # Pesuruhjaya Kehakiman Mahkamah Tinggi (Judicial Commissioners)
    48: "jumlah_pesuruhjaya_kehakiman_mahkamah_tinggi",          # Jumlah/Total
    49: "pesuruhjaya_kehakiman_mahkamah_tinggi_lelaki",          # Lelaki/Male
    50: "pesuruhjaya_kehakiman_mahkamah_tinggi_perempuan"        # Perempuan/Female
}

# TODO: Map the EXCEL COLUMN NUMBER to the YEAR STRING
COL_MAP = {
    8  : "2022",  
    9  : "2023",  
    10 : "2024"   
}

# ==========================================
# REPORT INJECTION ENGINE
# ==========================================
def populate_jadual_27_0(sheet, hierarchy, report_type):
    print("  -> Populating Jadual 27.0 (Bilangan Hakim) untuk Malaysia")
    
    metrics_data = get_metrics_dict("Malaysia", level='malaysia')
    if not metrics_data:
        print("     [Warning] No data found for Malaysia.")
        return

    # Titles (Openpyxl syntax)
    sheet["C3"] = ": Bilangan hakim di Badan Kehakiman mengikut jawatan dan jantina, Malaysia, 2022 - 2024"
    sheet["C4"] = ": Number of judges in the Judiciary board by position and sex, Malaysia, 2022 - 2024"

    inject_static_table(sheet, metrics_data, ROW_MAP, COL_MAP)