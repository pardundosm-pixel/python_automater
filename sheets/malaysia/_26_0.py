
from src.data_provider import get_metrics_dict
from src.excel_utils import inject_static_table

# ==========================================
# MAPPING CONFIGURATION FOR JADUAL 26.0 (Peratusan Ahli Parlimen)
# ==========================================
# TODO: Map the EXCEL ROW NUMBER to the METRIC NAME in the database
# Example: 7: "jumlah_penduduk", 11: "warganegara"
ROW_MAP = {
    # Dewan Negara (Senate)
    9:  "dewan_negara_lelaki",      # Lelaki/Male
    10: "dewan_negara_perempuan",   # Perempuan/Female

    # Dewan Rakyat (House of Representatives)
    13: "dewan_rakyat_lelaki",      # Lelaki/Male
    14: "dewan_rakyat_perempuan",   # Perempuan/Female

    # Menteri Kabinet (Cabinet Minister)
    17: "menteri_kabinet_lelaki",   # Lelaki/Male
    18: "menteri_kabinet_perempuan",# Perempuan/Female

    # Timbalan Menteri (Deputy Minister)
    21: "timbalan_menteri_lelaki",  # Lelaki/Male
    22: "timbalan_menteri_perempuan"# Perempuan/Female
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
def populate_jadual_26_0(sheet, hierarchy, report_type):
    print("  -> Populating Jadual 26.0 (Peratusan Ahli Parlimen) untuk Malaysia")
    
    metrics_data = get_metrics_dict("Malaysia", level='malaysia')
    if not metrics_data:
        print("     [Warning] No data found for Malaysia.")
        return

    # Titles (Openpyxl syntax)
    sheet["C3"] = ": Peratusan ahli parlimen dan anggota pentadbiran mengikut jantina, Malaysia, 2022 - 2024"
    sheet["C4"] = ": Percentage of members of parliament and administration by sex, Malaysia, 2022 - 2024"
        
    inject_static_table(sheet, metrics_data, ROW_MAP, COL_MAP)