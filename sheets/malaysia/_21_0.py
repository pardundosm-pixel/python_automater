
from src.data_provider import get_metrics_dict
from src.excel_utils import inject_static_table

# ==========================================
# MAPPING CONFIGURATION FOR JADUAL 21.0 (Pelancongan Domestik)
# ==========================================
# TODO: Map the EXCEL ROW NUMBER to the METRIC NAME in the database
# Example: 7: "jumlah_penduduk", 11: "warganegara"
ROW_MAP = {
    # Jumlah terimaan (RM juta)
    8 :  "jumlah_terimaan",
    10:  "jumlah_pelawat_domestik",
    11:  "isi_rumah_yang_dilawati",
    
    # Jumlah Pelawat ('000)
    13:  "jumlah_pelawat",
    15:  "pelawat_harian",
    16:  "pelancong",
    
    # Jumlah perjalanan Pelancongan ('000)
    18:  "jumlah_perjalanan_pelancongan"
    
}

# TODO: Map the EXCEL COLUMN NUMBER to the YEAR STRING
COL_MAP = {
    5 : "2023",  
    6 : "2024",  
    7 : "2025"   
}

def populate_jadual_21_0(sheet, hierarchy, report_type):
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