
from src.data_provider import get_metrics_dict
from src.excel_utils import inject_static_table

# ==========================================
# MAPPING CONFIGURATION FOR JADUAL 24.0 (IHPR)
# ==========================================
# TODO: Map the EXCEL ROW NUMBER to the METRIC NAME in the database
# Example: 7: "jumlah_penduduk", 11: "warganegara"
ROW_MAP = {
    # Indeks Harga Pengeluar (IHPR)
    12: "indeks_harga_pengeluar_jumlah",                                 # Jumlah / Total
    15: "indeks_harga_pengeluar_pertanian_perhutanan_dan_perikanan",    # Pertanian, perhutanan dan perikanan
    18: "indeks_harga_pengeluar_perlombongan",                          # Perlombongan
    21: "indeks_harga_pengeluar_pembuatan",                             # Pembuatan
    24: "indeks_harga_pengeluar_bekalan_elektrik_dan_gas",              # Bekalan elektrik dan gas
    27: "indeks_harga_pengeluar_bekalan_air",                           # Bekalan air

    # Perubahan Peratus Tahunan (Annual % Change)
    34: "perubahan_peratus_tahunan_jumlah",                             # Jumlah / Total
    37: "perubahan_peratus_tahunan_pertanian_perhutanan_dan_perikanan", # Pertanian...
    40: "perubahan_peratus_tahunan_perlombongan",                       # Perlombongan
    43: "perubahan_peratus_tahunan_pembuatan",                          # Pembuatan
    46: "perubahan_peratus_tahunan_bekalan_elektrik_dan_gas",           # Bekalan elektrik dan gas
    49: "perubahan_peratus_tahunan_bekalan_air"                         # Bekalan air
}

# TODO: Map the EXCEL COLUMN NUMBER to the YEAR STRING
COL_MAP = {
    5  : "2023",  
    6  : "2024",  
    7  : "2025"   
}

# ==========================================
# REPORT INJECTION ENGINE
# ==========================================
def populate_jadual_24_0(sheet, hierarchy, report_type):
    print("  -> Populating Jadual 24.0 (IHPR) untuk Malaysia")
    
    metrics_data = get_metrics_dict("Malaysia", level='malaysia')
    if not metrics_data:
        print("     [Warning] No data found for Malaysia.")
        return

    # Titles (Openpyxl syntax)
    sheet["C1"] = ": Indeks Harga Pengeluar (IHPR) Pengeluaran Tempatan dan Perubahan Peratus Tahunan mengikut "
    sheet["C2"] = "  Sektor, Malaysia, 2023 - 2025 "
    sheet["C3"] = ": Producer Price Index (PPI) Local Production and Annual Percentage Change by Sector, Malaysia, "
    sheet["C4"] = "  2023 - 2025"
        
    inject_static_table(sheet, metrics_data, ROW_MAP, COL_MAP)