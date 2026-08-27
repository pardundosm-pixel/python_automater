import pandas as pd
from src.excel_utils import get_dynamic_boundaries, inject_dynamic_table, inject_static_table
from src.data_provider import get_metrics_dict

# ==========================================
# MAPPING CONFIGURATION FOR JADUAL 2.1 (NEGERI)
# ==========================================
# TODO: Map the EXCEL ROW NUMBER to the METRIC NAME in the database
# Example: 7: "jumlah_penduduk", 11: "warganegara"
ROW_MAP = {
    7:  "jumlah_penduduk",
    11: "penduduk_warganegara",
    12: "penduduk_bukan_warganegara",
    14: "penduduk_lelaki",
    15: "penduduk_perempuan",
    18: "peratus_penduduk_warganegara",
    19: "peratus_penduduk_bukan_warganegara",
    21: "purata_pertumbuhan_penduduk",
    26: "peratus_penduduk_bumiputera",
    27: "peratus_penduduk_cina",
    28: "peratus_penduduk_india",
    29: "peratus_penduduk_lain_lain",
    33: "penduduk_umur_0_14",
    35: "penduduk_umur_15_30", #variable baru
    37: "penduduk_umur_15_64",
    39: "penduduk_umur_65_lebih",
    41: "penduduk_umur_18_lebih",
    46: "jumlah_nisbah_tanggungan",
    47: "nisbah_tanggungan_umur_muda",
    48: "nisbah_tanggungan_umur_tua",
    50: "nisbah_jantina",
    52: "kepadatan_penduduk"
}

# TODO: Map the EXCEL COLUMN NUMBER to the YEAR STRING
COL_MAP = {
    4: "2024",  
    5: "2025",  
    6: "2026p"   
}

def populate_jadual_2_1(sheet, hierarchy, report_type):
    metrics_data = get_metrics_dict(hierarchy['state_code'], level='negeri')
    if not metrics_data: return

    sheet["C2"] = f": Anggaran penduduk pertengahan tahun, {hierarchy['state_name']}, 2024 - 2026p"
    sheet["C3"] = f": Mid-year population estimates, {hierarchy['state_name']}, 2024 - 2026p"
    
    inject_static_table(sheet, metrics_data, ROW_MAP, COL_MAP)