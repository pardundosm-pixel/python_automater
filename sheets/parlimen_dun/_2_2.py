import pandas as pd
from src.excel_utils import get_dynamic_boundaries, inject_dynamic_table, inject_static_table
from src.data_provider import get_metrics_dict

# ==========================================
# 1. MAPPING CONFIGURATION 
# ==========================================
# Map the EXCEL ROW NUMBER to the METRIC NAME in the database
ROW_MAP = {
    8:  "jumlah_penduduk",
    12: "penduduk_warganegara",
    13: "penduduk_bukan_warganegara",
    15: "penduduk_lelaki",
    16: "penduduk_perempuan",
    19: "peratus_penduduk_warganegara",
    20: "peratus_penduduk_bukan_warganegara",
    24: "peratus_penduduk_bumiputera",
    25: "peratus_penduduk_cina",
    26: "peratus_penduduk_india",
    27: "peratus_penduduk_lain_lain",
    31: "penduduk_umur_0_14",
    33: "penduduk_umur_15_30",
    35: "penduduk_umur_15_64",
    37: "penduduk_umur_65_lebih",
    39: "penduduk_umur_18_lebih",
    44: "jumlah_nisbah_tanggungan",
    45: "umur_muda",
    46: "umur_tua",
    48: "nisbah_jantina",
    51: "kepadatan_penduduk"
}

# Map the EXCEL COLUMN INDEX to the YEAR STRING
COL_MAP = {
    6: "2024",
    7: "2025",
}

def populate_jadual_2_2(sheet, hierarchy, report_type):
    target_code = hierarchy.get('parl_code') or hierarchy.get('parent_parl_code')
    metrics_data = get_metrics_dict(target_code, level='parlimen')
    if not metrics_data: return

    parl_name = hierarchy.get('parl_name') or hierarchy.get('parent_parl_name')
    
    sheet["C3"] = f": Anggaran penduduk pertengahan tahun, Parlimen {parl_name}, {hierarchy.get('state_name')}, 2024 - 2025"
    sheet["C4"] = f": Mid-year population estimates, Parliament of {parl_name}, {hierarchy.get('state_name')}, 2024 - 2025"
    
    inject_static_table(sheet, metrics_data, ROW_MAP, COL_MAP)