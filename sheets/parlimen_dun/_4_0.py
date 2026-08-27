import pandas as pd
from src.excel_utils import get_dynamic_boundaries, inject_dynamic_table, inject_static_table
from src.data_provider import get_metrics_dict

# ==========================================
# MAPPING CONFIGURATION 
# ==========================================
# 1. Map Excel Row to Metric Name
ROW_MAP = {
    11: "tenaga_buruh",
    15: "penduduk_bekerja",
    19: "penganggur",
    23: "luar_tenaga_buruh",
    27: "kadar_penyertaan_tenaga_buruh",
    31: "kadar_pengangguran",

}

# 2. Map Excel Column to Year
COL_MAP = {
    11: "2024",  # E.g., Column D
}

def populate_jadual_4(sheet, hierarchy, report_type):
    target_code = hierarchy.get('parl_code') or hierarchy.get('parent_parl_code')
    metrics_data = get_metrics_dict(target_code, level='parlimen')
    if not metrics_data: return

    parl_name = hierarchy.get('parl_name') or hierarchy.get('parent_parl_name')
    
    sheet["C3"] = f": Statistik guna tenaga, Parlimen {parl_name}, {hierarchy.get('state_name')}, 2024"
    sheet["C4"] = f": Statistics of employment, Parliament {parl_name}, {hierarchy.get('state_name')}, 2024"
    
    inject_static_table(sheet, metrics_data, ROW_MAP, COL_MAP)