import pandas as pd
import re
from src.excel_utils import get_dynamic_boundaries, inject_dynamic_table
from src.data_provider import get_metrics_dict

# ==========================================
# 1. MAPPING CONFIGURATION 
# ==========================================
# TODO: Set the unique text in Column B to find this table's starting row
HEADER_ANCHOR = "FASILITI AWAM" # e.g., "MAKLUMAT ASAS" or "KEMISKINAN"

# TODO: Map the EXCEL ROW NUMBER to a TUPLE of ("METRIC_NAME", "YEAR")
# Format: { Excel_Row: ("metric_name_in_database", "Year") }
ROW_MAP = {
    8: ("bilangan_pasar_malam", "2025"),
    9: ("pasar_malam_1_50_tapak", "2025"),
    10: ("pasar_malam_51_99_tapak", "2025"),
    11: ("pasar_malam_lebih_100_tapak", "2025"),

    13: ("bilangan_pasar_tani", "2025"),
    14: ("pasar_tani_1_50_tapak", "2025"),
    15: ("pasar_tani_51_99_tapak", "2025"),
    16: ("pasar_tani_lebih_100_tapak", "2025"),

    18: ("bilangan_pasar_awam", "2025"),
    19: ("pasar_awam_1_50_tapak", "2025"),
    20: ("pasar_awam_51_99_tapak", "2025"),
    21: ("pasar_awam_lebih_100_tapak", "2025"),

    23: ("bilangan_pusat_penjaja", "2025"),
    24: ("penjaja_1_20", "2025"),
    25: ("penjaja_21_40", "2025"),
    26: ("penjaja_lebih_41", "2025"),

    28: ("bilangan_stesen_bas", "2025"),
    30: ("bilangan_hentian_bas", "2025"),
    32: ("bilangan_hentian_teksi", "2025"),
    34: ("bilangan_tandas_awam", "2025"),
    36: ("bilangan_pusat_dialisis", "2025"),
    38: ("bilangan_pusat_keagamaan", "2025"),
    41: ("jumlah_masjid", "2025"),
    42: ("masjid_kerajaan", "2025"),
    43: ("masjid_negeri", "2025"),
    44: ("masjid_daerah", "2025"),
    45: ("masjid_mukim", "2025"),
    47: ("bilangan_surau", "2025"),
    49: ("bilangan_tokong", "2025"),
    51: ("bilangan_kuil", "2025"),
    53: ("bilangan_gereja", "2025"),
    55: ("bilangan_inap_desa", "2025"),
    57: ("inap_1_20_tetamu", "2025"),
    58: ("inap_21_40_tetamu", "2025"),
    59: ("inap_lebih_41_tetamu", "2025"),
}

# TODO: Set the furthest row down that contains dummy template data for cleanup
MAX_ROW_TO_CLEAN = 50 # Adjust based on how far down the template's dummy data goes

# ==========================================
# 3. REPORT INJECTION ENGINE
# ==========================================
# TODO: Rename the function to match the specific jadual (e.g., populate_jadual_3_0)
def populate_jadual_13_0_1(sheet, hierarchy, report_type):
    target_row, start_col, end_col, max_slots = get_dynamic_boundaries(sheet, HEADER_ANCHOR)
    if not start_col: return

    locations = []
    if report_type == 'parlimen':
        parl = hierarchy['parl_code']
        locations.append((parl, hierarchy['parl_name'], get_metrics_dict(parl, 'parlimen')))
        for dun in hierarchy['duns']: locations.append((dun['code'], dun['name'], get_metrics_dict(dun['code'], 'dun', parent_code=parl)))
    elif report_type == 'dun':
        parent = hierarchy['parent_parl_code']
        locations.append((parent, hierarchy['parent_parl_name'], get_metrics_dict(parent, 'parlimen')))
        dun = hierarchy['dun_code']
        locations.append((dun, hierarchy['dun_name'], get_metrics_dict(dun, 'dun', parent_code=parent)))
    
    loc_bm = f"Parlimen {hierarchy.get('parl_name')}" if report_type == 'parlimen' else f"DUN {hierarchy.get('dun_name')}"
    loc_en = f"Parliament {hierarchy.get('parl_name')}" if report_type == 'parlimen' else f"DUN {hierarchy.get('dun_name')}"

    # TODO: Update Title Strings
    sheet["C3"] = f": Tajuk BM, {loc_bm}, {hierarchy.get('state_name')}, 2024"
    sheet["C4"] = f": English Title, {loc_en}, {hierarchy.get('state_name')}, 2024"

    inject_dynamic_table(sheet, target_row, start_col, max_slots, locations, ROW_MAP, MAX_ROW_TO_CLEAN, hierarchy)