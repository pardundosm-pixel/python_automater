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
    8: ("bilangan_taman_perumahan", "2025"),
    9: ("penduduk_taman_1_1000", "2025"),
    10: ("penduduk_taman_1001_9999", "2025"),
    11: ("penduduk_taman_lebih_10000", "2025"),

    13: ("bilangan_kampung", "2025"),
    14: ("penduduk_kampung_1_1000", "2025"),
    15: ("penduduk_kampung_1001_9999", "2025"),
    16: ("penduduk_kampung_lebih_10000", "2025"),

    18: ("bilangan_dewan", "2025"),
    19: ("dewan_1_200_orang", "2025"),
    20: ("dewan_201_500_orang", "2025"),
    21: ("dewa_lebih_501_orang", "2025"),

    23: ("bilangan_balai_raya", "2025"),
    25: ("bilangan_pusat_masyarakat", "2025"),
    28: ("bilangan_taman_permainan", "2025"),
    30: ("bilangan_padang_permainan", "2025"),
    32: ("bilangan_taman_rekreasi", "2025"),

    34: ("bilangan_trek_joging", "2025"),
    35: ("joging_1_9_99_km", "2025"),
    36: ("joging_lebih_10_km", "2025"),

    38: ("bilangan_trek_basikal", "2025"),

    40: ("bilangan_gelanggang_futsal", "2025"),
    41: ("futsal_1_5", "2025"),
    42: ("futsal_lebih_6", "2025"),

    44: ("bilangan_gelanggang_badminton", "2025"),
    46: ("badminton_1_5", "2025"),
    48: ("badminton_lebih_6", "2025"),

    51: ("bilangan_gimnasium", "2025"),
    52: ("gimansium_1_20_orang", "2025"),
    53: ("gimansium_lebih_21_orang", "2025"),

    55: ("bilangan_padang_bola", "2025"),
    56: ("bilangan_padang_rumput_semula_jadi", "2025"),
    58: ("bilangan_padang_rumput_tiruan", "2025"),

    61: ("bilangan_stadium", "2025"),
    62: ("stadium_1_3000_orang", "2025"),
    63: ("stadium_3001_9999_orang", "2025"),
    64: ("stadium_lebih_9999_orang", "2025"),

    66: ("bilangan_pusat_boling", "2025"),
    67: ("boling_1_20_lorong", "2025"),
    68: ("boling_lebih_21_lorong", "2025"),
}

# TODO: Set the furthest row down that contains dummy template data for cleanup
MAX_ROW_TO_CLEAN = 50 # Adjust based on how far down the template's dummy data goes

# ==========================================
# 3. REPORT INJECTION ENGINE
# ==========================================
# TODO: Rename the function to match the specific jadual (e.g., populate_jadual_3_0)
def populate_jadual_13(sheet, hierarchy, report_type):
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