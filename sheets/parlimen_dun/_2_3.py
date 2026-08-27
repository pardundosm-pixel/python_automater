import pandas as pd
import re
from src.excel_utils import get_dynamic_boundaries, inject_dynamic_table
from src.data_provider import get_metrics_dict

# ==========================================
# 1. MAPPING CONFIGURATION 
# ==========================================
# TODO: Set the unique text in Column B to find this table's starting row
HEADER_ANCHOR = "PENDUDUK" # e.g., "MAKLUMAT ASAS" or "KEMISKINAN"

# TODO: Map the EXCEL ROW NUMBER to a TUPLE of ("METRIC_NAME", "YEAR")
# Format: { Excel_Row: ("metric_name_in_database", "Year") }
ROW_MAP = {
    8:  ("kelahiran_hidup", "2023"),
    9:  ("kelahiran_hidup", "2024"),
    10:  ("kelahiran_hidup", "2025"),
    12:  ("hidup_lelaki", "2023"),
    13:  ("hidup_lelaki", "2024"),
    14:  ("hidup_lelaki", "2025"),
    16:  ("hidup_perempuan", "2023"),
    17:  ("hidup_perempuan", "2024"),
    18:  ("hidup_perempuan", "2025"),
    20:  ("kematian", "2023"),
    21:  ("kematian", "2024"),
    22:  ("kematian", "2025"),
    24:  ("mati_lelaki", "2023"),
    25:  ("mati_lelaki", "2024"),
    26:  ("mati_lelaki", "2025"),
    28:  ("mati_perempuan", "2023"),
    29:  ("mati_perempuan", "2024"),
    30:  ("mati_perempuan", "2025"),
}

# TODO: Set the furthest row down that contains dummy template data for cleanup
MAX_ROW_TO_CLEAN = 30 # Adjust based on how far down the template's dummy data goes

def populate_jadual_2_3(sheet, hierarchy, report_type):
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

    sheet["C3"] = f": Bilangan kelahiran hidup dan kematian mengikut jantina, {loc_bm}, {hierarchy.get('state_name')}, 2023 - 2025"
    sheet["C4"] = f": Number of live births and deaths by sex, {loc_en}, {hierarchy.get('state_name')}, 2023 - 2025"

    inject_dynamic_table(sheet, target_row, start_col, max_slots, locations, ROW_MAP, MAX_ROW_TO_CLEAN, hierarchy)