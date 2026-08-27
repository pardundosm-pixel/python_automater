from src.excel_utils import get_dynamic_boundaries, inject_dynamic_table
from src.data_provider import get_metrics_dict
import re

# ==========================================
# 1. MAPPING CONFIGURATION 
# ==========================================
# TODO: Set the unique text in Column B to find this table's starting row
HEADER_ANCHOR = "PENDAPATAN ISI RUMAH" # e.g., "MAKLUMAT ASAS" or "KEMISKINAN"

# TODO: Map the EXCEL ROW NUMBER to a TUPLE of ("METRIC_NAME", "YEAR")
# Format: { Excel_Row: ("metric_name_in_database", "Year") }
ROW_MAP = {
    11:  ("penengah", "2022"),
    12:  ("penengah", "2024"),
    14:  ("purata", "2022"),
    15:  ("purata", "2024"),
    17: ("perbelanjaan_purata", "2022"),
    18: ("perbelanjaan_purata", "2024"),
    23: ("pekali_gini", "2022"),
    24: ("pekali_gini", "2024"),
    26: ("kemiskinan", "2022"),
    27: ("kemiskinan", "2024"),
}

# TODO: Set the furthest row down that contains dummy template data for cleanup
MAX_ROW_TO_CLEAN = 28 # Adjust based on how far down the template's dummy data goes

def populate_jadual_5(sheet, hierarchy, report_type):
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

    sheet["C3"] = f": Statistik pendapatan dan kemiskinan, {loc_bm}, {hierarchy.get('state_name')}, 2022 - 2024"
    sheet["C4"] = f": Statistics of income and poverty, {loc_en}, {hierarchy.get('state_name')}, 2022 - 2024"

    inject_dynamic_table(sheet, target_row, start_col, max_slots, locations, ROW_MAP, MAX_ROW_TO_CLEAN, hierarchy)