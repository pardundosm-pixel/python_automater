import pandas as pd
import re
from src.excel_utils import get_dynamic_boundaries, inject_dynamic_table
from src.data_provider import get_metrics_dict

# ==========================================
# 1. MAPPING CONFIGURATION 
# ==========================================
# TODO: Set the unique text in Column B to find this table's starting row
HEADER_ANCHOR = "PERTUBUHAN/ SYARIKAT PERNIAGAAN" # e.g., "MAKLUMAT ASAS" or "KEMISKINAN"

# TODO: Map the EXCEL ROW NUMBER to a TUPLE of ("METRIC_NAME", "YEAR")
# Format: { Excel_Row: ("metric_name_in_database", "Year") }
ROW_MAP = {
    10: ("aktiviti_kewangan_insurans_takaful_aktiviti_hartanah_aktiviti_profesional_aktiviti_pentadbiran_dan_kesenian", "2025"),
    15: ("pendidikan_dan_aktiviti_kesihatan_kemanusiaan_kerja_sosial", "2025"),
    19: ("aktiviti_perkhidmatan_lain", "2025"),
    22: ("dandanan_rambut_dan_penjagaan_kecantikan_lain", "2025"),
    25: ("lain_lain_aktiviti_di_bawah_aktiviti_perkhidmatan_lain", "2025"),
}

# TODO: Set the furthest row down that contains dummy template data for cleanup
MAX_ROW_TO_CLEAN = 35 # Adjust based on how far down the template's dummy data goes

# ==========================================
# 3. REPORT INJECTION ENGINE
# ==========================================
# TODO: Rename the function to match the specific jadual (e.g., populate_jadual_3_0)
def populate_jadual_12_1_1(sheet, hierarchy, report_type):
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