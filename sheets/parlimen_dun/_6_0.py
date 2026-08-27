import pandas as pd
import re
from src.excel_utils import get_dynamic_boundaries, inject_dynamic_table
from src.data_provider import get_metrics_dict

def safe_add(val1, val2):
    """Safely adds two values, ignoring 'n.a' strings. Returns 'n.a' if both are missing."""
    def parse_val(v):
        if pd.isna(v) or str(v).strip().lower() in ['', 'n.a', 'n.a.', 'nan']:
            return None
        try:
            return float(v)
        except ValueError:
            return None

    num1 = parse_val(val1)
    num2 = parse_val(val2)

    if num1 is None and num2 is None:
        return "n.a"
    
    return (num1 or 0.0) + (num2 or 0.0)

# ==========================================
# 1. MAPPING CONFIGURATION 
# ==========================================
# TODO: Set the unique text in Column B to find this table's starting row
HEADER_ANCHOR = "PENDIDIKAN" # e.g., "MAKLUMAT ASAS" or "KEMISKINAN"

# TODO: Map the EXCEL ROW NUMBER to a TUPLE of ("METRIC_NAME", "YEAR")
# Format: { Excel_Row: ("metric_name_in_database", "Year") }
ROW_MAP = {
    9: ('bilangan_sekolah', '2024'), #tukar tahun
    10: ('bilangan_sekolah', '2025'),
    12: ('bilangan_sekolah_rendah', '2024'),
    13: ('bilangan_sekolah_rendah', '2025'),
    15: ('bilangan_sekolah_menengah', '2024'),
    16: ('bilangan_sekolah_menengah', '2025'),
    18: ('bilangan_guru', '2024'),
    19: ('bilangan_guru', '2025'),
    21: ('bilangan_guru_rendah', '2024'),
    22: ('bilangan_guru_rendah', '2025'),
    24: ('bilangan_guru_menengah', '2024'),
    25: ('bilangan_guru_menengah', '2025'),
    27: ('bilangan_murid', '2024'),
    28: ('bilangan_murid', '2025'),
    30: ('bilangan_murid_rendah', '2024'),
    31: ('bilangan_murid_rendah', '2025'),
    33: ('bilangan_murid_menengah', '2024'),
    34: ('bilangan_murid_menengah', '2025'),
    36: ('sekolah_1sesi', '2024'), # sum sekolah_rendah_1sesi + sekolah_menengah_1sesi
    37: ('sekolah_1sesi', '2025'), # sum sekolah_rendah_1sesi + sekolah_menengah_1sesi
    39: ('sekolah_rendah_1sesi', '2024'),
    40: ('sekolah_rendah_1sesi', '2025'),
    42: ('sekolah_menengah_1sesi', '2024'),
    43: ('sekolah_menengah_1sesi', '2025'),
    45: ('sekolah_2sesi', '2024'), # sum sekolah_rendah_2sesi + sekolah_menengah_2sesi
    46: ('sekolah_2sesi', '2025'), # sum sekolah_rendah_2sesi + sekolah_menengah_2sesi
    48: ('sekolah_rendah_2sesi', '2024'),
    49: ('sekolah_rendah_2sesi', '2025'),
    51: ('sekolah_menengah_2sesi', '2024'),
    52: ('sekolah_menengah_2sesi', '2025'),
    54: ('bilangan_bilik_darjah', '2024'),
    55: ('bilangan_bilik_darjah', '2025'),
    57: ('bilangan_bilik_darjah_rendah', '2024'),
    58: ('bilangan_bilik_darjah_rendah', '2025'),
    60: ('bilangan_bilik_darjah_menengah', '2024'),
    61: ('bilangan_bilik_darjah_menengah', '2025'),
}

# TODO: Set the furthest row down that contains dummy template data for cleanup
MAX_ROW_TO_CLEAN = 62 # Adjust based on how far down the template's dummy data goes

def populate_jadual_6(sheet, hierarchy, report_type):
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

    sheet["C3"] = f": Statistik pendidikan, {loc_bm}, {hierarchy.get('state_name')}, 2024 - 2025"
    sheet["C4"] = f": Statistics of education, {loc_en}, {hierarchy.get('state_name')}, 2024 - 2025"

    inject_dynamic_table(sheet, target_row, start_col, max_slots, locations, ROW_MAP, MAX_ROW_TO_CLEAN, hierarchy)