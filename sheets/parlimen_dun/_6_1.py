import pandas as pd
import re
from src.excel_utils import get_dynamic_boundaries, inject_dynamic_table
from src.data_provider import get_metrics_dict

def safe_add(*values):
    """
    Menambah berbilang nilai secara selamat.
    Mengabaikan nilai 'n.a', None, atau string kosong.
    Mengembalikan 'n.a' jika KESEMUA nilai adalah tiada/n.a.
    """
    parsed = []
    for v in values:
        if pd.isna(v) or str(v).strip().lower() in ['', 'n.a', 'n.a.', 'nan']:
            continue
        try:
            parsed.append(float(v))
        except ValueError:
            continue

    if not parsed:
        return "n.a"
    return sum(parsed)

# ==========================================
# 1. MAPPING CONFIGURATION 
# ==========================================
# TODO: Set the unique text in Column B to find this table's starting row
HEADER_ANCHOR = "PENDIDIKAN SWASTA" # e.g., "MAKLUMAT ASAS" or "KEMISKINAN"

# TODO: Map the EXCEL ROW NUMBER to a TUPLE of ("METRIC_NAME", "YEAR")
# Format: { Excel_Row: ("metric_name_in_database", "Year") }
ROW_MAP = {
    10:  ("bilangan_sekolah", "2024"), # sum var_smk + var_srk
    11:  ("bilangan_sekolah", "2025"), # sum var_smk + var_srk
    13:  ("sekolah_rendah", "2024"), # sum sekolah_rendah_akademik + sekolah_rendah_agama = var_srk
    14:  ("sekolah_rendah", "2025"), # sum sekolah_rendah_akademik + sekolah_rendah_agama = var_srk
    16:  ("sekolah_rendah_akademik", "2024"),
    17:  ("sekolah_rendah_akademik", "2025"),
    19:  ("sekolah_rendah_agama", "2024"),
    20:  ("sekolah_rendah_agama", "2025"),
    22:  ("sekolah_menengah", "2024"), # sum sekolah_menengah_akademik + sekolah_menengah_agama + sekolah_menengah_cina = var_smk
    23:  ("sekolah_menengah", "2025"), # sum sekolah_menengah_akademik + sekolah_menengah_agama + sekolah_menengah_cina = var_smk
    25:  ("sekolah_menengah_akademik", "2024"),
    26:  ("sekolah_menengah_akademik", "2025"),
    28:  ("sekolah_menengah_agama", "2024"),
    29:  ("sekolah_menengah_agama", "2025"),
    31:  ("sekolah_menengah_cina", "2024"),
    32:  ("sekolah_menengah_cina", "2025"),
    34:  ("sekolah_pendidikan_khas", "2024"),
    35:  ("sekolah_pendidikan_khas", "2025"),
    37:  ("sekolah_antarabangsa", "2024"),
    38:  ("sekolah_antarabangsa", "2025"),
    40:  ("sekolah_ekspatriat", "2024"),
    41:  ("sekolah_ekspatriat", "2025"),
    43:  ("bilangan_guru", "2024"), # sum var_srk_guru + var_smk_guru
    44:  ("bilangan_guru", "2025"), # sum var_srk_guru + var_smk_guru
    46:  ("bilangan_guru_sekolah_rendah", "2024"), # sum guru_sekolah_rendah_akademik + guru_sekolah_rendah_agama = var_srk_guru
    47:  ("bilangan_guru_sekolah_rendah", "2025"), # sum guru_sekolah_rendah_akademik + guru_sekolah_rendah_agama = var_srk_guru
    49:  ("guru_sekolah_rendah_akademik", "2024"),
    50:  ("guru_sekolah_rendah_akademik", "2025"),
    52:  ("guru_sekolah_rendah_agama", "2024"),
    53:  ("guru_sekolah_rendah_agama", "2025"),
    55:  ("bilangan_guru_sekolah_menengah", "2024"), # sum guru_sekolah_menengah_akademik + guru_sekolah_menengah_agama + guru_sekolah_menengah_cina = var_smk_guru
    56:  ("bilangan_guru_sekolah_menengah", "2025"), # sum guru_sekolah_menengah_akademik + guru_sekolah_menengah_agama + guru_sekolah_menengah_cina = var_smk_guru
    58:  ("guru_sekolah_menengah_akademik", "2024"),
    59:  ("guru_sekolah_menengah_akademik", "2025"),
    61:  ("guru_sekolah_menengah_agama", "2024"),
    62:  ("guru_sekolah_menengah_agama", "2025"),
    64:  ("guru_sekolah_menengah_cina", "2024"),
    65:  ("guru_sekolah_menengah_cina", "2025"),
    67:  ("guru_sekolah_pendidikan_khas", "2024"),
    68:  ("guru_sekolah_pendidikan_khas", "2025"),
    70:  ("guru_antarabangsa", "2024"),
    71:  ("guru_antarabangsa", "2025"),
    73:  ("guru_ekspatriat", "2024"),
    74:  ("guru_ekspatriat", "2025"),
    91:  ("bilangan_murid", "2024"), # sum var_srk_rendah + var_smk_murid
    92:  ("bilangan_murid", "2025"), # sum var_srk_rendah + var_smk_murid
    94:  ("bilangan_murid_sekolah_rendah", "2024"), # sum = murid_sekolah_rendah_akademik + murid_sekolah_rendah_agama = var_srk_rendah
    95:  ("bilangan_murid_sekolah_rendah", "2025"), # sum = murid_sekolah_rendah_akademik + murid_sekolah_rendah_agama = var_srk_rendah
    97:  ("murid_sekolah_rendah_akademik", "2024"),
    98:  ("murid_sekolah_rendah_akademik", "2025"),
    100:  ("murid_sekolah_rendah_agama", "2024"),
    101:  ("murid_sekolah_rendah_agama", "2025"),
    103:  ("bilangan_murid_sekolah_menengah", "2024"), # sum murid_sekolah_menengah_akademik + murid_sekolah_menengah_agama + murid_sekolah_menengah_cina = var_smk_murid
    104:  ("bilangan_murid_sekolah_menengah", "2025"), # sum murid_sekolah_menengah_akademik + murid_sekolah_menengah_agama + murid_sekolah_menengah_cina = var_smk_murid
    106:  ("murid_sekolah_menengah_akademik", "2024"),
    107:  ("murid_sekolah_menengah_akademik", "2025"),
    109:  ("murid_sekolah_menengah_agama", "2024"),
    110:  ("murid_sekolah_menengah_agama", "2025"),
    112:  ("murid_sekolah_menengah_cina", "2024"),
    113:  ("murid_sekolah_menengah_cina", "2025"),
    115:  ("murid_sekolah_pendidikan_khas", "2024"),
    116:  ("murid_sekolah_pendidikan_khas", "2025"),
    118:  ("murid_antarabangsa", "2024"),
    119:  ("murid_antarabangsa", "2025"),
    121:  ("murid_ekspatriat", "2024"),
    122:  ("murid_ekspatriat", "2025"),
}

# TODO: Set the furthest row down that contains dummy template data for cleanup
MAX_ROW_TO_CLEAN = 132 # Adjust based on how far down the template's dummy data goes

def populate_jadual_6_1(sheet, hierarchy, report_type):
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

    sheet["C3"] = f": Statistik pendidikan swasta, {loc_bm}, {hierarchy.get('state_name')}, 2024 - 2025"
    sheet["C4"] = f": Statistics of private education, {loc_en}, {hierarchy.get('state_name')}, 2024 - 2025"

    inject_dynamic_table(sheet, target_row, start_col, max_slots, locations, ROW_MAP, MAX_ROW_TO_CLEAN, hierarchy)