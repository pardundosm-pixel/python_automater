from src.excel_utils import get_dynamic_boundaries, inject_dynamic_table
from src.data_provider import get_metrics_dict

# ==========================================
# 1. MAPPING CONFIGURATION 
# ==========================================
# Set the unique text in Column B to find this table's starting row
HEADER_ANCHOR = "MAKLUMAT ASAS"

# ==========================================
# MAPPING CONFIGURATION FOR JADUAL 1.0
# Format: { Excel_Row: ("Metric_Name", "Year") }
# ==========================================
ROW_MAP = {
    8:  ("luas_kawasan", "2023"),
    9:  ("luas_kawasan", "2024"),
    10: ("luas_kawasan", "2025"),
    12: ("jumlah_penduduk", "2023"),
    13: ("jumlah_penduduk", "2024"),
    14: ("jumlah_penduduk", "2025"),
    16: ("kepadatan_penduduk", "2023"),
    17: ("kepadatan_penduduk", "2024"),
    18: ("kepadatan_penduduk", "2025"),
    20: ("jumlah_pemilih", "PRU 14"), 
    21: ("jumlah_pemilih", "PRU 15"), 
    23: ("jumlah_undian_oleh_pemilih", "PRU 14"), 
    24: ("jumlah_undian_oleh_pemilih", "PRU 15"),
    26: ("jumlah_kertas_undi_yang_ditolak", "PRU 14"), #variable 
    27: ("jumlah_kertas_undi_yang_ditolak", "PRU 15"), #variable baru
    29: ("bilangan_kertas_undi_yang_dikeluarkan_tetapi_tidak_dimasukkan_ke_dalam_peti_undi_dan_tidak_dikembalikan", "PRU 14"), #variable baru
    30: ("bilangan_kertas_undi_yang_dikeluarkan_tetapi_tidak_dimasukkan_ke_dalam_peti_undi_dan_tidak_dikembalikan", "PRU 15"), #variable baru
    32: ("jumlah_kertas_undi_yang_dikeluarkan_dalam_pengundian", "PRU 14"), #variable baru
    33: ("jumlah_kertas_undi_yang_dikeluarkan_dalam_pengundian", "PRU 15") #variable baru
}

# The furthest row down that contains dummy template data for cleanup
MAX_ROW_TO_CLEAN = 33 

def populate_jadual_1(sheet, hierarchy, report_type):
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

    sheet["C3"] = f": Statistik maklumat asas, {loc_bm}, {hierarchy.get('state_name')}, 2023 - 2025"
    sheet["C4"] = f": Statistics of basic information, {loc_en}, {hierarchy.get('state_name')}, 2023 - 2025"

    inject_dynamic_table(sheet, target_row, start_col, max_slots, locations, ROW_MAP, MAX_ROW_TO_CLEAN, hierarchy)