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
    13: ("perkhidmatan", "2025"),
    16: ("bekalan_elektrik_gas_wap_dan_pendinginan_udara", "2025"),
    19: ("bekalan_air_pembentungan_pengurusan_sisa_dan_aktiviti_pemulihan", "2025"),
    24: ("perdagangan_borong_dan_runcit_pembaikan_kenderaan_bermotor", "2025"),
    29: ("pembaikan_dan_penyelenggaraan_motosikal_bengkel_motorsikal", "2025"),
    34: ("lain_lain_aktiviti_di_bawah_perdagangan_borong_dan_runcit_pembaikan_kenderaan_bermotor", "2025"),
    37: ("pengangkutan_dan_penyimpanan", "2025"),
    40: ("penginapan_dan_aktiviti_perkhidmatan_makanan_dan_minuman", "2025"),
    44: ("hotel_dan_hotel_resort", "2025"),
    47: ("motel", "2025"),
    50: ("homestay", "2025"),
    53: ("gerai_penjaja_makanan", "2025"),
    56: ("makanan_dan_minuman_penyediaan_makanan_dan_minuman_di_dalam_gerai_penjaja", "2025"),
    61: ("gerai_penjaja_minuman", "2025"),
    64: ("restoran_dan_restoran_yang_juga_kelab_malam", "2025"),
    67: ("restoran_makanan_segera", "2025"),
    70: ("lain_lain_aktiviti_di_bawah_penginapan_dan_aktiviti_perkhidmatan_makanan_dan_minuman", "2025"),
}

# TODO: Set the furthest row down that contains dummy template data for cleanup
MAX_ROW_TO_CLEAN = 74 # Adjust based on how far down the template's dummy data goes

# ==========================================
# 3. REPORT INJECTION ENGINE
# ==========================================
# TODO: Rename the function to match the specific jadual (e.g., populate_jadual_3_0)
def populate_jadual_12_1(sheet, hierarchy, report_type):
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