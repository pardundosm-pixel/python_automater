from src.excel_utils import get_dynamic_boundaries, inject_dynamic_table
from src.data_provider import get_metrics_dict

HEADER_ANCHOR = "KESIHATAN"
MAX_ROW_TO_CLEAN = 34 

ROW_MAP = {
    11: ("hospital_kkm", "2024"),
    14: ("hospital_swasta", "2024"),
    20: ("klinik_kesihatan_kerajaan", "2024"),
    23: ("klinik_desa", "2024"),
    26: ("klinik_komuniti", "2024"),
    29: ("bilangan_katil_kerajaan", "2024"),
    32: ("bilangan_katil_swasta", "2024"),
}

def populate_jadual_7(sheet, hierarchy, report_type):
    target_row, start_col, end_col, max_slots = get_dynamic_boundaries(sheet, HEADER_ANCHOR)
    if not start_col: return

    # 1. Build Payload
    locations = []
    if report_type == 'parlimen':
        parl = hierarchy['parl_code']
        locations.append((parl, hierarchy['parl_name'], get_metrics_dict(parl, 'parlimen')))
        for dun in hierarchy['duns']:
            locations.append((dun['code'], dun['name'], get_metrics_dict(dun['code'], 'dun', parent_code=parl)))
            
    elif report_type == 'dun':
        parent = hierarchy['parent_parl_code']
        locations.append((parent, hierarchy['parent_parl_name'], get_metrics_dict(parent, 'parlimen')))
        dun = hierarchy['dun_code']
        locations.append((dun, hierarchy['dun_name'], get_metrics_dict(dun, 'dun', parent_code=parent)))
    
    # 2. Set Titles
    loc_bm = f"Parlimen {hierarchy.get('parl_name')}" if report_type == 'parlimen' else f"DUN {hierarchy.get('dun_name')}"
    loc_en = f"Parliament {hierarchy.get('parl_name')}" if report_type == 'parlimen' else f"DUN {hierarchy.get('dun_name')}"

    sheet["C3"] = f": Statistik kemudahan kesihatan, {loc_bm}, {hierarchy.get('state_name')}, 2024"
    sheet["C4"] = f": Statistics of health facilities, {loc_en}, {hierarchy.get('state_name')}, 2024"

    # 3. Inject Data via Utility
    inject_dynamic_table(
        sheet=sheet, target_row=target_row, start_col=start_col, max_slots=max_slots,
        locations_to_inject=locations, row_map=ROW_MAP, max_row_clean=MAX_ROW_TO_CLEAN, hierarchy=hierarchy
    )