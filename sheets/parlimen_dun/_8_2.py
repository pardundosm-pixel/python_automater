import pandas as pd
import re
from src.data_provider import get_metrics_dict

# ==========================================
# 1. MAPPING CONFIGURATION 
# ==========================================
# TODO: Set the unique text in Column B to find this table's starting row
HEADER_ANCHOR = "KEMISKINAN" # e.g., "MAKLUMAT ASAS" or "KEMISKINAN"

# TODO: Map the EXCEL ROW NUMBER to a TUPLE of ("METRIC_NAME", "YEAR")
# Format: { Excel_Row: ("metric_name_in_database", "Year") }
ROW_MAP = {
    10: ("jumlah_keseluruhan_agama", "2024"),
    11: ("jumlah_keseluruhan_agama", "2025"),
    13: ("jumlah_keseluruhan_miskin_tegar", "2024"),
    14: ("jumlah_keseluruhan_miskin_tegar", "2025"),
    16: ("jumlah_miskin_tegar_islam", "2024"),
    17: ("jumlah_miskin_tegar_islam", "2025"),
    19: ("jumlah_miskin_tegar_kristian", "2024"),
    20: ("jumlah_miskin_tegar_kristian", "2025"),
    22: ("jumlah_miskin_tegar_buddha", "2024"),
    23: ("jumlah_miskin_tegar_buddha", "2025"),
    25: ("jumlah_miskin_tegar_hindu", "2024"),
    26: ("jumlah_miskin_tegar_hindu", "2025"),
    28: ("jumlah_miskin_tegar_lain_lain", "2024"),
    29: ("jumlah_miskin_tegar_lain_lain", "2025"),
    31: ("jumlah_keseluruhan_miskin", "2024"),
    32: ("jumlah_keseluruhan_miskin", "2025"),
    34: ("jumlah_miskin_islam", "2024"),
    35: ("jumlah_miskin_islam", "2025"),
    37: ("jumlah_miskin_kristian", "2024"),
    38: ("jumlah_miskin_kristian", "2025"),
    40: ("jumlah_miskin_buddha", "2024"),
    41: ("jumlah_miskin_buddha", "2025"),
    43: ("jumlah_miskin_hindu", "2024"),
    44: ("jumlah_miskin_hindu", "2025"),
    46: ("jumlah_miskin_lain_lain", "2024"),
    47: ("jumlah_miskin_lain_lain", "2025"),
}

# TODO: Set the furthest row down that contains dummy template data for cleanup
MAX_ROW_TO_CLEAN = 50 # Adjust based on how far down the template's dummy data goes

# ==========================================
# 2. DYNAMIC BOUNDARY SCANNER (Universal)
# ==========================================
def get_dynamic_boundaries(sheet, header_text):
    """
    Scans the Excel sheet to dynamically find the Target Row, Start Column, 
    End Column, and Total Available Slots based on P.XXX and N.XXX signatures.
    """
    target_row = None
    
    # Step A: Find the Target Row in Column B (Index 2)
    for row_idx in range(1, 100):  # Scan first 100 rows safely
        val = sheet.range((row_idx, 2)).value
        if val and isinstance(val, str) and header_text.lower() in val.lower():
            target_row = row_idx
            break
            
    if not target_row:
        return None, None, None, None
        
    start_col = None
    end_col = None
    
    # Step B: Scan horizontally across the Target Row to find P.XXX and N.XXX
    for col_idx in range(3, 30):  # Scan across 30 columns safely
        val = sheet.range((target_row, col_idx)).value
        if val and isinstance(val, str):
            clean_val = val.strip()
            
            # 1. Detect Parliament signature (e.g., "P.037\nKULIM...")
            if re.match(r"^P\.\d+", clean_val, re.IGNORECASE) and start_col is None:
                start_col = col_idx
                
            # 2. Detect DUN signature (e.g., "N.18\nSimanggang")
            if re.match(r"^N\.\d+", clean_val, re.IGNORECASE):
                end_col = col_idx # This overwrites until it finds the VERY LAST DUN
                
    # Edge Case: Table only has a Parliament column and no DUNs drawn
    if start_col and not end_col:
        end_col = start_col
        
    if start_col and end_col:
        max_slots = (end_col - start_col) + 1
        return target_row, start_col, end_col, max_slots
        
    return target_row, None, None, None

# ==========================================
# 3. REPORT INJECTION ENGINE
# ==========================================
# TODO: Rename the function to match the specific jadual (e.g., populate_jadual_3_0)
def populate_jadual_8_2(sheet, hierarchy, report_type):
    loc_name_debug = hierarchy.get('parl_code') if report_type == 'parlimen' else hierarchy.get('dun_code')
    print(f"  -> Populating Jadual for {loc_name_debug} | Anchor: '{HEADER_ANCHOR}'")

    # 1. Dynamically locate the grid coordinates
    target_row, start_col, end_col, max_slots = get_dynamic_boundaries(sheet, HEADER_ANCHOR)
    
    if not start_col:
        print(f"     [Error] Could not locate P.XXX boundaries for '{HEADER_ANCHOR}'. Skipping.")
        return

    # 2. Build the Data Payload Sequence 
    locations_to_inject = []
    
    if report_type == 'parlimen':
        parl_code = hierarchy['parl_code']
        parl_metrics = get_metrics_dict(parl_code, 'parlimen')
        locations_to_inject.append((parl_code, hierarchy['parl_name'], parl_metrics))
        
        for dun in hierarchy['duns']:
            dun_metrics = get_metrics_dict(dun['code'], 'dun', parent_code=parl_code)
            locations_to_inject.append((dun['code'], dun['name'], dun_metrics))
            
    elif report_type == 'dun':
        parent_code = hierarchy['parent_parl_code']
        parl_metrics = get_metrics_dict(parent_code, 'parlimen')
        locations_to_inject.append((parent_code, hierarchy['parent_parl_name'], parl_metrics))
        
        dun_code = hierarchy['dun_code']
        dun_metrics = get_metrics_dict(dun_code, 'dun', parent_code=parent_code)
        locations_to_inject.append((dun_code, hierarchy['dun_name'], dun_metrics))
    
    # ==========================================
    # DYNAMIC TABLE TITLE MODIFICATION
    # ==========================================
    # 1. Determine the correct location string based on the report type
    if report_type == 'parlimen':
        loc_string_bm = f"Parlimen {hierarchy.get('parl_name')}"
        loc_string_en = f"Parlimen {hierarchy.get('parl_name')}"
    else:
        loc_string_bm = f"DUN {hierarchy.get('dun_name')}"
        loc_string_en = f"DUN {hierarchy.get('dun_name')}"

    # TODO: Format the specific base text titles for this Jadual
    title_bm = f": Tajuk BM di sini bagi {loc_string_bm}, {hierarchy.get('state_name')}, 2022 - 2024"
    title_en = f": English Title here for {loc_string_en}, {hierarchy.get('state_name')}, 2022 - 2024"

    # TODO: Inject into the sheet (Choose ONE method and comment out the other)
    
    # Method A: Fixed cell coordinates (if the title never moves rows)
    # sheet.range("C3").value = title_bm
    # sheet.range("C4").value = title_en
    
    # Method B: Dynamic coordinates based on Target Row (if the title shifts with the table)
    # sheet.range((target_row - 2, 3)).value = title_bm  # Assuming Column C is Index 3
    # sheet.range((target_row - 1, 3)).value = title_en
    # ==========================================

    # Safety constraint: Never exceed the available slots dynamically detected
    payload = len(locations_to_inject)
    if payload > max_slots:
        payload = max_slots
        locations_to_inject = locations_to_inject[:max_slots]

    # 3. Execute Right-Aligned Math
    skip_offset = max_slots - payload
    actual_start_col = start_col + skip_offset

    # 4. Left-Side Template Cleanup
    # Wipe from the first dummy column up to the actual starting column
    if skip_offset > 0:
        # Erase Headers (P.XXX / N.XXX row)
        sheet.range((target_row, start_col), (target_row, actual_start_col - 1)).value = None
        # Erase Data Rows underneath
        sheet.range((target_row + 1, start_col), (MAX_ROW_TO_CLEAN, actual_start_col - 1)).value = None

    # 5. Inject Data Flush-Right
    current_col = actual_start_col
    for loc_code, loc_name, metrics_data in locations_to_inject:
        
        # Write Column Header dynamically
        header_cell = sheet.range((target_row, current_col))
        header_cell.value = f"{loc_code}\n{loc_name}"
        
        # BOLD PARLIAMENT HEADERS 
        # Check if the current column belongs to the Parliament. If yes, make it bold.
        is_parliament = (loc_code == hierarchy.get('parl_code') or loc_code == hierarchy.get('parent_parl_code'))
        header_cell.font.bold = is_parliament
        
        # Inject Mapped Rows
        for row_idx, (metric_name, year) in ROW_MAP.items():
            year_data = metrics_data.get(str(year), {})
            val = year_data.get(metric_name, "n.a")
            
            if pd.notna(val) and val != "n.a" and val != "":
                try: 
                    val = float(val)
                except (ValueError, TypeError): 
                    pass
            else:
                val = "n.a"
                
            sheet.range((row_idx, current_col)).value = val
            
        current_col += 1