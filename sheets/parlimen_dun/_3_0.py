import pandas as pd
import re
from src.data_provider import get_metrics_dict

# ==========================================
# 1. MAPPING CONFIGURATION 
# ==========================================
# TODO: Set the unique text in Column B to find this table's starting row
HEADER_ANCHOR = "PERUMAHAN" # e.g., "MAKLUMAT ASAS" or "KEMISKINAN"

# TODO: Map the EXCEL ROW NUMBER to a TUPLE of ("METRIC_NAME", "YEAR")
# Format: { Excel_Row: ("metric_name_in_database", "Year") }
ROW_MAP = {
    11:  ("dimiliki", "2019"),
    12:  ("dimiliki", "2022"),
    14:  ("disewa", "2019"),
    15:  ("disewa", "2022"),
    17: ("kuarters", "2019"),
    18: ("kuarters", "2022"),
}

# TODO: Set the furthest row down that contains dummy template data for cleanup
MAX_ROW_TO_CLEAN = 50 # Adjust based on how far down the template's dummy data goes

# ==========================================
# 2. DYNAMIC BOUNDARY SCANNER (Upgraded)
# ==========================================
def get_dynamic_boundaries(sheet, header_text):
    target_anchor_row = None
    
    # Step A: Find the Target Title Row in Column B (Index 2)
    for row_idx in range(1, 100): 
        val = sheet.range((row_idx, 2)).value
        if val and isinstance(val, str) and header_text.lower() in val.lower():
            target_anchor_row = row_idx
            break
            
    if not target_anchor_row:
        return None, None, None, None
        
    start_col = None
    end_col = None
    target_header_row = target_anchor_row
    
    # Step B: Scan horizontally for P.XXX and N.XXX (Check title row + up to 3 rows below)
    for offset in range(0, 4):
        search_row = target_anchor_row + offset
        
        for col_idx in range(3, 30):
            val = sheet.range((search_row, col_idx)).value
            if val and isinstance(val, str):
                clean_val = val.strip()
                
                # Detect Parliament signature
                if re.match(r"^P\.\d+", clean_val, re.IGNORECASE) and start_col is None:
                    start_col = col_idx
                    target_header_row = search_row # Lock in the EXACT row the headers live on
                    
                # Detect DUN signature
                if re.match(r"^N\.\d+", clean_val, re.IGNORECASE):
                    end_col = col_idx 
                    
        # If we found the start_col on this row, stop searching downwards
        if start_col:
            break
            
    # Edge Case: Table only has a Parliament column
    if start_col and not end_col:
        end_col = start_col
        
    if start_col and end_col:
        max_slots = (end_col - start_col) + 1
        # Return the target_header_row so the engine writes headers to the correct level
        return target_header_row, start_col, end_col, max_slots
        
    return target_anchor_row, None, None, None

# ==========================================
# 3. REPORT INJECTION ENGINE
# ==========================================
# TODO: Rename the function to match the specific jadual (e.g., populate_jadual_3_0)
def populate_jadual_3(sheet, hierarchy, report_type):
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
        loc_string_en = f"Parliament {hierarchy.get('parl_name')}"
    else:
        loc_string_bm = f"DUN {hierarchy.get('dun_name')}"
        loc_string_en = f"DUN {hierarchy.get('dun_name')}"

    # TODO: Format the specific base text titles for this Jadual
    title_bm = f": Statistik isi rumah mengikut jenis rumah yang didiami, {loc_string_bm}, {hierarchy.get('state_name')}, 2019 - 2022"
    title_en = f": Statistics of household by type of occupied dwelling, {loc_string_en}, {hierarchy.get('state_name')}, 2019 - 2022"

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