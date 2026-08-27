import pandas as pd
import re
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
def populate_jadual_6_1(sheet, hierarchy, report_type):
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

    # 5. Inject Data Flush-Right (VECTORIZED)
        headers = []
        bold_indices = []
    
        # --- A. Prepare and Inject Headers in One Call ---
        for i, (loc_code, loc_name, metrics_data) in enumerate(locations_to_inject):
            headers.append(f"{loc_code}\n{loc_name}")
            
            # Track which columns belong to a Parliament to bold them later
            if loc_code == hierarchy.get('parl_code') or loc_code == hierarchy.get('parent_parl_code'):
                bold_indices.append(i)
    
        # Inject the entire header row at once
        sheet.range((target_row, actual_start_col)).value = headers
        
        # Apply bold styling only to Parliament headers
        for i in bold_indices:
            sheet.range((target_row, actual_start_col + i)).font.bold = True
    
        # --- B. Prepare and Inject Data Rows in One Call Per Row ---
        for row_idx, (metric_name, year) in ROW_MAP.items():
            row_payload = []
            
            # Build the horizontal list for this specific metric
            for loc_code, loc_name, metrics_data in locations_to_inject:
                year_data = metrics_data.get(str(year), {})
                val = year_data.get(metric_name, "n.a")
                
                if pd.notna(val) and val != "n.a" and val != "":
                    try: 
                        val = float(val)
                    except (ValueError, TypeError): 
                        pass
                else:
                    val = "n.a"
                    
                row_payload.append(val)
                
            # 🚀 INJECT THE ENTIRE ROW ACROSS ALL COLUMNS IN 1 COM CALL
            sheet.range((row_idx, actual_start_col)).value = row_payload