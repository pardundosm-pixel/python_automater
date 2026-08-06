import os
import xlwings as xw
from config.settings import TEMPLATE_PATHS, OUTPUT_DIR
from src.data_provider import get_location_hierarchy, get_dun_hierarchy

# ==========================================
# 1. IMPORT YOUR SHEET MAPPERS
# ==========================================
# (Parlimen & DUN Template Mappers)
from sheets.parlimen_dun._1 import populate_jadual_1 as jadual_1
from sheets.parlimen_dun._2 import populate_jadual_2 as jadual_2
from sheets.parlimen_dun._2_1 import populate_jadual_2_1 as jadual_2_1    
from sheets.parlimen_dun._2_2 import populate_jadual_2_2 as jadual_2_2    
from sheets.parlimen_dun._2_3 import populate_jadual_2_3 as jadual_2_3    
from sheets.parlimen_dun._3_0 import populate_jadual_3 as jadual_3
from sheets.parlimen_dun._4_0 import populate_jadual_4 as jadual_4
from sheets.parlimen_dun._5_0 import populate_jadual_5 as jadual_5
from sheets.parlimen_dun._6_0 import populate_jadual_6 as jadual_6 
from sheets.parlimen_dun._6_1 import populate_jadual_6_1 as jadual_6_1
from sheets.parlimen_dun._7_0 import populate_jadual_7 as jadual_7
from sheets.parlimen_dun._8_0 import populate_jadual_8 as jadual_8
from sheets.parlimen_dun._8_1 import populate_jadual_8_1 as jadual_8_1
from sheets.parlimen_dun._8_2 import populate_jadual_8_2 as jadual_8_2
from sheets.parlimen_dun._8_3 import populate_jadual_8_3 as jadual_8_3
from sheets.parlimen_dun._9_0 import populate_jadual_9 as jadual_9
from sheets.parlimen_dun._10_0 import populate_jadual_10 as jadual_10
from sheets.parlimen_dun._11_0 import populate_jadual_11 as jadual_11
from sheets.parlimen_dun._12_0 import populate_jadual_12 as jadual_12
from sheets.parlimen_dun._12_1 import populate_jadual_12_1 as jadual_12_1
from sheets.parlimen_dun._12_1_1 import populate_jadual_12_1_1 as jadual_12_1_1
from sheets.parlimen_dun._12_2 import populate_jadual_12_2 as jadual_12_2
from sheets.parlimen_dun._13_0 import populate_jadual_13 as jadual_13
from sheets.parlimen_dun._13_0_1 import populate_jadual_13_0_1 as jadual_13_0_1
from sheets.parlimen_dun._13_1 import populate_jadual_13_1 as jadual_13_1
from sheets.parlimen_dun._13_2 import populate_jadual_13_2 as jadual_13_2

# --- NEW: IMPORT MALAYSIA MAPPERS ---
# (Adjust the folder name 'sheets.malaysia' if you named the folder differently!)
from sheets.malaysia._14_0 import populate_jadual_14 as jadual_14
from sheets.malaysia._31_0 import populate_jadual_31 as jadual_31

# ==========================================
# 2. BUILD THE NESTED ROUTING REGISTRY
# ==========================================
MASTER_REGISTRY = {
    # Profile 1: The standard Parlimen & DUN template
    "parlimen_dun": {
        "1.0_Maklumat_Asas": jadual_1,
        "2.0_Penduduk_Malaysia": jadual_2,
        "2.1_Penduduk_Negeri": jadual_2_1,
        "2.2_Penduduk_Parlimen": jadual_2_2,
        "2.3_Penduduk": jadual_2_3,
        "3.0_Perumahan": jadual_3,
        "4.0_Guna_Tenaga": jadual_4,
        "5.0_Pendapatan": jadual_5,
        "6.0_Pendidikan_SK": jadual_6,
        "6.1_Pendidikan_Swasta": jadual_6_1,
        "7.0_Kesihatan": jadual_7,
        "8.0_Jantina": jadual_8,
        "8.1_Etnik": jadual_8_1,
        "8.2_Agama": jadual_8_2,
        "8.3_Umur": jadual_8_3,
        "9.0_Keselamatan_Awam": jadual_9,
        "10.0_IMS": jadual_10,
        "11.0_Air_Elektrik_Sampah": jadual_11,
        "12.0_Pertubuhan": jadual_12,
        "12.1_Prtubuhn_Perkhdmtn": jadual_12_1,
        "12.1_Prtubuhn_Perkhdmtn_(samb.)": jadual_12_1_1,
        "12.2_Kemudahan_Asas": jadual_12_2,
        "13.0_Fasiliti_awam": jadual_13,
        "13.0_Fasiliti_awam_samb": jadual_13_0_1,
        "13.1_Statistik_Perkhid._lain": jadual_13_1,
        "13.2_Koperasi": jadual_13_2
    },
    
    # --- Profile 2 (Malaysia) ---
    "malaysia": {
        "14.0_KDNK": jadual_14,
        "31.0_Jenayah_violent": jadual_31,
        # Add future Malaysia sheets here
    }
}

# ==========================================
# 3. REPORT ORCHESTRATOR
# ==========================================
def generate_report(location_code: str, report_type: str, excel_app: xw.App, parent_code: str = None, template_key: str = "parlimen_dun"):
    print(f"\n=============================================")
    print(f"Generating {report_type.upper()} Report [{template_key}] for: {location_code}")
    
    # --- A. NEW HIERARCHY FETCHING ---
    if report_type == 'parlimen':
        hierarchy = get_location_hierarchy(location_code)
    elif report_type == 'dun':
        hierarchy = get_dun_hierarchy(location_code, parent_code)
    elif report_type == 'malaysia':
        # Malaysia doesn't need to query the dim_geo table, we just build a dummy dictionary
        hierarchy = {'state_name': 'Malaysia', 'location_name': 'Malaysia', 'state_code': '00'}
    else:
        print(f"  -> [Error] Invalid report_type '{report_type}'. Must be 'parlimen', 'dun', or 'malaysia'.")
        return
        
    if not hierarchy:
        print(f"  -> [Error] Could not find {location_code} in the dimension table.")
        return

    try:
        # --- B. OPEN THE CORRECT TEMPLATE ---
        target_template_path = TEMPLATE_PATHS.get(template_key)
        if not target_template_path:
            print(f"  -> [Error] Template key '{template_key}' not found in settings.py TEMPLATE_PATHS.")
            return
            
        wb = excel_app.books.open(target_template_path)
        
        # --- C. GET THE CORRECT MAPPERS FOR THIS TEMPLATE ---
        active_mappers = MASTER_REGISTRY.get(template_key, {})
        if not active_mappers:
            print(f"  -> [Warning] No mappers defined in MASTER_REGISTRY for '{template_key}'.")
            
        sorted_prefixes = sorted(active_mappers.keys(), key=len, reverse=True)
        
        # --- D. DYNAMIC ROUTING ---
        for sheet in wb.sheets:
            sheet_name_clean = str(sheet.name).strip()
            
            for prefix in sorted_prefixes:
                if sheet_name_clean.startswith(prefix):
                    mapper_function = active_mappers[prefix]
                    mapper_function(sheet, hierarchy, report_type)
                    break 

        # --- E. CONSTRUCT OUTPUT DIRECTORY & SAVE LOGIC ---
        state_name = str(hierarchy.get('state_name', 'Unknown_State')).strip()
        safe_state_name = state_name.replace('/', '_').replace('\\', '_')
        
        if report_type == 'parlimen':
            parl_code = str(hierarchy.get('parl_code', '')).strip()
            parl_name = str(hierarchy.get('parl_name', '')).strip()
            target_dir = os.path.join(os.path.abspath(OUTPUT_DIR), safe_state_name, "Parlimen")
            file_name = f"{parl_code} {parl_name} - {template_key}.xlsx"
            
        elif report_type == 'dun':
            dun_code = str(hierarchy.get('dun_code', '')).strip()
            dun_name = str(hierarchy.get('dun_name', '')).strip()
            parent_code_hierarchy = str(hierarchy.get('parent_parl_code', '')).strip()
            parent_name = str(hierarchy.get('parent_parl_name', '')).strip()
            safe_parent_folder = f"{parent_code_hierarchy}_{parent_name}".replace('/', '_').replace('\\', '_')
            target_dir = os.path.join(os.path.abspath(OUTPUT_DIR), safe_state_name, "DUN", safe_parent_folder)
            file_name = f"{dun_code} {dun_name} - {template_key}.xlsx"
            
        elif report_type == 'malaysia':
            # NEW: Save Malaysia reports directly into the root Output folder
            target_dir = os.path.abspath(OUTPUT_DIR)
            file_name = f"MALAYSIA Profile - {template_key}.xlsx"
        
        os.makedirs(target_dir, exist_ok=True)
        save_path = os.path.join(target_dir, file_name)
        
        # --- Save Report and Close Workbook ---
        wb.save(save_path)
        wb.close()
        print(f"  -> Success! Saved to:\n     {save_path}")

    except Exception as e:
        print(f"  -> [Critical Error]: {e}")
        try:
            wb.close()
        except:
            pass