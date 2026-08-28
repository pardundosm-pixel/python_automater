import os
import openpyxl
import importlib
import pkgutil
from config.settings import TEMPLATE_PATHS, OUTPUT_DIR
from src.data_provider import get_location_hierarchy, get_dun_hierarchy, get_negeri_hierarchy

# Import the base packages so pkgutil can scan them
import sheets.parlimen_dun
import sheets.malaysia
import sheets.negeri

# ==========================================
# 1. DYNAMIC PLUGIN REGISTRY BUILDER
# ==========================================
MASTER_REGISTRY = {
    "parlimen_dun": {},
    "malaysia": {},
    "negeri": {}
}

PACKAGE_MAP = {
    "parlimen_dun": sheets.parlimen_dun,
    "malaysia": sheets.malaysia,
    "negeri": sheets.negeri
}

print("[ENGINE] Booting up and discovering sheet mappers...")

for profile, package in PACKAGE_MAP.items():
    # Scan the folder for any python files (e.g., '_14_0.py')
    for _, module_name, is_pkg in pkgutil.iter_modules(package.__path__):
        if is_pkg or not module_name.startswith('_'):
            continue
            
        full_module_name = f"{package.__name__}.{module_name}"
        mod = importlib.import_module(full_module_name)
        
        # Convert file name '_14_0' to prefix '14.0'
        prefix = module_name.lstrip('_').replace('_', '.')
        
        for attr_name in dir(mod):
            if attr_name.startswith("populate_jadual_"):
                func = getattr(mod, attr_name)
                MASTER_REGISTRY[profile][prefix] = func
                break

print(f"[ENGINE] Discovery complete. Loaded {sum(len(v) for v in MASTER_REGISTRY.values())} total sheet mappers.")


# ==========================================
# 2. REPORT ORCHESTRATOR
# ==========================================
def generate_report(location_code: str, report_type: str, parent_code: str = None, template_key: str = "parlimen_dun", allowed_sheets: list = None):
    print(f"\n=============================================")
    print(f"Generating {report_type.upper()} Report [{template_key}] for: {location_code}")
    
    # --- A. HIERARCHY FETCHING ---
    if report_type == 'parlimen':
        hierarchy = get_location_hierarchy(location_code)
    elif report_type == 'dun':
        hierarchy = get_dun_hierarchy(location_code, parent_code)
    elif report_type == 'negeri':
        hierarchy = get_negeri_hierarchy(location_code)
    elif report_type == 'malaysia':
        hierarchy = {'state_name': 'Malaysia', 'location_name': 'Malaysia', 'state_code': '00'}
    else:
        print(f"  -> [Error] Invalid report_type '{report_type}'. Must be 'parlimen', 'dun', 'negeri', or 'malaysia'.")
        return
        
    if not hierarchy:
        print(f"  -> [Error] Could not find {location_code} in the dimension table.")
        return

    wb = None
    try:
        # --- B. OPEN THE CORRECT TEMPLATE ---
        target_template_path = TEMPLATE_PATHS.get(template_key)
        if not target_template_path:
            print(f"  -> [Error] Template key '{template_key}' not found in settings.py TEMPLATE_PATHS.")
            return
            
        wb = openpyxl.load_workbook(target_template_path)
        
        # --- C. GET THE CORRECT MAPPERS ---
        active_mappers = MASTER_REGISTRY.get(template_key, {})
        if not active_mappers:
            print(f"  -> [Warning] No mappers defined in MASTER_REGISTRY for '{template_key}'.")
            
        sorted_prefixes = sorted(active_mappers.keys(), key=len, reverse=True)
        
        # --- D. DYNAMIC ROUTING & CLI OVERRIDE ---
        populated_sheets = set()
    
        for sheet in wb.worksheets:
            sheet_name_clean = str(sheet.title).strip()
            for prefix in sorted_prefixes:
                if sheet_name_clean.startswith(prefix):
                    
                    # Intercept via CLI allowlist
                    if allowed_sheets and prefix not in allowed_sheets:
                        print(f"  -> Skipping {prefix} (Not in --sheets allowlist)")
                        break
                    
                    mapper_function = active_mappers[prefix]
                    mapper_function(sheet, hierarchy, report_type)
                    populated_sheets.add(sheet.title)
                    break

        # --- E. UNUSED SHEET PURGING ---
        print(f"  -> Purging {len(wb.worksheets) - len(populated_sheets)} unused sheets...")
        sheets_to_delete = [sheet.title for sheet in wb.worksheets if sheet.title not in populated_sheets]
        
        if len(wb.worksheets) > len(sheets_to_delete):
            for sheet_name in sheets_to_delete:
                del wb[sheet_name]

        # --- F. CONSTRUCT OUTPUT DIRECTORY & SAVE ---
        state_name = str(hierarchy.get('state_name', 'Unknown_State')).strip()
        safe_state_name = state_name.replace('/', '_').replace('\\', '_')
        state_code = str(hierarchy.get('state_code', '00')).strip().zfill(2)
        
        base_output_dir = os.path.abspath(OUTPUT_DIR)
        
        if template_key == "parlimen_dun":
            category_folder = os.path.join(base_output_dir, "Jadual 1 - 13 (Parlimen & DUN)")
        elif template_key == "malaysia":
            category_folder = os.path.join(base_output_dir, "Jadual Malaysia")
        elif template_key == "negeri":
            category_folder = os.path.join(base_output_dir, "Jadual Negeri")
        else:
            category_folder = os.path.join(base_output_dir, f"Lain-lain ({template_key})")

        if report_type == 'parlimen':
            parl_code = str(hierarchy.get('parl_code', '')).strip()
            parl_name = str(hierarchy.get('parl_name', '')).strip()
            target_dir = os.path.join(category_folder, safe_state_name, "Parlimen")
            file_name = f"{parl_code} {parl_name}.xlsx"
        elif report_type == 'dun':
            dun_code = str(hierarchy.get('dun_code', '')).strip()
            dun_name = str(hierarchy.get('dun_name', '')).strip()
            target_dir = os.path.join(category_folder, safe_state_name, "DUN")
            file_name = f"{state_code}_{dun_code} {dun_name}.xlsx"
        elif report_type == 'negeri':
            target_dir = category_folder
            file_name = f"Jadual_42_56_{safe_state_name}.xlsx"
        elif report_type == 'malaysia':
            target_dir = category_folder
            file_name = "Jadual_14_41_Malaysia.xlsx"
        
        os.makedirs(target_dir, exist_ok=True)
        save_path = os.path.realpath(os.path.join(target_dir, file_name))
        
        wb.save(save_path)
        wb.close()
        print(f"  -> Success! Saved to:\n     {save_path}")

    except Exception as e:
        print(f"  -> [Critical Error]: {e}")
        if wb:
            try:
                wb.close()
            except Exception:
                pass