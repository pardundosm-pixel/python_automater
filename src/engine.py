import os
import xlwings as xw
from config.settings import TEMPLATE_PATHS, OUTPUT_DIR
from src.data_provider import get_location_hierarchy, get_dun_hierarchy, get_negeri_hierarchy

# ==========================================
# 1. IMPORT YOUR SHEET MAPPERS
# ==========================================
# (Parlimen & DUN Template Mappers)
from sheets.parlimen_dun._1 import populate_jadual_1 as jadual_1
from sheets.parlimen_dun._2_0 import populate_jadual_2 as jadual_2
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

# (Malaysia Template Mappers)
from sheets.malaysia._14_0 import populate_jadual_14_0 as jadual_14_0_malaysia
from sheets.malaysia._14_1 import populate_jadual_14_1 as jadual_14_1_malaysia
from sheets.malaysia._14_2 import populate_jadual_14_2 as jadual_14_2_malaysia
from sheets.malaysia._15_0 import populate_jadual_15_0 as jadual_15_0_malaysia
from sheets.malaysia._16_0 import populate_jadual_16_0 as jadual_16_0_malaysia
from sheets.malaysia._17_0 import populate_jadual_17_0 as jadual_17_0_malaysia
from sheets.malaysia._18_0 import populate_jadual_18_0 as jadual_18_0_malaysia
from sheets.malaysia._19_0 import populate_jadual_19_0 as jadual_19_0_malaysia

# (Negeri Template Mappers)
from sheets.negeri._42_1 import populate_jadual_42_1 as jadual_42_1_negeri
from sheets.negeri._42_0 import populate_jadual_42 as jadual_42_negeri
from sheets.negeri._43_0 import populate_jadual_43 as jadual_43_negeri

# ==========================================
# 2. BUILD THE NESTED ROUTING REGISTRY
# ==========================================
MASTER_REGISTRY = {
    # Profile 1: The standard Parlimen & DUN template
    "parlimen_dun": {
        # "1.0_Maklumat_Asas": jadual_1,
        "2.0_Penduduk_Malaysia": jadual_2,
        "2.1_Penduduk_Negeri": jadual_2_1,
        "2.2_Penduduk_Parlimen": jadual_2_2,
        "2.3_Penduduk": jadual_2_3,
        # "3.0_Perumahan": jadual_3,
        # "4.0_Guna_Tenaga": jadual_4,
        # "5.0_Pendapatan": jadual_5,
        # "6.0_Pendidikan_SK": jadual_6,
        # "6.1_Pendidikan_Swasta": jadual_6_1,
        # "7.0_Kesihatan": jadual_7,
        # "8.0_Jantina": jadual_8,
        # "8.1_Etnik": jadual_8_1,
        # "8.2_Agama": jadual_8_2,
        # "8.3_Umur": jadual_8_3,
        # "9.0_Keselamatan_Awam": jadual_9,
        # "10.0_IMS": jadual_10,
        # "11.0_Air_Elektrik_Sampah": jadual_11,
        # "12.0_Pertubuhan": jadual_12,
        # "12.1_Prtubuhn_Perkhdmtn": jadual_12_1,
        # "12.1_Prtubuhn_Perkhdmtn_(samb.)": jadual_12_1_1,
        # "12.2_Kemudahan_Asas": jadual_12_2,
        # "13.0_Fasiliti_awam": jadual_13,
        # "13.0_Fasiliti_awam_samb": jadual_13_0_1,
        # "13.1_Statistik_Perkhid._lain": jadual_13_1,
        # "13.2_Koperasi": jadual_13_2
    },
    
    # Profile 2: Malaysia
    "malaysia": {
        "14.0_KDNK": jadual_14_0_malaysia,
        "14.0_KDNK_1": jadual_14_1_malaysia,
        "14.0_KDNK_2": jadual_14_2_malaysia,
        "15.0_Pasaran_Kewagn": jadual_15_0_malaysia,
        "16.0_Akaun_ SP": jadual_16_0_malaysia,
        "17.0_Akaun_ STMK": jadual_17_0_malaysia,
        "18.0_Stok_Modal": jadual_18_0_malaysia,
        "19.0_Imbagn_Pembayaran": jadual_19_0_malaysia,
        # "14.0_KDNK_1": jadual_14_1_malaysia
        # "20.0": jadual_20_malaysia,
        # "31.0": jadual_31_malaysia,
    },
    
    # Profile 3: Negeri
    "negeri": {
        "42.1_KDNK": jadual_42_1_negeri,
        "42.0_KDNK": jadual_42_negeri,
        "43.0_Dagangan": jadual_43_negeri
        # Future Negeri Mappers will go here
    }
}


# ==========================================
# 3. REPORT ORCHESTRATOR
# ==========================================
import sheets  # requires sheets/__init__.py

LOCATION_REGISTRY: dict[int, callable] = {}   # jadual → populate_column
YEAR_REGISTRY:     dict[int, callable] = {}   # jadual → populate_page

print("[ENGINE] Discovering jadual mappers...")
for _, module_name, is_pkg in pkgutil.iter_modules(sheets.__path__):
    if is_pkg or not module_name.startswith("jadual_"):
        continue

    mod = importlib.import_module(f"sheets.{module_name}")

    # Skip placeholder / unfinished mappers
    if not hasattr(mod, "JADUAL_NUMBER"):
        print(f"  [skip] {module_name}: JADUAL_NUMBER not defined")
        continue
    if mod.JADUAL_NUMBER is None:
        print(f"  [skip] {module_name}: JADUAL_NUMBER is None (not configured)")
        continue

    num = str(mod.JADUAL_NUMBER)

    if hasattr(mod, "populate_column"):
        LOCATION_REGISTRY[num] = mod.populate_column
        print(f"  loaded sheets/{module_name}.py  →  Jadual {num}  [location-mode]")
    elif hasattr(mod, "populate_page"):
        YEAR_REGISTRY[num] = mod.populate_page
        print(f"  loaded sheets/{module_name}.py  →  Jadual {num}  [year-mode]")
    else:
        print(f"  -> [Error] Invalid report_type '{report_type}'. Must be 'parlimen', 'dun', 'negeri', or 'malaysia'.")
        return
        
    if not hierarchy:
        print(f"  -> [Error] Could not find {location_code} in the dimension table.")
        return

print(f"[ENGINE] {len(LOCATION_REGISTRY)} location-mapper(s), "
      f"{len(YEAR_REGISTRY)} year-mapper(s) registered.\n")


# ==========================================
# 2. HEADER / PAGE HELPERS
# ==========================================
def detect_location_start_col(ws, header_row: int = KATEGORI_ROW) -> int:
    """
    For location-mode jaduals: if column C's header contains 'Tahun'/'Year',
    the jadual has a year column and locations start at column D.
    Otherwise locations start at C.
    """
    c_val = ws.cell(row=header_row, column=COL_C).value
    has_year = c_val is not None and "tahun" in str(c_val).lower()
    return COL_D if has_year else COL_C


def parse_header_locations(ws,
                           header_row: int = KATEGORI_ROW,
                           start_col: int | None = None,
                           end_col: int = COL_L) -> list[dict]:
    """
    Walk one page's header row and return ordered location dicts:
        {col, kind, kod, nama, parent_parlimen}
    kind ∈ {'negeri', 'parlimen', 'dun'}
    """
    if start_col is None:
        start_col = detect_location_start_col(ws, header_row)

    locations: list[dict] = []
    current_parlimen: str | None = None

    for col in range(start_col, end_col + 1):
        val = ws.cell(row=header_row, column=col).value
        if val is None:
            continue
        text = str(val).strip()
        if not text:
            continue

        # "P.140\nSegamat"  →  kod="P.140", nama="Segamat"
        if "\n" in text:
            kod, nama = text.split("\n", 1)
            kod, nama = kod.strip(), nama.strip()
        else:
            kod, nama = None, text

        if kod and kod.upper().startswith("P."):
            current_parlimen = kod
            locations.append({
                "col": col, "kind": "parlimen",
                "kod": kod, "nama": nama,
                "parent_parlimen": None,
            })
        elif kod and kod.upper().startswith("N."):
            locations.append({
                "col": col, "kind": "dun",
                "kod": kod, "nama": nama,
                "parent_parlimen": current_parlimen,
            })
        else:
            # Fallback: no prefix ⇒ state cell ("Johor")
            locations.append({
                "col": col, "kind": "negeri",
                "kod": None, "nama": nama,
                "parent_parlimen": None,
            })

    return locations


def parse_header_years(ws,
                       header_row: int = KATEGORI_ROW,
                       start_col: int = COL_C,
                       end_col: int = COL_L) -> dict[int, str]:
    """
    For year-mode jaduals: walk the header row and return {col_idx: year_str}.
    Skips blanks. Starts scanning at column C by default.
    """
    col_map: dict[int, str] = {}
    for col in range(start_col, end_col + 1):
        val = ws.cell(row=header_row, column=col).value
        if val is None:
            continue
        text = str(val).strip()
        if text:
            col_map[col] = text
    return col_map


def detect_page_header_rows(ws) -> list[int]:
    """
    Return absolute row numbers of every page-header in the sheet.

    Page 1 header is always at KATEGORI_ROW. Later pages land at
    KATEGORI_ROW + N * ROWS_PER_PAGE (per Stage-1's stacking rule).
    Stop as soon as a candidate row has an empty column-B label.
    """
    headers = [KATEGORI_ROW]
    page = 1
    while True:
        test_row = KATEGORI_ROW + page * ROWS_PER_PAGE
        if test_row > ws.max_row:
            break
        val = ws.cell(row=test_row, column=COL_B).value
        if val is None or not str(val).strip():
            break
        headers.append(test_row)
        page += 1
    return headers


def extract_jadual_number(filename: str) -> str | None:
    """'Jadual 7.xlsx' → '7'. 'Jadual 2.1.xlsx' → '2.1'."""
    # Capture digits, an optional dot, and more digits
    m = re.search(r"Jadual[\s_\-]+(\d+(?:\.\d+)?)", filename, re.IGNORECASE)
    return m.group(1) if m else None


# ==========================================
# 3. METRIC LOOKUP
# ==========================================
def fetch_metrics_for_location(loc: dict, state_code: str) -> dict:
    """Return {year: {metric: value}} for one location, or {} on failure."""
    kind = loc["kind"]
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
        state_code = str(hierarchy.get('state_code', '00')).strip().zfill(2)
        
        base_output_dir = os.path.abspath(OUTPUT_DIR)
        
        # 1. Determine Root Category Folder based on template profile
        if template_key == "parlimen_dun":
            category_folder = os.path.join(base_output_dir, "Jadual 1 - 13 (Parlimen & DUN)")
        elif template_key == "malaysia":
            category_folder = os.path.join(base_output_dir, "Jadual Malaysia")
        elif template_key == "negeri":
            category_folder = os.path.join(base_output_dir, "Jadual Negeri")
        else:
            category_folder = os.path.join(base_output_dir, f"Lain-lain ({template_key})")

        # 2. Determine Subfolders and Naming
        if report_type == 'parlimen':
            parl_code = str(hierarchy.get('parl_code', '')).strip()
            parl_name = str(hierarchy.get('parl_name', '')).strip()
            
            # e.g., output/Jadual 1 - 13 (Parlimen & DUN)/Johor/Parlimen/P.143 Pagoh.xlsx
            target_dir = os.path.join(category_folder, safe_state_name, "Parlimen")
            file_name = f"{parl_code} {parl_name}.xlsx"
            
        elif report_type == 'dun':
            dun_code = str(hierarchy.get('dun_code', '')).strip()
            dun_name = str(hierarchy.get('dun_name', '')).strip()
            
            # e.g., output/Jadual 1 - 13 (Parlimen & DUN)/Johor/DUN/01_N.07 Bukit Kepong.xlsx
            target_dir = os.path.join(category_folder, safe_state_name, "DUN")
            file_name = f"{state_code}_{dun_code} {dun_name}.xlsx"
            
        elif report_type == 'negeri':
            # e.g., output/Jadual Negeri/Jadual_42_56_Johor.xlsx
            target_dir = category_folder
            file_name = f"Jadual_42_56_{safe_state_name}.xlsx"
            
        elif report_type == 'malaysia':
            # e.g., output/Jadual Malaysia/Jadual_14_41_Malaysia.xlsx
            target_dir = category_folder
            file_name = "Jadual_14_41_Malaysia.xlsx"
        
        # 3. Create the directory safely
        os.makedirs(target_dir, exist_ok=True)
        
        # Use realpath to resolve any Mac iCloud symlinks before giving it to AppleScript!
        save_path = os.path.realpath(os.path.join(target_dir, file_name))
        
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