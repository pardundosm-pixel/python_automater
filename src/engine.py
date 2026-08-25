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
# ===== Jadual 14.0 to 22.01 =====
from sheets.malaysia._14_0 import populate_jadual_14_0 as jadual_14_0_malaysia
from sheets.malaysia._14_1 import populate_jadual_14_1 as jadual_14_1_malaysia
from sheets.malaysia._14_2 import populate_jadual_14_2 as jadual_14_2_malaysia
from sheets.malaysia._15_0 import populate_jadual_15_0 as jadual_15_0_malaysia
from sheets.malaysia._16_0 import populate_jadual_16_0 as jadual_16_0_malaysia
from sheets.malaysia._17_0 import populate_jadual_17_0 as jadual_17_0_malaysia
from sheets.malaysia._18_0 import populate_jadual_18_0 as jadual_18_0_malaysia
from sheets.malaysia._19_0 import populate_jadual_19_0 as jadual_19_0_malaysia
from sheets.malaysia._20_0 import populate_jadual_20 as jadual_20_0_malaysia
from sheets.malaysia._21_0 import populate_jadual_21_0 as jadual_21_0_malaysia
from sheets.malaysia._22_0 import populate_jadual_22_0 as jadual_22_0_malaysia
from sheets.malaysia._22_01 import populate_jadual_22_01 as jadual_22_01_malaysia
# ===== Jadual 23.0 to 25.0 =====
from sheets.malaysia._23_0 import populate_jadual_23_0 as jadual_23_0_malaysia
from sheets.malaysia._23_1 import populate_jadual_23_1 as jadual_23_1_malaysia
from sheets.malaysia._23_1_1 import populate_jadual_23_1_1 as jadual_23_1_1_malaysia
from sheets.malaysia._24_0 import populate_jadual_24_0 as jadual_24_0_malaysia
from sheets.malaysia._25_0 import populate_jadual_25_0 as jadual_25_0_malaysia
from sheets.malaysia._25_0_1 import populate_jadual_25_0_1 as jadual_25_0_1_malaysia
# ===== Jadual 26.0 to 30.0 =====
from sheets.malaysia._26_0 import populate_jadual_26_0 as jadual_26_0_malaysia
from sheets.malaysia._27_0 import populate_jadual_27_0 as jadual_27_0_malaysia
from sheets.malaysia._28_0 import populate_jadual_28_0 as jadual_28_0_malaysia
from sheets.malaysia._29_0 import populate_jadual_29_0 as jadual_29_0_malaysia
from sheets.malaysia._30_0 import populate_jadual_30_0 as jadual_30_0_malaysia
# ===== Jadual 31.0 to 35.0 =====
from sheets.malaysia._31_0 import populate_jadual_31 as jadual_31_0_malaysia
from sheets.malaysia._31_1 import populate_jadual_31_1 as jadual_31_1_malaysia
from sheets.malaysia._32_0 import populate_jadual_32 as jadual_32_0_malaysia
from sheets.malaysia._33_0 import populate_jadual_33 as jadual_33_0_malaysia
from sheets.malaysia._34_0 import populate_jadual_34 as jadual_34_0_malaysia
from sheets.malaysia._35_0 import populate_jadual_35 as jadual_35_0_malaysia
# ===== Jadual 36.0 to 41.0 =====
from sheets.malaysia._36_0 import populate_jadual_36_0 as jadual_36_0_malaysia
from sheets.malaysia._36_1 import populate_jadual_36_1 as jadual_36_1_malaysia
from sheets.malaysia._37_0 import populate_jadual_37 as jadual_37_0_malaysia
from sheets.malaysia._37_0_1 import populate_jadual_37_0_1 as jadual_37_0_1_malaysia
from sheets.malaysia._37_0_2 import populate_jadual_37_0_2 as jadual_37_0_2_malaysia
from sheets.malaysia._37_0_3 import populate_jadual_37_0_3 as jadual_37_0_3_malaysia
from sheets.malaysia._38_0 import populate_jadual_38 as jadual_38_0_malaysia
from sheets.malaysia._38_0_1 import populate_jadual_38_0_1 as jadual_38_0_1_malaysia
from sheets.malaysia._38_0_2 import populate_jadual_38_0_2 as jadual_38_0_2_malaysia
from sheets.malaysia._38_0_3 import populate_jadual_38_0_3 as jadual_38_0_3_malaysia
from sheets.malaysia._39_0 import populate_jadual_39 as jadual_39_0_malaysia
from sheets.malaysia._39_0_1 import populate_jadual_39_0_1 as jadual_39_0_1_malaysia
from sheets.malaysia._40_0 import populate_jadual_40 as jadual_40_0_malaysia
from sheets.malaysia._40_1 import populate_jadual_40_1 as jadual_40_1_malaysia
from sheets.malaysia._41_0 import populate_jadual_41 as jadual_41_0_malaysia

# (Negeri Template Mappers)
from sheets.negeri._42_1 import populate_jadual_42_1 as jadual_42_1_negeri
from sheets.negeri._43_0 import populate_jadual_43 as jadual_43_negeri
from sheets.negeri._42_1 import populate_jadual_42_1 as jadual_42_1_negeri
from sheets.negeri._42_0 import populate_jadual_42 as jadual_42_0_negeri
from sheets.negeri._43_0 import populate_jadual_43 as jadual_43_negeri
from sheets.negeri._44_0 import populate_jadual_44_0 as jadual_44_negeri
from sheets.negeri._45_0 import populate_jadual_45_0 as jadual_45_0_negeri
from sheets.negeri._45_0_1 import populate_jadual_45_0_1 as jadual_45_0_1_negeri
from sheets.negeri._46_0 import populate_jadual_46_0 as jadual_46_0_negeri
from sheets.negeri._46_1 import populate_jadual_46_1 as jadual_46_1_negeri
from sheets.negeri._46_1_1 import populate_jadual_46_1_1 as jadual_46_1_1_negeri
from sheets.negeri._47_0 import populate_jadual_47_0 as jadual_47_0_negeri
from sheets.negeri._47_0_1 import populate_jadual_47_0_1 as jadual_47_0_1_negeri
from sheets.negeri._48_0 import populate_jadual_48_0 as jadual_48_0_negeri
from sheets.negeri._49_0 import populate_jadual_49_0 as jadual_49_0_negeri
from sheets.negeri._49_0_1 import populate_jadual_49_0_1 as jadual_49_0_1_negeri
from sheets.negeri._50_0 import populate_jadual_50_0 as jadual_50_0_negeri
from sheets.negeri._51_0 import populate_jadual_51 as jadual_51_0_negeri
from sheets.negeri._52_0 import populate_jadual_52 as jadual_52_0_negeri
from sheets.negeri._54_0 import populate_jadual_54_0 as jadual_54_0_negeri
from sheets.negeri._55_0 import populate_jadual_55_0 as jadual_55_0_negeri

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
    
    # Profile 2: Malaysia
    "malaysia": {
        # ===== 14.0 to 19.0 =====
        "14.0_KDNK": jadual_14_0_malaysia,
        "14.0_KDNK_1": jadual_14_1_malaysia,
        "14.0_KDNK_2": jadual_14_2_malaysia,
        "15.0_Pasaran_Kewagn": jadual_15_0_malaysia,
        "16.0_Akaun_ SP": jadual_16_0_malaysia,
        "17.0_Akaun_ STMK": jadual_17_0_malaysia,
        "18.0_Stok_Modal": jadual_18_0_malaysia,
        "19.0_Imbagn_Pembayaran": jadual_19_0_malaysia,

        # ===== 20.0 to 22.0 =====
        "20.0_Eksport_Import": jadual_20_0_malaysia,
        "21.0_Pelancongn_Domestik": jadual_21_0_malaysia,
        "22.0_Pasaran_Buruh": jadual_22_0_malaysia,
        "22.0_Pasaran_Buruh_1": jadual_22_01_malaysia,

        # ===== 23.0 to 25.0 =====
        "23.0_IHP": jadual_23_0_malaysia,
        "23.1_Harga_Item_Terpilih": jadual_23_1_malaysia,
        "23.1_Harga_Item_Terpilih_1": jadual_23_1_1_malaysia,
        "24.0_IHPR": jadual_24_0_malaysia,
        "25.0_Kecederaan_Pekerja": jadual_25_0_malaysia,
        "25.0_Kecederaan_Pekerja_1": jadual_25_0_1_malaysia,

        # ===== 26.0 to 30.0 =====
        "26.0_Peratusan_Ahli_Parlimen": jadual_26_0_malaysia,
        "27.0_Bilangan_Hakim": jadual_27_0_malaysia,
        "28.0_Bilangan_Hakim_Syariah": jadual_28_0_malaysia,
        "29.0_Bil_KSU": jadual_29_0_malaysia,
        "30.0_Bil_Murid": jadual_30_0_malaysia,

        # ===== 31.0 to 35.0 =====
        "31.0_Jenayah_kekerasan": jadual_31_0_malaysia,
        "31.1_Jenayah_harta_benda": jadual_31_1_malaysia,
        "32.0_Kemalangan_Jalan_Raya": jadual_32_0_malaysia,
        "33.0 Oku_kumulatif": jadual_33_0_malaysia,          # note the space
        "34.0_Capaian_ICT_Isi_Rumh": jadual_34_0_malaysia,
        "35.0_Pendpatn_Perbelanjaan": jadual_35_0_malaysia,

        # ===== 36.0 to 37.0 =====
        "36.0_Penggunaan_Perkapita_PCC": jadual_36_0_malaysia,
        "36.1_Perkapita_Pertanian": jadual_36_1_malaysia,
        "37.0_Purata_Suhu": jadual_37_0_malaysia,
        "37.0_Purata_Suhu_1": jadual_37_0_1_malaysia,
        "37.0_Purata_Suhu_2": jadual_37_0_2_malaysia,
        "37.0_Purata_Suhu_3": jadual_37_0_3_malaysia,

        # ===== 38.0 to 39.0 =====
        "38.0_hujan": jadual_38_0_malaysia,
        "38.0_hujan_1": jadual_38_0_1_malaysia,
        "38.0_hujan_2": jadual_38_0_2_malaysia,
        "38.0_hujan_3": jadual_38_0_3_malaysia,
        "39.0_Kelembapan": jadual_39_0_malaysia,
        "39_0_Kelembapan_1": jadual_39_0_1_malaysia,

        # ===== 40.0 to 41.0 =====
        "40.0_Kes_COVID 19": jadual_40_0_malaysia,           # note the space
        "40.1_Penerima_Vaksin_Malaysia": jadual_40_1_malaysia,
        "41.0_Agihan_Zakat": jadual_41_0_malaysia,
    },
    
    # Profile 3: Negeri
    "negeri": {
        "42.0_KDNK":jadual_42_0_negeri,
        "42.1_KDNK_DP": jadual_42_1_negeri,
        "43.0_Dagangan": jadual_43_negeri,
        "44.0_Pelancongan": jadual_44_negeri,
        "45.0_Buruh": jadual_45_0_negeri,
        "45.0_Buruh_1": jadual_45_0_1_negeri,
        "46.0_Harga": jadual_46_0_negeri,
        "46.1_AUP": jadual_46_1_negeri,
        "46.1_AUP_1": jadual_46_1_1_negeri,
        "47.0_Kemalangan_Pekerjaan": jadual_47_0_negeri,
        "47.0_Kemalangan_Pekerjaan_1": jadual_47_0_1_negeri,
        "48.0_Bil_Murid_negeri": jadual_48_0_negeri, 
        "49.0_Taman_Asuhan": jadual_49_0_negeri,
        "49.0_Asuhan_bil_kanak": jadual_49_0_1_negeri,
        "50.0_Jenayah": jadual_50_0_negeri,
        "51.0_Kemalangan_Jalan_Raya": jadual_51_0_negeri,
        "52.0_ICT": jadual_52_0_negeri,
        "54.0_Pertanian": jadual_54_0_negeri,
        "55.0_Penerima_Vaksin_state": jadual_55_0_negeri,
    }
}


# ==========================================
# 3. REPORT ORCHESTRATOR
# ==========================================
def generate_report(location_code: str, report_type: str, excel_app: xw.App, parent_code: str = None, template_key: str = "parlimen_dun"):
    print(f"\n=============================================")
    print(f"Generating {report_type.upper()} Report [{template_key}] for: {location_code}")
    
    # --- A. HIERARCHY FETCHING ---
    if report_type == 'parlimen':
        hierarchy = get_location_hierarchy(location_code)
    elif report_type == 'dun':
        hierarchy = get_dun_hierarchy(location_code, parent_code)
    elif report_type == 'negeri':
        hierarchy = get_negeri_hierarchy(location_code)  # <--- NEW: FETCH NEGERI DATA
    elif report_type == 'malaysia':
        hierarchy = {'state_name': 'Malaysia', 'location_name': 'Malaysia', 'state_code': '00'}
    else:
        print(f"  -> [Error] Invalid report_type '{report_type}'. Must be 'parlimen', 'dun', 'negeri', or 'malaysia'.")
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