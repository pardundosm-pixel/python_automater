import pandas as pd
from src.data_provider import get_metrics_dict
from src.excel_utils import safe_write

# ==========================================
# MAPPING CONFIGURATION FOR JADUAL 46.1 (NEGERI)
# ==========================================
# Map the EXCEL ROW NUMBER to the METRIC NAME in the database
ROW_MAP = {
    # Sayuran
    12: "bayam_hijau",
    13: "bendi",
    14: "petola",
    15: "bawang_besar_india",
    16: "cili_padi_burung_import",
    17: "cili_merah_kulai",
    18: "cili_merah_minyak",
    19: "kacang_panjang",
    20: "kacang_buncis",
    21: "kubis_bulat",
    22: "kubis_bunga",
    23: "lobak_merah",
    24: "sawi_jepun",
    25: "tomato",
    26: "terung",
    27: "timun",

    # Buah-Buahan
    32: "epal_fuji",
    33: "epal_hijau",
    34: "epal_merah",
    35: "betik",
    36: "nanas",
    37: "pisang_emas",
    38: "tembikai_susu",
    39: "pisang_berangan",
    40: "tembikai_tanpa_biji",

    # Kelapa dan Telur
    45: "kelapa_parut",
    46: "santan",
    47: "telur_ayam_gred_a",
    48: "telur_ayam_gred_b",
    49: "telur_ayam_gred_c",

    #Ikan, Ayam dan Daging
    54: "ikan_bawal_hitam",
    55: "ikan_cencaru",
    56: "ikan_kembung",
    57: "ikan_kerisi",
    58: "ikan_merah",
    59: "ikan_tenggiri_batang",
    60: "ikan_tongkol_hitam",
    61: "ikan_selayang",
    62: "ikan_siakap",
    63: "ayam",
    64: "daging_lembu_tempatan",

    # Udang, Sotong dan Ketam
    69: "udang_8_12_sm",
    70: "sotong_10_12_sm",
    71: "ketam_bunga",

    # Makanan dan Minuman
    76: "nasi_lemak",
    77: "nasi_kosong",
    78: "nasi_goreng",
    79: "kuey_teow_goreng",
    80: "mee_hoon_goreng",
    81: "nasi_biryani",
    82: "roti_canai",
    83: "air_mineral",
    84: "satay_ayam",
    85: "nasi_ayam",
    86: "teh_tarik",
    87: "kopi_o",
    88: "teh_o",
    89: "milo"
}

# Map the EXCEL COLUMN NUMBER to the YEAR STRING
# Based on screenshot: Column J = 10, Column K = 11, Column L = 12
COL_MAP = {
    7: "2023",  
    8: "2024",  
    9: "2025"   
}

# ==========================================
# 2. REPORT INJECTION ENGINE
# ==========================================
def populate_jadual_46_1(sheet, hierarchy, report_type):
    state_name = hierarchy.get('state_name', 'Unknown State')
    state_code = hierarchy.get('state_code')
    print(f"  -> Populating Jadual 42.0 (KDNK) untuk {state_name}")

    metrics_data = get_metrics_dict(state_code, level='negeri')
    if not metrics_data:
        print(f"     [Warning] No data found for {state_name}.")
        return

    # 1. Inject Titles Safely (Row, Column)
    title_bm = f": Keluaran Dalam Negeri Kasar (KDNK), {state_name}, 2023 - 2025p"
    title_en = f": Gross Domestic Product (GDP), {state_name}, 2023 - 2025p"
    safe_write(sheet, 3, 3, title_bm) # C3
    safe_write(sheet, 4, 3, title_en) # C4