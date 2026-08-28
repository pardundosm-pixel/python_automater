import pandas as pd
from src.data_provider import get_metrics_dict
from src.excel_utils import safe_write

# ==========================================
# MAPPING CONFIGURATION FOR JADUAL 46.1 (samb.)(NEGERI)
# ==========================================
# Map the EXCEL ROW NUMBER to the METRIC NAME in the database
ROW_MAP = {
    # Perkakas Isi Rumah & Barangan Elektrik
    14: "periuk_nasi_elektrik",
    15: "peti_sejuk_2_pintu",
    16: "televisyen_warna_full_hd",

    # Perkakas Lain, Barang-Barang dan Produk Untuk Penjagaan Diri
    21: "lampin_bayi",
    22: "syampu",

    # Susu Segar, Susu pekat, Susu Tepung & Keluaran Susu Lain
    27: "susu_segar_uht",
    28: "susu_krimer_pekat_manis",
    29: "susu_tepung_bayi_pek",
    30: "keju",

    # Barang Pengeluaran Perubatan
    35: "ubat_batuk",
    36: "paracetamol",

    # Ikan & Makanan Laut yang diproses dan Minyak Masak
    41: "ikan_dalam_sos_tomato",
    42: "minyak_masak_3_kg",
    43: "minyak_masak_5_kg",
    44: "minyak_masak_1_kg",

    # Beras, Minuman Bermalt dan Minuman Isotonik
    49: "beras_super_special_tempatan",
    50: "minuman_bermalt",
    51: "minuman_isotonik",

    # Mee Kering, Kicap Kacang Soya Manis, Sos, Mayonis, Minuman Beralkohol dan Ubat Nyamuk
    56: "mi_kering_500gm",
    57: "mi_kering_pek",
    58: "kicap_kacang_soya_manis",
    59: "sos_tomato",
    60: "sos_cili",
    61: "mayonis",
    62: "minuman_beralkohol",
    63: "ubat_nyamuk_10_keping",
    64: "ubat_nyamuk_30_keping",
    65: "ubat_nyamuk_360_gm",

    # Barangan & Penyelenggaraan Isi Rumah, Buku & Alat tulis dan Barangan Untuk Penjagaan Diri
    70: "plastik_sampah",
    71: "pelembut_fabrik",
    72: "pen_sebatang",
    73: "pensil_warna",
    74: "kertas_fotostat",
    75: "buku_latihan",
    76: "menggunting_rambut_lelaki",
    77: "berus_gigi_satu",
    78: "ubat_gigi",
    79: "tisu_tandas",
    80: "sabun_mandi_1_pek",
    81: "sabun_mandi_250_ml"
}

# Map the EXCEL COLUMN NUMBER to the YEAR STRING
# Based on screenshot: Column J = 10, Column K = 11, Column L = 12
COL_MAP = {
    9: "2023",  
    10: "2024",  
    11: "2025"   
}

# ==========================================
# 2. REPORT INJECTION ENGINE
# ==========================================
def populate_jadual_46_1_1(sheet, hierarchy, report_type):
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