import pandas as pd
from src.data_provider import get_metrics_dict

# ==========================================
# MAPPING CONFIGURATION FOR JADUAL 23.1 (Harga Item Terpilih)
# ==========================================
# TODO: Map the EXCEL ROW NUMBER to the METRIC NAME in the database
# Example: 7: "jumlah_penduduk", 11: "warganegara"
ROW_MAP = {
    # Sayuran (Vegetables)
    16: "bayam_hijau",                 # Bayam/Spinach
    17: "bendi",                       # Bendi/Ladies Fingers
    18: "petola",                      # Petola/Angel Gourd
    19: "bawang_besar_india",          # Bawang Besar, India
    20: "cili_padi_burung_import",     # Cili (Burung/padi) Import
    21: "cili_merah_kulai",            # Cili Merah (Kulai)
    22: "cili_merah_minyak",           # Cili Merah (Minyak)  <-- added
    23: "kacang_panjang",              # Kacang Panjang
    24: "kacang_buncis",               # Kacang Buncis
    25: "kubis_bulat",                 # Kobis Bulat
    26: "kubis_bunga",                 # Kobis Bunga
    27: "lobak_merah",                 # Lobak Merah
    28: "sawi_jepun",                  # Sawi Jepun
    29: "tomato",                      # Tomato
    30: "terung",                      # Terung
    31: "timun",                       # Timun
    
    # Buah-Buahan (Fruits)
    36: "epal_fuji",                   # Epal Fuji
    37: "epal_hijau",                  # Epal (Green Skin)
    38: "epal_merah",                  # Epal Merah
    39: "betik",                       # Betik/Papaya
    40: "nanas",                       # Nenas/Pineapple
    41: "pisang_emas",                 # Pisang Emas
    42: "tembikai_susu",               # Tembikai Susu/Honeydew
    43: "pisang_berangan",             # Pisang Berangan
    44: "tembikai_tanpa_biji",         # Tembikai Tanpa Biji
    
    # Kelapa dan Telur (Coconut & Eggs)
    49: "kelapa_parut",                # Kelapa Parut
    50: "santan",                      # Santan
    51: "telur_ayam_gred_a",           # Telur Gred A
    52: "telur_ayam_gred_b",           # Telur Gred B
    53: "telur_ayam_gred_c",           # Telur Gred C

    
    # Ikan, Ayam dan Daging (Fish, Chicken & Meat)
    58: "ikan_bawal_hitam",            # Ikan Bawal Hitam
    59: "ikan_cencaru",                # Ikan Cencaru
    60: "ikan_kembung",                # Ikan Kembang
    61: "ikan_kerisi",                 # Ikan Kerisi
    62: "ikan_merah",                  # Ikan Merah
    63: "ikan_tenggiri_batang",        # Ikan Tenggiri, Batang
    64: "ikan_tongkol_hitam",          # Ikan Tongkol, Hitam
    65: "ikan_selayang",               # Ikan Selayang
    66: "ikan_siakap",                 # Ikan Siakap
    67: "ayam",                        # Ayam
    68: "daging_lembu_tempatan",       # Daging Lembu Tempatan
    
    # Udang, Sotong dan Ketam (Prawn, Cuttlefish & Crab)
    73: "udang_8_12_sm",               # Udang (8-12 sm)
    74: "sotong_10_12_sm",             # Sotong (10-12 sm)
    75: "ketam_bunga",                 # Ketam
    
    # Makanan dan Minuman
    80: "nasi_lemak",                  # Nasi Lemak
    81: "nasi_kosong",                 # Nasi Kosong
    82: "nasi_goreng",                 # Nasi Goreng
    83: "kuey_teow_goreng",            # Kuey Teow Goreng
    84: "mee_hoon_goreng",             # Mee Hoon Goreng
    85: "nasi_biryani",                # Nasi Biryani
    86: "roti_canai",                  # Roti Canai
    87: "air_mineral",                 # Air Mineral
    88: "satay_ayam",                  # Satay Ayam
    89: "nasi_ayam",                   # Nasi Ayam
    90: "teh_tarik",                   # Teh Tarik
    91: "kopi_o",                      # Kopi-O
    92: "teh_o",                       # Teh-O
    93: "milo"                         # Milo
    
}

# TODO: Map the EXCEL COLUMN NUMBER to the YEAR STRING
COL_MAP = {
    7  : "2023",  
    9  : "2024",  
    11 : "2025"   
}

def populate_jadual_23_1(sheet, hierarchy, report_type):
    print(f"  -> Populating Jadual 23.1 (Harga Item Terpilih) untuk Malaysia_23_1")
    
    # 1. Fetch the Data Payload strictly for Malaysia
    metrics_data = get_metrics_dict("00", level='negeri')
    
    if not metrics_data:
            print(f"     [Warning] No data found for Malaysia.")
            return

    # ==========================================
        # DYNAMIC TABLE TITLE MODIFICATION
    # ==========================================
    title_bm = ": Harga purata item terpilih, Malaysia, 2023 - 2025"
    title_en = ": Average price for selected items, Malaysia, 2023 - 2025"
        
    # Set the exact cells where your title sits in the template
    # Targeting Column C based on standard template behavior
    sheet.range("C3").value = title_bm
    sheet.range("C5").value = title_en

    # Standard Injection Loop
    for col_idx, year in COL_MAP.items():
            year_data = metrics_data.get(str(year), {})
            
            for row_idx, metric_name in ROW_MAP.items():
                val = year_data.get(metric_name, "n.a")
                
                # Clean and parse missing values
                if pd.notna(val) and val != "n.a" and val != "":
                    try: 
                        val = float(val)
                    except (ValueError, TypeError): 
                        pass
                else:
                    val = "n.a"
                    
                sheet.range((row_idx, col_idx)).value = val