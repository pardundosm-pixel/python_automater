import pandas as pd
from src.data_provider import get_metrics_dict

# ==========================================
# MAPPING CONFIGURATION FOR JADUAL 23.1 (Harga Item Terpilih 1)
# ==========================================
# TODO: Map the EXCEL ROW NUMBER to the METRIC NAME in the database
# Example: 7: "jumlah_penduduk", 11: "warganegara"
ROW_MAP = {
    # Perkakas Isi Rumah & Barangan Elektrik
    16: "periuk_nasi_elektrik",        # Periuk Nasi Elektrik, 1.8LT
    17: "peti_sejuk_2_pintu",          # Peti Sejuk 2-Pintu, kapasiti 280-300L
    18: "televisyen_warna_full_hd",    # Televisyen Warna, Full HD 40"-43"

    # Perkakas lain, Barang-Barang dan Produk Untuk Penjagaan Diri
    23: "lampin_bayi",                 # Lampin Bayi, Saiz M
    24: "syampu",                      # Syampu, 300 ML

    # Susu Segar, Susu Pekat, Susu Tepung & Keluaran Susu Lain
    29: "susu_segar_uht",              # Susu Segar UHT
    30: "susu_krimer_pekat_manis",     # Susu, Krimer Pekat Manis
    31: "susu_tepung_bayi_pek",        # Susu Tepung Bayi
    32: "keju",                        # Keju

    # Barang Pengeluaran Perubatan
    37: "ubat_batuk",                  # Ubat batuk
    38: "paracetamol",                 # Paracetamol

    # Ikan & Makanan Laut yang diproses dan Minyak Masak
    43: "ikan_dalam_sos_tomato",       # Ikan dalam Sos Tomato
    44: "minyak_masak_3_kg",           # Minyak Masak (3 Kg)
    45: "minyak_masak_5_kg",           # Minyak Masak (5 Kg)
    46: "minyak_masak_1_kg",           # Minyak Masak (1 Kg)

    # Beras, Minuman Bermalt dan Minuman Isotonik
    51: "beras_super_special_tempatan",# Beras SST 5%
    52: "minuman_bermalt",             # Minuman bermalt (pek)
    53: "minuman_isotonik",            # Minuman Isotonik

    # Mee Kering, Kicap Kacang Soya Manis, Sos, Mayonis, Minuman Beralkohol dan Ubat Nyamuk
    58: "mi_kering_500gm",             # Mee Kering (500 gm)
    59: "mi_kering_pek",               # Mee Kering (5 pek)
    60: "kicap_kacang_soya_manis",     # Kicap Kacang Soya, manis
    61: "sos_tomato",                  # Sos Tomato
    62: "sos_cili",                    # Sos Cili
    63: "mayonis",                     # Mayonis
    64: "minuman_beralkohol",          # Minuman Beralkohol
    65: "ubat_nyamuk_10_keping",       # Ubat Nyamuk (10 keping)
    66: "ubat_nyamuk_30_keping",       # Ubat Nyamuk (30 keping)
    67: "ubat_nyamuk_360_gm",          # Ubat Nyamuk (360 gm)

    # Barangan & Penyelenggaraan Isi Rumah, Buku & Alat tulis dan Barangan Untuk Penjagaan Diri
    72: "plastik_sampah",              # Plastik Sampah
    73: "pelembut_fabrik",             # Pelembut Fabrik
    74: "pen_sebatang",                # Pen
    75: "pensil_warna",                # Pensel Warna
    76: "kertas_fotostat",             # Kertas Fotostat
    77: "buku_latihan",                # Buku Latihan
    78: "menggunting_rambut_lelaki",   # Potong Rambut (Lelaki)
    79: "berus_gigi_satu",             # Berus Gigi
    80: "ubat_gigi",                   # Ubat Gigi
    81: "tisu_tandas",                 # Tisu Tandas
    82: "sabun_mandi_1_pek",           # Sabun Mandian Buku (1 pek 3, 70 gm)
    83: "sabun_mandi_250_ml"           # Sabun Mandi (250 ml)
}

# TODO: Map the EXCEL COLUMN NUMBER to the YEAR STRING
COL_MAP = {
    7  : "2023",  
    9  : "2024",  
    11 : "2025"   
}

def populate_jadual_23_1_1(sheet, hierarchy, report_type):
    print(f"  -> Populating Jadual 23.1 (Harga Item Terpilih 1) untuk Malaysia_23_1_1")
    
    # 1. Fetch the Data Payload strictly for Malaysia
    metrics_data = get_metrics_dict("Malaysia", level='negeri')
    
    if not metrics_data:
            print(f"     [Warning] No data found for Malaysia.")
            return

    # ==========================================
        # DYNAMIC TABLE TITLE MODIFICATION
    # ==========================================
    title_bm = ": Harga purata item terpilih, Malaysia, 2023 - 2025 (samb.)"
    title_en = ": Average price for selected items, Malaysia, 2023 - 2025 (cont'd)"
        
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