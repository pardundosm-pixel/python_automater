import pandas as pd
from src.data_provider import get_metrics_dict

# ==========================================
# MAPPING CONFIGURATION FOR JADUAL 17.0 (Akaun STMK)
# ==========================================
# TODO: Map the EXCEL ROW NUMBER to the METRIC NAME in the database
# Example: 7: "jumlah_penduduk", 11: "warganegara"
ROW_MAP = {
    # =============================================
    # Akaun Satelit Teknologi Maklumat dan Komunikasi (Nilai)
    # =============================================
    11: "astmk_nilai_ditambah_kasar_industri_tmk",      # Nilai Ditambah Kasar Industri TMK (RM juta)
    14: "astmk_eksport_produk_tmk",                     # Eksport Produk TMK (RM juta)
    17: "astmk_import_produk_tmk",                      # Import Produk TMK (RM juta)
    20: "astmk_nilai_ditambah_kasar_e_dagang",          # Nilai Ditambah Kasar e-dagang (RM juta)
    23: "astmk_sumbangan_tmk_kepada_ekonomi",           # Sumbangan TMK kepada ekonomi (RM juta)
    26: "astmk_guna_tenaga_dalam_industri_tmk",         # Guna Tenaga Dalam Industri TMK ('000)

    # =============================================
    # Akaun Satelit Teknologi Maklumat dan Komunikasi - Perubahan Peratusan Tahunan (%)
    # =============================================
    32: "astmk_ perubahan_peratusan_tahunan_nilai_ditambah_kasar_industri_tmk",   # NDKTMK %
    35: "astmk_perubahan_peratusan_tahunan_eksport_produk_tmk",                  # Eksport Produk TMK %
    38: "astmk_perubahan_peratusan_tahunan_import_produk_tmk",                   # Import Produk TMK %
    41: "astmk_perubahan_peratusan_tahunan_nilai_ditambah_kasar_e_dagang",       # e-dagang %
    44: "astmk_perubahan_peratusan_tahunan_sumbangan_tmk_kepada_ekonomi",        # Sumbangan TMK %
    47: "astmk_perubahan_peratusan_tahunan_guna_tenaga_dalam_industri_tmk",      # Guna Tenaga %

    # =============================================
    # Akaun Satelit Teknologi Maklumat dan Komunikasi - Peratus Sumbangan kepada KDNK (%)
    # =============================================
    54: "peratus_kdnk_nilai_ditambah_kasar_tmk",         # NDKTMK % sumbangan kepada KDNK
    57: "peratus_kdnk_nilai_ditambah_kasar_e_dagang",    # e-dagang % sumbangan kepada KDNK
    60: "peratus_kdnk_nilai_ditambah_kasar_tmk_e_dagang" # Sumbangan TMK % kepada KDNK
}

# TODO: Map the EXCEL COLUMN NUMBER to the YEAR STRING
COL_MAP = {
    7 : "2023",  
    8 : "2024",  
    9 : "2025"   
}

def populate_jadual_17_0(sheet, hierarchy, report_type):
    print(f"  -> Populating Jadual 17.0 (Akaun STMK) untuk Malaysia_17_0")
    
    # 1. Fetch the Data Payload strictly for Malaysia
    metrics_data = get_metrics_dict("Malaysia", level='malaysia')
    
    if not metrics_data:
            print(f"     [Warning] No data found for Malaysia.")
            return

    # ==========================================
        # DYNAMIC TABLE TITLE MODIFICATION
    # ==========================================
    title_bm = ": Statistik utama Akaun Satelit Teknologi Maklumat dan Komunikasi (ASTMK), Malaysia, 2023 - 2025"
    title_en = ": Principal statistics of Information and Communication Technology Satellite Account (ICTSA), Malaysia, 2023 - 2025"
        
    # Set the exact cells where your title sits in the template
    # Targeting Column C based on standard template behavior
    sheet.range("C3").value = title_bm
    sheet.range("C4").value = title_en

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