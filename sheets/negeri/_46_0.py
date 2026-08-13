import pandas as pd
from src.data_provider import get_metrics_dict

# ==========================================
# MAPPING CONFIGURATION FOR JADUAL 45.0 (NEGERI)
# ==========================================
# Map the EXCEL ROW NUMBER to the METRIC NAME in the database
ROW_MAP = {
    # Indeks Harga Pengguna
    10: "indeks_harga_pengguna_jumlah",
    13: "ihp_makanan_dan_minuman",
    16: "ihp_minuman_alkohol_dan_tembakau",
    19: "ihp_pakaian_dan_kasut",
    22: "ihp_perumahan_air_elektrik_gas_dan_bahan_api_lain",
    25: "ihp_hiasan_perkakasan_dan_penyelenggaraan_isi_rumah",
    28: "ihp_kesihatan",
    31: "ihp_pengangkutan",
    34: "ihp_maklumat_dan_komunikasi",
    37: "ihp_rekreasi_sukan_dan_kebudayaan",
    40: "ihp_pendidikan",
    43: "ihp_restoran_dan_perkhidmatan_penginapan",
    46: "ihp_penjagaan_diri_perlindungan_sosial_dan_pelbagai_barangan_dan_perkhidmatan",

    # Inflasi Tahunan
    52: "inflasi_tahunan_jumlah",
    55: "inflasi_tahunan_makanan_dan_minuman",
    58: "inflasi_tahunan_minuman_alkohol_dan_tembakau",
    61: "inflasi_tahunan_pakaian_dan_kasut",
    64: "inflasi_tahunan_perumahan_air_elektrik_gas_dan_bahan_api_lain",
    67: "inflasi_tahunan_hiasan_perkakasan_dan_penyelenggaraan_isi_rumah",
    70: "inflasi_tahunan_kesihatan",
    73: "inflasi_tahunan_pengangkutan",
    76: "inflasi_tahunan_maklumat_dan_komunikasi",
    79: "inflasi_tahunan_rekreasi_sukan_dan_kebudayaan",
    82: "inflasi_tahunan_pendidikan",
    85: "inflasi_tahunan_restoran_dan_perkhidmatan_penginapan",
    88: "inflasi_tahunan_penjagaan_diri_perlindungan_sosial_dan_pelbagai_barangan_dan_perkhidmatan"

}

# Map the EXCEL COLUMN NUMBER to the YEAR STRING
# Based on screenshot: Column J = 10, Column K = 11, Column L = 12
COL_MAP = {
    8: "2023",  
    9: "2024",  
    10: "2025"   
}

# ==========================================
# 2. REPORT INJECTION ENGINE
# ==========================================
def populate_jadual_46_0(sheet, hierarchy, report_type):
    state_name = hierarchy.get('state_name', 'Unknown State')
    state_code = hierarchy.get('state_code')
    print(f"  -> Populating Jadual 46.0 (Harga) untuk {state_name}")

    # 1. Fetch the Data Payload for the specific Negeri
    metrics_data = get_metrics_dict(state_code, level='negeri')
    
    if not metrics_data:
        print(f"     [Warning] No data found for {state_name}.")
        return

    # ==========================================
    # DYNAMIC TABLE TITLE MODIFICATION
    # ==========================================
    title_bm = f": Indeks Harga Pengguna dan Inflasi Tahunan mengikut kumpulan utama, {state_name}, 2023 - 2025"
    title_en = f": Consumer Price Index and Annual Inflation by main group, {state_name}, 2023 - 2025"

    # Set the exact cells where your title sits in the template
    sheet.range("C3").value = title_bm
    sheet.range("C4").value = title_en
    # ==========================================

    # 2. Inject Data Flush to the Grid
    for col_idx, year in COL_MAP.items():
        year_data = metrics_data.get(str(year), {})
        
        for row_idx, metric_name in ROW_MAP.items():
            val = year_data.get(metric_name, "n.a")
            
            # Sanitization and missing value fallback
            if pd.notna(val) and val != "n.a" and val != "":
                try: 
                    val = float(val)
                except (ValueError, TypeError): 
                    pass
            else:
                val = "n.a"
                
            sheet.range((row_idx, col_idx)).value = val