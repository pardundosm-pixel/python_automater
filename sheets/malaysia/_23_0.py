import pandas as pd
from src.data_provider import get_metrics_dict

# ==========================================
# MAPPING CONFIGURATION FOR JADUAL 23.0 (IHP)
# ==========================================
# TODO: Map the EXCEL ROW NUMBER to the METRIC NAME in the database
# Example: 7: "jumlah_penduduk", 11: "warganegara"
ROW_MAP = {
    # Indeks Harga pengguna
    9 :  "indeks_harga_pengguna_jumlah",
    12:  "ihp_makanan_dan_minuman",
    15:  "ihp_minuman_alkohol_dan_tembakau",
    18:  "ihp_pakaian_dan_kasut",
    21:  "ihp_perumahan_air_elektrik_gas_dan_bahan_api_lain",
    24:  "ihp_hiasan_perkakasan_dan_penyelenggaraan_isi_rumah",
    27:  "ihp_kesihatan",
    30:  "ihp_pengangkutan",
    33:  "ihp_maklumat_dan_komunikasi",
    36:  "ihp_rekreasi_sukan_dan_kebudayaan",
    39:  "ihp_pendidikan",
    42:  "ihp_restoran_dan_perkhidmatan_penginapan",
    45:  "ihp_insurans_dan_perkhidmatan_kewangan",
    48:  "ihp_penjagaan_diri_perlindungan_sosial_dan_pelbagai_barangan_dan_perkhidmatan",
    
    # Inflasi Tahunan (Perubahan %)
    
    54:  "inflasi_tahunan_jumlah",
    57:  "inflasi_tahunan_makanan_dan_minuman",
    60:  "inflasi_tahunan_minuman_alkohol_dan_tembakau",
    63:  "inflasi_tahunan_pakaian_dan_kasut",
    66:  "inflasi_tahunan_perumahan_air_elektrik_gas_dan_bahan_api_lain",
    69:  "inflasi_tahunan_hiasan_perkakasan_dan_penyelenggaraan_isi_rumah",
    72:  "inflasi_tahunan_kesihatan",
    75:  "inflasi_tahunan_pengangkutan",
    78:  "inflasi_tahunan_maklumat_dan_komunikasi",
    81:  "inflasi_tahunan_rekreasi_sukan_dan_kebudayaan",
    84:  "inflasi_tahunan_pendidikan",
    87:  "inflasi_tahunan_restoran_dan_perkhidmatan_penginapan",
    90:  "inflasi_tahunan_insurans_dan_perkhidmatan_kewangan",
    93:  "inflasi_tahunan_penjagaan_diri_perlindungan_sosial_dan_pelbagai_barangan_dan_perkhidmatan"
    
}

# TODO: Map the EXCEL COLUMN NUMBER to the YEAR STRING
COL_MAP = {
    7 : "2023",  
    8 : "2024",  
    9 : "2025"   
}

def populate_jadual_23_0(sheet, hierarchy, report_type):
    print(f"  -> Populating Jadual 23.0 (IHP) untuk Malaysia_23_0")
    
    # 1. Fetch the Data Payload strictly for Malaysia
    metrics_data = get_metrics_dict("00", level='negeri')
    
    if not metrics_data:
            print(f"     [Warning] No data found for Malaysia.")
            return

    # ==========================================
        # DYNAMIC TABLE TITLE MODIFICATION
    # ==========================================
    title_bm = ": Indeks Harga Pengguna dan Inflasi Tahunan mengikut Kumpulan Utama, Malaysia, 2023 - 2025"
    title_en = ": Consumer Price Index and Annual Inflation by Main Group, Malaysia, 2023 - 2025"
        
    # Set the exact cells where your title sits in the template
    # Targeting Column C based on standard template behavior
    sheet.range("D1").value = title_bm
    sheet.range("D2").value = title_en

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