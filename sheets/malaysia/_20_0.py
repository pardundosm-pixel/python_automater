import pandas as pd
from src.data_provider import get_metrics_dict

# ==========================================
# 1. MAPPING CONFIGURATION FOR JADUAL 20.0 (MALAYSIA)
# ==========================================

# --- SET A: NILAI (VALUE) MAPPINGS ---
# Map the Rows for the 'Value' metrics
ROW_MAP_VALUE = {
    # Eksport
    9:  "eksport_nilai_eksport",
    11: "eksport_nilai_pertanian",
    12: "eksport_nilai_perlombongan",
    13: "eksport_nilai_pembuatan",
    14: "eksport_nilai_lain_lain",
    18: "eksport_nilai_produk_elektrik_dan_elektronik",
    19: "eksport_nilai_keluaran_petroleum_bertapis",
    20: "eksport_nilai_minyak_sawit_dan_produk_berasaskan_minyak_sawit",
    
    # Import
    23: "import_nilai_import",
    24: "import_nilai_tertangguh",
    26: "import_nilai_modal",
    27: "import_nilai_perantaraaan",
    28: "import_nilai_penggunaan",
    29: "import_nilai_lain_lain",
    33: "import_nilai_elektrik_dan_elektronik",
    34: "import_nilai_keluaran_petroleum_bertapis",
    35: "import_nilai_petroleum_mentah",
    
    # Jumlah Dagangan
    37: "nilai_jumlah_dagangan",
    
    # Imbangan Dagangan
    39: "nilai_imbangan_dagangan",
    
    # 5 Rakan Dagangan Utama (Jumlah Dagangan)
    41: "nilai_5_rakan",
    42: "nilai_china",
    43: "nilai_singapura",
    44: "nilai_amerika",
    45: "nilai_kesatuan_eropah",
    46: "nilai_taiwan"

}

# Map Excel Columns E, F, G to Years
COL_MAP_VALUE = {
    6: "2023",  
    7: "2024",  
    8: "2025"   
}

# --- SET B: PERTUMBUHAN (GROWTH) MAPPINGS ---
# Map the Rows for the 'Percentage' metrics (Same rows, different database metric names)
ROW_MAP_GROWTH = {
    # Eksport
    9:  "eksport_pertumbuhan_tahunan_eksport",
    11: "eksport_pertumbuhan_tahunan_pertanian",
    12: "eksport_pertumbuhan_tahunan_perlombongan",
    13: "eksport_pertumbuhan_tahunan_pembuatan",
    14: "eksport_pertumbuhan_tahunan_lain_lain",
    18: "eksport_pertumbuhan_tahunan_produk_elektrik_dan_elektronik",
    19: "eksport_pertumbuhan_tahunan_keluaran_petroleum_bertapis",
    20: "eksport_pertumbuhan_tahunan_minyak_sawit_dan_produk_berasaskan_minyak_sawit",
    
    # Import
    23: "import_pertumbuhan_tahunan_import",
    24: "import_pertumbuhan_tahunan_tertangguh",
    26: "import_pertumbuhan_tahunan_modal",
    27: "import_pertumbuhan_tahunan_perantaraaan",
    28: "import_pertumbuhan_tahunan_penggunaan",
    29: "import_pertumbuhan_tahunan_lain_lain",
    33: "import_pertumbuhan_tahunan_elektrik_dan_elektronik",
    34: "import_pertumbuhan_tahunan_keluaran_petroleum_bertapis",
    35: "import_pertumbuhan_tahunan_petroleum_mentah",
    
    # Jumlah Dagangan
    37: "import_pertumbuhan_tahunan_jumlah_dagangan",
    
    # Imbangan Dagangan
    39: "import_pertumbuhan_tahunan_imbangan_dagangan",
    
    # 5 Rakan Dagangan Utama (Jumlah Dagangan)
    41: "pertumbuhan_tahunan_5_rakan",
    42: "pertumbuhan_tahunan_china",
    43: "pertumbuhan_tahunan_singapura",
    44: "pertumbuhan_tahunan_amerika",
    45: "pertumbuhan_tahunan_kesatuan_eropah",
    46: "pertumbuhan_tahunan_taiwan",
    
    # Sumbangan kepada Jumlah Dagangan Malaysia (%)
    48: "nilai_sumbangan_jumlah_dagangan"
}

# Map Excel Columns I, J, K to Years
COL_MAP_GROWTH = {
    10: "2023",  
    11: "2024",  
    12: "2025"   
}


# ==========================================
# 2. REPORT INJECTION ENGINE
# ==========================================
def populate_jadual_20(sheet, hierarchy, report_type):
    print(f"  -> Populating Jadual 20.0 (Eksport Import) untuk Malaysia")

    # 1. Fetch the Data Payload strictly for Malaysia
    metrics_data = get_metrics_dict("Malaysia", level='malaysia')
    
    if not metrics_data:
        print(f"     [Warning] No data found for Malaysia.")
        return

    # 2. Static Title Injection
    title_bm = ": Eksport, import, jumlah dagangan dan imbangan dagangan, Malaysia, 2023 - 2025"
    title_en = ": Exports, imports, total trade and balance of trade, Malaysia, 2023 - 2025"
    sheet.range("C3").value = title_bm
    sheet.range("C4").value = title_en

    # 3. Inject Data using a helper function to avoid repeating logic
    def inject_grid(col_map, row_map):
        for col_idx, year in col_map.items():
            year_data = metrics_data.get(str(year), {})
            for row_idx, metric_name in row_map.items():
                val = year_data.get(metric_name, "n.a")
                
                if pd.notna(val) and val != "n.a" and val != "":
                    try: 
                        val = float(val)
                    except (ValueError, TypeError): 
                        pass
                else:
                    val = "n.a"
                    
                sheet.range((row_idx, col_idx)).value = val

    # Execute injection for both sides of the table
    inject_grid(COL_MAP_VALUE, ROW_MAP_VALUE)
    inject_grid(COL_MAP_GROWTH, ROW_MAP_GROWTH)