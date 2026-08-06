import pandas as pd
from src.data_provider import get_metrics_dict

# ==========================================
# 1. MAPPING CONFIGURATION FOR JADUAL 20.0 (MALAYSIA)
# ==========================================

# --- SET A: NILAI (VALUE) MAPPINGS ---
# Map the Rows for the 'Value' metrics
ROW_MAP_VALUE = {
    9:  "eksport_nilai",
    11: "eksport_pertanian_nilai",
    12: "eksport_perlombongan_nilai",
    13: "eksport_pembuatan_nilai",
    14: "eksport_lain_lain_nilai",
    18: "eksport_elektrik_elektronik_nilai",
    19: "eksport_petroleum_bertapis_nilai",
    20: "eksport_minyak_sawit_nilai",
    23: "import_nilai",
    24: "import_tertangguh_nilai",
    26: "import_barangan_modal_nilai",
    27: "import_barangan_perantaraan_nilai",
    28: "import_barangan_penggunaan_nilai",
    29: "import_lain_lain_nilai",
    33: "import_elektrik_elektronik_nilai",
    34: "import_petroleum_bertapis_nilai",
    35: "import_petroleum_mentah_nilai"
}

# Map Excel Columns E, F, G to Years
COL_MAP_VALUE = {
    5: "2022",  
    6: "2023",  
    7: "2024"   
}

# --- SET B: PERTUMBUHAN (GROWTH) MAPPINGS ---
# Map the Rows for the 'Percentage' metrics (Same rows, different database metric names)
ROW_MAP_GROWTH = {
    9:  "eksport_peratus",
    11: "eksport_pertanian_peratus",
    12: "eksport_perlombongan_peratus",
    13: "eksport_pembuatan_peratus",
    14: "eksport_lain_lain_peratus",
    18: "eksport_elektrik_elektronik_peratus",
    19: "eksport_petroleum_bertapis_peratus",
    20: "eksport_minyak_sawit_peratus",
    23: "import_peratus",
    24: "import_tertangguh_peratus",
    26: "import_barangan_modal_peratus",
    27: "import_barangan_perantaraan_peratus",
    28: "import_barangan_penggunaan_peratus",
    29: "import_lain_lain_peratus",
    33: "import_elektrik_elektronik_peratus",
    34: "import_petroleum_bertapis_peratus",
    35: "import_petroleum_mentah_peratus"
}

# Map Excel Columns I, J, K to Years
COL_MAP_GROWTH = {
    9:  "2022",  
    10: "2023",  
    11: "2024"   
}


# ==========================================
# 2. REPORT INJECTION ENGINE
# ==========================================
def populate_jadual_20(sheet, hierarchy, report_type):
    print(f"  -> Populating Jadual 20.0 (Perdagangan) untuk Malaysia")

    # 1. Fetch the Data Payload strictly for Malaysia
    metrics_data = get_metrics_dict("Malaysia", level='malaysia')
    
    if not metrics_data:
        print(f"     [Warning] No data found for Malaysia.")
        return

    # 2. Static Title Injection
    title_bm = ": Eksport, import, jumlah dagangan dan imbangan dagangan, Malaysia, 2022 - 2024"
    title_en = ": Exports, imports, total trade and balance of trade, Malaysia, 2022 - 2024"
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