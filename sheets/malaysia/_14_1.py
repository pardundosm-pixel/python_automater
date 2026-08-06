import pandas as pd
from src.data_provider import get_metrics_dict

# ==========================================
# MAPPING CONFIGURATION FOR JADUAL 14.1 (NEGERI)
# ==========================================
# TODO: Map the EXCEL ROW NUMBER to the METRIC NAME in the database
# Example: 7: "jumlah_penduduk", 11: "warganegara"
ROW_MAP = {
    7:  "jumlah_penduduk",
    11: "warganegara",
    12: "bukan_warganegara",
    14: "lelaki",
    15: "perempuan",
    18: "peratus_warganegara",
    19: "peratus_bukan_warganegara",
    21: "purata_pertumbuhan_penduduk",
    26: "peratus_bumiputera",
    27: "peratus_cina",
    28: "peratus_india",
    29: "peratus_lain_lain",
    33: "umur_0_14",
    35: "umur_15_64",
    37: "umur_65_lebih",
    39: "umur_18_lebih",
    44: "jumlah_nisbah_tanggungan",
    45: "umur_muda",
    46: "umur_tua",
    48: "nisbah_jantina",
    50: "kepadatan_penduduk"
}

# TODO: Map the EXCEL COLUMN NUMBER to the YEAR STRING
COL_MAP = {
    4: "2023",  
    5: "2024",  
    6: "2025p"   
}

def populate_jadual_14_1(sheet, hierarchy):
    print(f"  -> Populating Jadual 14.1 (Penduduk Negeri) for {hierarchy['state_name']}")
    
    # Fetch Negeri data using the state_code from hierarchy
    metrics = get_metrics_dict(hierarchy['state_code'], 'negeri')

    # Dynamic Titles (Update cell reference B2 if needed)
    title_bm = f"Anggaran penduduk pertengahan tahun, {hierarchy['state_name']}, 2023 - 2025p"
    sheet.range("C2").value = title_bm

    # Standard Injection Loop
    for col_idx, year in COL_MAP.items():
        year_data = metrics.get(str(year), {})
        
        for row_idx, metric_name in ROW_MAP.items():
            val = year_data.get(metric_name, "n.a")

            if pd.notna(val) and val != "n.a" and val != "":
                try: 
                    val = float(val)
                except (ValueError, TypeError): 
                    pass
            else:
                val = "n.a"
            
            sheet.range((row_idx, col_idx)).value = val