import pandas as pd
from src.data_provider import get_metrics_dict

# ==========================================
# MAPPING CONFIGURATION FOR JADUAL 2.1 (NEGERI)
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

# ==========================================
# 2. REPORT INJECTION ENGINE
# ==========================================
# TODO: Rename the function to match the specific jadual (e.g., populate_jadual_2_2)
def populate_jadual_2_1(sheet, hierarchy, report_type):
    # Even though the report is running for P.143, Jadual 2.1 is ALWAYS Malaysia data
    print(f"  -> Populating Jadual 2.1 (Negeri) for {hierarchy['state_name']}")

    # 1. Fetch the Data Payload strictly for Malaysia
    metrics_data = get_metrics_dict("Malaysia", level='malaysia')
    
    if not metrics_data:
        print(f"     [Warning] No data found for Malaysia.")
        return

    # ==========================================
    # DYNAMIC TABLE TITLE MODIFICATION
    # ==========================================
    title_bm = f": Anggaran penduduk pertengahan tahun, {hierarchy['state_name']}, 2022 - 2025p"
    title_en = f": Mid-year population estimates, {hierarchy['state_name']}, 2022 - 2025p"

    # Set the exact cells where your title sits in the template
    sheet.range("C2").value = title_bm
    sheet.range("C3").value = title_en
    # ==========================================

    # 2. Inject Data
    # Loop over the columns (Years) first
    for col_idx, year in COL_MAP.items():
        year_data = metrics_data.get(str(year), {})
        
        # --- THE FIX IS HERE ---
        # We only unpack row_idx and metric_name (no year tuple!)
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