import pandas as pd
from src.data_provider import get_metrics_dict

# ==========================================
# MAPPING CONFIGURATION FOR JADUAL 2.1 (NEGERI)
# ==========================================
# TODO: Map the EXCEL ROW NUMBER to the METRIC NAME in the database
# Example: 7: "jumlah_penduduk", 11: "warganegara"
ROW_MAP = {
    7:  "jumlah_penduduk",
    11: "penduduk_warganegara",
    12: "penduduk_bukan_warganegara",
    14: "penduduk_lelaki",
    15: "penduduk_perempuan",
    18: "peratus_penduduk_warganegara",
    19: "peratus_penduduk_bukan_warganegara",
    21: "purata_pertumbuhan_penduduk",
    26: "peratus_penduduk_bumiputera",
    27: "peratus_penduduk_cina",
    28: "peratus_penduduk_india",
    29: "peratus_penduduk_lain_lain",
    33: "penduduk_umur_0_14",
    35: "penduduk_umur_15_30", #variable baru
    37: "penduduk_umur_15_64",
    39: "penduduk_umur_65_lebih",
    41: "penduduk_umur_18_lebih",
    46: "jumlah_nisbah_tanggungan",
    47: "nisbah_tanggungan_umur_muda",
    48: "nisbah_tanggungan_umur_tua",
    50: "nisbah_jantina",
    52: "kepadatan_penduduk"
}

# TODO: Map the EXCEL COLUMN NUMBER to the YEAR STRING
COL_MAP = {
    4: "2024",  
    5: "2025",  
    6: "2026p"   
}

# ==========================================
# 2. REPORT INJECTION ENGINE
# ==========================================
# TODO: Rename the function to match the specific jadual (e.g., populate_jadual_2_2)
def populate_jadual_2_1(sheet, hierarchy, report_type):
    # Even though the report is running for P.143, Jadual 2.1 is ALWAYS Malaysia data
    print(f"  -> Populating Jadual 2.1 (Negeri) for {hierarchy['state_name']}")

    # 1. Fetch the Data Payload strictly for Malaysia
    metrics_data = get_metrics_dict("Malaysia", level='negeri')
    
    if not metrics_data:
        print(f"     [Warning] No data found for Negeri.")
        return

    # ==========================================
    # DYNAMIC TABLE TITLE MODIFICATION
    # ==========================================
    title_bm = f": Anggaran penduduk pertengahan tahun, {hierarchy['state_name']}, 2024 - 2026p"
    title_en = f": Mid-year population estimates, {hierarchy['state_name']}, 2024 - 2026p"

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