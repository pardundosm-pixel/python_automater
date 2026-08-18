import pandas as pd
from src.data_provider import get_metrics_dict

# ==========================================
# MAPPING CONFIGURATION FOR JADUAL 19.0 (Imbangan Pembayaran)
# ==========================================
# Map the EXCEL ROW NUMBER to the METRIC NAME in the database
ROW_MAP = {
    # Akaun Semasa / Current Account
    11: "akaun_semasa",
    12: "barangan",
    13: "perkhidmatan",
    14: "pendapatan_primer",
    15: "pendapatan_sekunder",
    
    # Akaun Modal / Capital Account
    17: "akaun_modal",
    
    # Akaun Kewangan
    19: "akaun_kewangan",
    
    # Pelaburan langsung
    21: "pelaburan_langsung",
    22: "pelaburan_langsung_di_luar_negeri",
    23: "pelaburan_langsung_asing_di_malaysia",
    
    # Kedudukan Pelaburan Antarabangsa
    25: "kedudukan_pelaburan_antarabangsa",
    
    # Pelaburan Swasta Diluluskan (MIDA) / Approved Private Investment
    27: "pelaburan_diluluskan",
    28: "johor",
    29: "kedah",
    30: "kelantan",
    31: "melaka",
    32: "negeri_sembilan",
    33: "pahang",
    34: "pulau_pinang",
    35: "perak",
    36: "perlis",
    37: "selangor",
    38: "terengganu",
    39: "sabah",
    40: "sarawak",
    41: "wp_kuala_lumpur",
    42: "wp_labuan",
    43: "wp_putrajaya"
}

# Map the EXCEL COLUMN NUMBER to the YEAR STRING
COL_MAP = {
    8  : "2023",  
    9  : "2024",  
    10 : "2025"   
}

# ==========================================
# REPORT INJECTION ENGINE
# ==========================================
def populate_jadual_19_0(sheet, hierarchy, report_type):
    print(f"  -> Populating Jadual 19.0 (Imbangan Pembayaran) untuk Malaysia_19_0")
    
    # 1. Fetch the Data Payload
    # SPECIAL CASE: Pointing to fact_metrics_negeri by setting level='negeri'
    metrics_data = get_metrics_dict("Malaysia", level='negeri')
    
    if not metrics_data:
        print(f"     [Warning] No data found for Malaysia in fact_metrics_negeri.")
        return

    # ==========================================
    # DYNAMIC TABLE TITLE MODIFICATION
    # ==========================================
    title_bm = ": Statistik imbangan pembayaran, Malaysia, 2023 - 2025"
    title_en = ": Balance of payments statistics, Malaysia, 2023 - 2025"
        
    # Set the exact cells where your title sits in the template
    sheet.range("C3").value = title_bm
    sheet.range("C4").value = title_en

    # 2. Standard Injection Loop
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