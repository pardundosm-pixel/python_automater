
from src.data_provider import get_metrics_dict
from src.excel_utils import inject_static_table

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
    print("  -> Populating Jadual 14.0 (KDNK) untuk Malaysia")
    
    metrics_data = get_metrics_dict("Malaysia", level='malaysia')
    if not metrics_data:
        print("     [Warning] No data found for Malaysia.")
        return

    # Titles (Updated to Openpyxl syntax)
    sheet["C3"] = ": Keluaran Dalam Negeri Kasar (KDNK), Malaysia, 2023 - 2025"
    sheet["C4"] = ": Gross Domestic Product (GDP), Malaysia, 2023 - 2025"

    # Single-line data injection
    inject_static_table(sheet, metrics_data, ROW_MAP, COL_MAP)