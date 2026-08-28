import pandas as pd
from src.data_provider import get_metrics_dict
from src.excel_utils import safe_write

# ==========================================
# 1. MAPPING CONFIGURATION FOR JADUAL 43.0 (NEGERI)
# ==========================================
# Map the EXCEL ROW NUMBER to the METRIC NAME in the database
ROW_MAP = {

    9:  "jumlah_terimaan",
    11: "jumlah_pelawat_domestik",
    12: "isi_rumah_yang_dilawati",
    
    14: "jumlah_pelawat",
    16: "pelawat_harian",
    17: "pelancong",

    19: "jumlah_perjalanan_pelancongan",

    22: "destinasi_1",
    23: "destinasi_2",
    24: "destinasi_3",
    25: "destinasi_4",
    26: "destinasi_5"
}

# Map the EXCEL COLUMN NUMBER to the YEAR STRING
# Based on screenshot: Column G = 7, Column H = 8, Column I = 9
COL_MAP = {
    5: "2023",  
    6: "2024",  
    7: "2025"   
}

# ==========================================
# 2. REPORT INJECTION ENGINE
# ==========================================
def populate_jadual_44_0(sheet, hierarchy, report_type):
    state_name = hierarchy.get('state_name', 'Unknown State')
    state_code = hierarchy.get('state_code')
    print(f"  -> Populating Jadual 42.0 (KDNK) untuk {state_name}")

    metrics_data = get_metrics_dict(state_code, level='negeri')
    if not metrics_data:
        print(f"     [Warning] No data found for {state_name}.")
        return

    # 1. Inject Titles Safely (Row, Column)
    title_bm = f": Keluaran Dalam Negeri Kasar (KDNK), {state_name}, 2023 - 2025p"
    title_en = f": Gross Domestic Product (GDP), {state_name}, 2023 - 2025p"
    safe_write(sheet, 3, 3, title_bm) # C3
    safe_write(sheet, 4, 3, title_en) # C4