import pandas as pd
from src.data_provider import get_metrics_dict

# ==========================================
# MAPPING CONFIGURATION FOR JADUAL 47.0 (NEGERI)
# ==========================================
# Map the EXCEL ROW NUMBER to the METRIC NAME in the database
ROW_MAP = {
    # Bilangan Kecederaan Pekerjaan (accident counts)
    9: "jumlah_kemalangan",

    12: "kemalangan_warganegara",
    13: "kemalangan_bukan_warganegara",

    16: "kemalangan_lelaki",
    17: "kemalangan_perempuan",

    21: "kemalangan_pertanian_perhutanan_dan_perikanan",
    22: "kemalangan_perlombongan_dan_pengkuarian",
    23: "kemalangan_pembuatan",
    24: "kemalangan_pembinaan",
    25: "kemalangan_utiliti",
    26: "kemalangan_perdagangan_borong_dan_runcit",
    27: "kemalangan_pengangkutan_penyimpanan_dan_komunikasi",
    29: "kemalangan_hotel_dan_restoran",
    30: "kemalangan_kewangan_insurans_hartanah_dan_perkhidmatan_perniagaan",
    32: "kemalangan_perkhidmatan",

    # Bilangan Kematian Pekerjaan (fatality counts)
    39: "jumlah_kematian",

    42: "kematian_warganegara",
    43: "kematian_bukan_warganegara",

    46: "kematian_lelaki",
    47: "kematian_perempuan",

    51: "kematian_pertanian_perhutanan_dan_perikanan",
    52: "kematian_perlombongan_dan_pengkuarian",
    53: "kematian_pembuatan",
    54: "kematian_pembinaan",
    55: "kematian_utiliti",
    56: "kematian_perdagangan_borong_dan_runcit",
    57: "kematian_pengangkutan_penyimpanan_dan_komunikasi",
    59: "kematian_hotel_dan_restoran",
    60: "kematian_kewangan_insurans_hartanah_dan_perkhidmatan_perniagaan",
    62: "kematian_perkhidmatan"
}

# Map the EXCEL COLUMN NUMBER to the YEAR STRING
# Based on screenshot: Column J = 10, Column K = 11, Column L = 12
COL_MAP = {
    5: "2022",  
    6: "2023",  
    7: "2024"   
}

# ==========================================
# 2. REPORT INJECTION ENGINE
# ==========================================
def populate_jadual_47_0(sheet, hierarchy, report_type):
    state_name = hierarchy.get('state_name', 'Unknown State')
    state_code = hierarchy.get('state_code')
    print(f"  -> Populating Jadual 47.0 (Kemalangan Pekerjaan) untuk {state_name}")

    # 1. Fetch the Data Payload for the specific Negeri
    metrics_data = get_metrics_dict(state_code, level='negeri')
    
    if not metrics_data:
        print(f"     [Warning] No data found for {state_name}.")
        return

    # ==========================================
    # DYNAMIC TABLE TITLE MODIFICATION
    # ==========================================
    title_bm = f": Statistik utama kecederaan pekerjaan, {state_name}, 2022 - 2024"
    title_en = f": Principal statistics of occupational injury, {state_name}, 2022 - 2024"

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