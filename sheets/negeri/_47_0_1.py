import pandas as pd
from src.data_provider import get_metrics_dict

# ==========================================
# MAPPING CONFIGURATION FOR JADUAL 47.0 (samb.) (NEGERI)
# ==========================================
# Map the EXCEL ROW NUMBER to the METRIC NAME in the database
ROW_MAP = {
    # Kadar Kecederaan Pekerjaan (accident rate per 1,000 workers)
    9: "jumlah_kadar_kemalangan",
    12: "kadar_kemalangan_warganegara",
    13: "kadar_kemalangan_bukan_warganegara",
    16: "kadar_kemalangan_lelaki",
    17: "kadar_kemalangan_perempuan",
    21: "kadar_kemalangan_pertanian_perhutanan_dan_perikanan",
    22: "kadar_kemalangan_perlombongan_dan_pengkuarian",
    23: "kadar_kemalangan_pembuatan",
    24: "kadar_kemalangan_pembinaan",
    25: "kadar_kemalangan_utiliti",
    26: "kadar_kemalangan_perdagangan_borong_dan_runcit",
    27: "kadar_kemalangan_pengangkutan_penyimpanan_dan_komunikasi",
    29: "kadar_kemalangan_hotel_dan_restoran",
    30: "kadar_kemalangan_kewangan_insurans_hartanah_dan_perkhidmatan_perniagaan",
    32: "kadar_kemalangan_perkhidmatan",

    # Kadar Kematian Pekerjaan (fatality rate per 100,000 workers)
    39: "jumlah_kadar_kematian",
    42: "kadar_kematian_warganegara",
    43: "kadar_kematian_bukan_warganegara",
    46: "kadar_kematian_lelaki",
    47: "kadar_kematian_perempuan",
    51: "kadar_kematian_pertanian_perhutanan_dan_perikanan",
    52: "kadar_kematian_perlombongan_dan_pengkuarian",
    53: "kadar_kematian_pembuatan",
    54: "kadar_kematian_pembinaan",
    55: "kadar_kematian_utiliti",
    56: "kadar_kematian_perdagangan_borong_dan_runcit",
    57: "kadar_kematian_pengangkutan_penyimpanan_dan_komunikasi",
    59: "kadar_kematian_hotel_dan_restoran",
    60: "kadar_kematian_kewangan_insurans_hartanah_dan_perkhidmatan_perniagaan",
    62: "kadar_kematian_perkhidmatan"
}

# Map the EXCEL COLUMN NUMBER to the YEAR STRING
# Based on screenshot: Column J = 10, Column K = 11, Column L = 12
COL_MAP = {
    6: "2022",  
    7: "2023",  
    8: "2024"   
}

# ==========================================
# 2. REPORT INJECTION ENGINE
# ==========================================
def populate_jadual_47_0_1(sheet, hierarchy, report_type):
    state_name = hierarchy.get('state_name', 'Unknown State')
    state_code = hierarchy.get('state_code')
    print(f"  -> Populating Jadual 47.0 (Kemalangan Pekerjaan)(samb.) untuk {state_name}")

    # 1. Fetch the Data Payload for the specific Negeri
    metrics_data = get_metrics_dict(state_code, level='negeri')
    
    if not metrics_data:
        print(f"     [Warning] No data found for {state_name}.")
        return

    # ==========================================
    # DYNAMIC TABLE TITLE MODIFICATION
    # ==========================================
    title_bm = f": Statistik utama kecederaan pekerjaan, {state_name}, 2022 - 2024 (samb.)"
    title_en = f": Principal statistics of occupational injury, {state_name}, 2022 - 2024 (cont'd)"

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