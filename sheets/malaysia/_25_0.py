import pandas as pd
from src.data_provider import get_metrics_dict

# ==========================================
# MAPPING CONFIGURATION FOR JADUAL 23.1 (Harga Item Terpilih)
# ==========================================
# TODO: Map the EXCEL ROW NUMBER to the METRIC NAME in the database
# Example: 7: "jumlah_penduduk", 11: "warganegara"
ROW_MAP = {
    # Jumlah Kecederaan Pekerjaan
    11: "jumlah_kemalangan",
    
    # Kewarganegaraan
    14: "kemalangan_warganegara",
    15: "kemalangan_bukan_warganegara",
    
    # Jantina
    18: "kemalangan_lelaki",
    19: "kemalangan_perempuan",
    
    # Sektor OSHA (Kecederaan)
    23: "kemalangan_pertanian_perhutanan_dan_perikanan",
    24: "kemalangan_perlombongan_dan_pengkuarian",
    25: "kemalangan_pembuatan",
    26: "kemalangan_pembinaan",
    27: "kemalangan_utiliti",
    28: "kemalangan_perdagangan_borong_dan_runcit",
    29: "kemalangan_pengangkutan_penyimpanan_dan_komunikasi",
    31: "kemalangan_hotel_dan_restoran",
    32: "kemalangan_kewangan_insurans_hartanah_dan_perkhidmatan_perniagaan",
    34: "kemalangan_perkhidmatan",
    
    # Jumlah Kematian Pekerjaan
    41: "jumlah_kematian_pekerjaan",
    
    # Kewarganegaraan (Kematian)
    44: "kematian_warganegara",
    45: "kematian_bukan_warganegara",
    
    # Jantina (Kematian)
    48: "kematian_lelaki",
    49: "kematian_perempuan",
    
    # Sektor OSHA (Kematian)
    53: "kematian_pertanian_perhutanan_dan_perikanan",
    54: "kematian_perlombongan_dan_pengkuarian",
    55: "kematian_pembuatan",
    56: "kematian_pembinaan",
    57: "kematian_utiliti",
    58: "kematian_perdagangan_borong_dan_runcit",
    59: "kematian_pengangkutan_penyimpanan_dan_komunikasi",
    61: "kematian_hotel_dan_restoran",
    62: "kematian_kewangan_insurans_hartanah_dan_perkhidmatan_perniagaan",
    64: "kematian_perkhidmatan"
}

# TODO: Map the EXCEL COLUMN NUMBER to the YEAR STRING
COL_MAP = {
    5  : "2022",  
    6  : "2023",  
    7  : "2024"   
}

def populate_jadual_25_0(sheet, hierarchy, report_type):
    print(f"  -> Populating Jadual 25.0 (Kemalangan Pekerja) untuk Malaysia_25_0")
    
    # 1. Fetch the Data Payload strictly for Malaysia
    metrics_data = get_metrics_dict("00", level='negeri')
    
    if not metrics_data:
            print(f"     [Warning] No data found for Malaysia.")
            return

    # ==========================================
        # DYNAMIC TABLE TITLE MODIFICATION
    # ==========================================
    title_bm  = ": Statistik utama kecederaan pekerjaan, Malaysia, 2022 - 2024"
    title_en  = ": Principal statistics of occupational injury , Malaysia, 2022 - 2024"

        
    # Set the exact cells where your title sits in the template
    # Targeting Column C based on standard template behavior
    sheet.range("C3").value = title_bm
    sheet.range("C4").value = title_en

    # Standard Injection Loop
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