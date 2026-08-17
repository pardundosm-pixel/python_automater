import pandas as pd
from src.data_provider import get_metrics_dict

# ==========================================
# MAPPING CONFIGURATION FOR JADUAL 26.0 (Peratusan Ahli Parlimen)
# ==========================================
# TODO: Map the EXCEL ROW NUMBER to the METRIC NAME in the database
# Example: 7: "jumlah_penduduk", 11: "warganegara"
ROW_MAP = {
    # Dewan Negara (Senate)
    9:  "dewan_negara_lelaki",      # Lelaki/Male
    10: "dewan_negara_perempuan",   # Perempuan/Female

    # Dewan Rakyat (House of Representatives)
    13: "dewan_rakyat_lelaki",      # Lelaki/Male
    14: "dewan_rakyat_perempuan",   # Perempuan/Female

    # Menteri Kabinet (Cabinet Minister)
    17: "menteri_kabinet_lelaki",   # Lelaki/Male
    18: "menteri_kabinet_perempuan",# Perempuan/Female

    # Timbalan Menteri (Deputy Minister)
    21: "timbalan_menteri_lelaki",  # Lelaki/Male
    22: "timbalan_menteri_perempuan"# Perempuan/Female
}

# TODO: Map the EXCEL COLUMN NUMBER to the YEAR STRING
COL_MAP = {
    8  : "2022",  
    9  : "2023",  
    10 : "2024"   
}

def populate_jadual_26_0(sheet, hierarchy, report_type):
    print(f"  -> Populating Jadual 26.0 (Peratusan Ahli Parlimen) untuk Malaysia_26_0")
    
    # 1. Fetch the Data Payload strictly for Malaysia
    metrics_data = get_metrics_dict("Malaysia", level='malaysia')
    
    if not metrics_data:
            print(f"     [Warning] No data found for Malaysia.")
            return

    # ==========================================
        # DYNAMIC TABLE TITLE MODIFICATION
    # ==========================================
    title_bm = ": Peratusan ahli parlimen dan anggota pentadbiran mengikut jantina, Malaysia, 2022 - 2024"
    title_en = ": Percentage of members of parliament and administration by sex, Malaysia, 2022 - 2024"
        
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