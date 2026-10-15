import pandas as pd
from src.data_provider import get_metrics_dict

# ==========================================
# MAPPING CONFIGURATION FOR JADUAL 18.0 (Stok Modal)
# ==========================================
# TODO: Map the EXCEL ROW NUMBER to the METRIC NAME in the database
# Example: 7: "jumlah_penduduk", 11: "warganegara"
ROW_MAP = {
    # Stok Modal apda harga malar 2015 (RM juta)
    11:  "jumlah_harga_malar_stok_modal_kasar",
    14:  "jumlah_harga_malar_stok_modal_bersih",
    17:  "jumlah_harga_malar_stok_modal_produktif",
    20:  "jumlah_harga_malar_penggunaan_modal_tetap",
    
    # Stok Modal pada harga malar 2015 - Perubahan Peratusan Tahunan (%)
    27:  "peratus_perubahan_peratusan_tahunan_stok_modal_kasar",
    30:  "peratus_perubahan_peratusan_tahunan_stok_modal_bersih",
    33:  "peratus_perubahan_peratusan_tahunan_stok_modal_produktif",
    36:  "peratus_perubahan_peratusan_tahunan_penggunaan_modal_tetap"

}

# TODO: Map the EXCEL COLUMN NUMBER to the YEAR STRING
COL_MAP = {
    6 : "2023",  
    7 : "2024",  
    8 : "2025"   
}

def populate_jadual_18_0(sheet, hierarchy, report_type):
    print(f"  -> Populating Jadual 18.0 (Stok Modal) untuk Malaysia_18_0")
    
    # 1. Fetch the Data Payload strictly for Malaysia
    metrics_data = get_metrics_dict("Malaysia", level='malaysia')
    
    if not metrics_data:
            print(f"     [Warning] No data found for Malaysia.")
            return

    # ==========================================
        # DYNAMIC TABLE TITLE MODIFICATION
    # ==========================================
    title_bm = ": Statistik utama Stok Modal, Malaysia, 2023 - 2025"
    title_en = ": Principal statistics of Capital Stock, Malaysia, 2023 - 2025"
        
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