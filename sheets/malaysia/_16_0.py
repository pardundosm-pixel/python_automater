import pandas as pd
from src.data_provider import get_metrics_dict

# ==========================================
# MAPPING CONFIGURATION FOR JADUAL 16.0 (Akaun SP)
# ==========================================
# TODO: Map the EXCEL ROW NUMBER to the METRIC NAME in the database
# Example: 7: "jumlah_penduduk", 11: "warganegara"
ROW_MAP = {
    # Akaun Satelit Pelancongan (Nilai)
    11:  "akaun_satelit_pelancongan_nilai_nilai_ditambah_kasar_industri_pelancongan",
    14:  "akaun_satelit_pelancongan_nilai_keluaran_dalam_negeri_kasar_pelancongan_langsung",
    17:  "akaun_satelit_pelancongan_nilai_perbelanjaan_pelancongan_inbound_bagi_pelawat",
    20:  "akaun_satelit_pelancongan_nilai_perbelanjaan_pelancongan_domestik_bagi_pelawat",
    23:  "akaun_satelit_pelancongan_nilai_perbelanjaan_pelancongan_outbound_bagi_pelawat",
    26:  "akaun_satelit_pelancongan_nilai_guna_tenaga_dalam_industri_pelancongan",
    
    # Akaun Satelit Pelancongan - Perubahan Peratusan Tahunan (%)
    32:  "akaun_satelit_pelancongan_perubahan_peratusan_tahunan__nilai_ditambah_kasar_industri_pelancongan",
    35:  "akaun_satelit_pelancongan_perubahan_peratusan_tahunan__keluaran_dalam_negeri_kasar_pelancongan_langsung",
    38:  "akaun_satelit_pelancongan_perubahan_peratusan_tahunan__perbelanjaan_pelancongan_inbound_bagi_pelawat",
    41:  "akaun_satelit_pelancongan_perubahan_peratusan_tahunan__perbelanjaan_pelancongan_domestik_bagi_pelawat",
    44:  "akaun_satelit_pelancongan_perubahan_peratusan_tahunan__perbelanjaan_pelancongan_outbound_bagi_pelawat",
    47:  "akaun_satelit_pelancongan_perubahan_peratusan_tahunan__guna_tenaga_dalam_industri_pelancongan",
    
    #Akaun Satelit Pelancongan - Peratus Sumbangan kepada KDNK (%)
    54:  "peratus_sumbangan_kepada_keluaran_dalam_negeri_kasar_keluaran_dalam_negeri_kasar_pelancongan_langsung",
    57:  "peratus_sumbangan_kepada_keluaran_dalam_negeri_kasar_nilai_ditambah_kasar_industri_pelancongan"
}

# TODO: Map the EXCEL COLUMN NUMBER to the YEAR STRING
COL_MAP = {
    6 : "2023",  
    7 : "2024",  
    8 : "2025"   
}

def populate_jadual_16_0(sheet, hierarchy, report_type):
    print(f"  -> Populating Jadual 16.0 (Akaun SP) untuk Malaysia_16_0")
    
    # 1. Fetch the Data Payload strictly for Malaysia
    metrics_data = get_metrics_dict("Malaysia", level='malaysia')
    
    if not metrics_data:
            print(f"     [Warning] No data found for Malaysia.")
            return

    # ==========================================
        # DYNAMIC TABLE TITLE MODIFICATION
    # ==========================================
    title_bm = ": Statistik utama Akaun Satelit Pelancongan (ASP), Malaysia, 2023 - 2025"
    title_en = ": Principal statistics Tourism Satellite Account (TSA), Malaysia, 2023 - 2025"
        
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