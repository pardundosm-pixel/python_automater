import pandas as pd
from src.data_provider import get_metrics_dict

# ==========================================
# MAPPING CONFIGURATION FOR JADUAL 14.2 (MALAYSIA)
# ==========================================
# TODO: Map the EXCEL ROW NUMBER to the METRIC NAME in the database
# Example: 7: "jumlah_penduduk", 11: "warganegara"
ROW_MAP = {
    # KDNK pada harga pembeli (RM juta)
    9:  "kdnk_pada_harga_pembeli_jumlah",
    12: "kdnk_pada_harga_pembeli_pampasan_pekerja",
    13: "kdnk_pada_harga_pembeli_lebihan_kendalian_kasar",
    14: "kdnk_pada_harga_pembeli_cukai_tolak_subsidi_ke_atas_pengeluaran_dan_import",
    
    # KDNK pada harga pembeli - Perubahan Peratusan Tahunan (%)
    16: "",
    19: "kdnk_perubahan_peratusan_tahunan_pampasan_pekerja",
    20: "kdnk_perubahan_peratusan_tahunan_lebihan_kendalian_kasar",
    21: "kdnk_perubahan_peratusan_tahunan_cukai_tolak_subsidi_ke_atas_pengeluaran_dan_import",
    
    # KDNK pada harga pembeli - Peratus Sumbangan kepada KDNK (%)
    23: "peratus_sumbangan_kepada_kdnk_jumlah",
    26: "peratus_sumbangan_kepada_kdnk_pampasan_pekerja",
    27: "peratus_sumbangan_kepada_kdnk_lebihan_kendalian_kasar",
    28: "peratus_sumbangan_kepada_kdnk_cukai_tolak_subsidi_ke_atas_pengeluaran_dan_import",
    
    # Nilai Ditambah PMKS Pada Harga Malar 2015 (RM juta)
    30: "nilai_ditambah_pmks_pada_harga_malar_2015",
    
    # Eksport PMKS Barangan dan Perkhidmatan (RM juta)
    33: "eksport_pmks_barangan_dan_perkhidmatan",
    
    # Guna Tenaga PMKS ('000)
    36: "guna_tenaga_pmks",
    
    # Nilai Ditambah PMKS Pada Harga Malar 2015 - Perubahan Peratusan Tahunan (%)
    39: "nilai_ditambah_pmks_pada_harga_malar_2015_perubahan_peratusan_tahunan",
    
    # Eksport PMKS Barangan dan Perkhidmatan - Perubahan Peratusan Tahunan (%)
    42: "eksport_pmks_barangan_dan_perkhidmatan_perubahan_peratusan_tahunan",
    
    # Guna Tenaga PMKS - Perubahan Peratusan Tahunan (%)
    45: "gtpmks_perubahan_peratusan_tahunan",
    
    # Nilai Ditambah PMKS Pada Harga Malar 2015 - Peratus Sumbangan kepada KDNK (%)
    48: "nilai_ditambah_pmks_pada_harga_malar_2015_peratus_sumbangan_kepada_kdnk",
    
    # Eksport PMKS Barangan dan Perkhidmatan - Peratus Sumbangan kepada Jumlah Eksport Barangan dan Perkhidmatan (%)
    51: "eksport_pmks_barangan_dan_perkhidmatan_peratus_sumbangan_kepada_kdnk",
    
    # Guna Tenaga PMKS - Peratus Sumbangan kepada Guna Tenaga Malaysia (%)
    54: "guna_tenaga_pmks_peratus_sumbangan_kepada_kdnk",
    
    # Pembentukan Modal Tetap Kasar (PMTK) - RM juta
    58: "kdnk_pembentukan_modal_tetap_kasar",
    
    # Pembentukan Modal Tetap Kasar (PMTK) - Perubahan Peratusan Tahunan (%)
    61: "pembentukan_modal_tetap_kasar_perubahan_peratusan_tahunan",

}

# TODO: Map the EXCEL COLUMN NUMBER to the YEAR STRING
COL_MAP = {
    8 : "2023",  
    9 : "2024",  
    10: "2025"   
}

def populate_jadual_14_2(sheet, hierarchy, report_type):
    print(f"  -> Populating Jadual 14.2 (KDNK) untuk Malaysia_14_2")
    
    # 1. Fetch the Data Payload strictly for Malaysia
    metrics_data = get_metrics_dict("Malaysia", level='malaysia')
    
    if not metrics_data:
            print(f"     [Warning] No data found for Malaysia.")
            return

    # ==========================================
        # DYNAMIC TABLE TITLE MODIFICATION
    # ==========================================
    title_bm = ": Keluaran Dalam Negeri Kasar (KDNK), Malaysia, 2023 - 2025 (samb.)"
    title_en = ": Gross Domestic Product (GDP), Malaysia, 2023 - 2025 (cont'd)"
        
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