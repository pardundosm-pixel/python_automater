import pandas as pd
from src.data_provider import get_metrics_dict

# ==========================================
# MAPPING CONFIGURATION FOR JADUAL 15.0 (Pasaran Kewangan)
# ==========================================
# TODO: Map the EXCEL ROW NUMBER to the METRIC NAME in the database
# Example: 7: "jumlah_penduduk", 11: "warganegara"
ROW_MAP = {
    # Kewangan Kerajaan (RM juta)
    10:  "kewangan_kerajaan_hasil",
    11:  "kewangan_kerajaan_perbelanjaan_mengurus",
    12:  "kewangan_kerajaan_baki_semasa",
    13:  "kewangan_kerajaan_perbelanjaan_pembangunan",
    14:  "kewangan_kerajaan_kumpulan_wang_covid_19",
    15:  "kewangan_kerajaan_baki_keseluruhan",
    16:  "kewangan_kerajaan_peratus_kdnk",
    
    # Hasil Terperinci Kerajaan Persekutuan (RM juta)
    18:  "kewangan_kerajaan_hasil_terperinci_kerajaan_persekutuan",
    
    # Jumlah cukai langsung
    20:  "jumlah_cukai_langsung",
    21:  "jumlah_cukai_langsung_cukai_pendapatan_individu",
    22:  "jumlah_cukai_langsung_cukai_pendapatan_syarikat",
    23:  "jumlah_cukai_langsung_cukai_pendapatan_petroleum",
    24:  "jumlah_cukai_langsung_cukai_pendapatan_koperasi",
    25:  "jumlah_cukai_langsung_cukai_pegangan",
    26:  "jumlah_cukai_langsung_lain_lain",
    
    # Jumlah cukai tidak langsung
    28:  "jumlah_cukai_tidak_langsung",
    29:  "jumlah_cukai_tidak_gregat_kewangan_langsung_gst",
    30:  "jumlah_cukai_tidak_langsung_duti_eksais",
    31:  "jumlah_cukai_tidak_langsung_duti_import",
    32:  "jumlah_cukai_tidak_langsung_duti_eksport",
    33:  "jumlah_cukai_tidak_langsung_lain_lain",
    
    # Hasil bukan cukai
    35:  "hasil_bukan_cukai",
    36:  "hasil_bukan_cukai_lesen_dan_permit",
    37:  "hasil_bukan_cukai_pendapatan_pelaburan",
    38:  "hasil_bukan_cukai_lain_lain",
    
    # Agregat Kewangan (RM juta)
    41:  "agregat_kewangan_m1",
    42:  "agregat_kewangan_m2",
    43:  "agregat_kewangan_m3",
    
    # Kadar Dasar Semalaman pada akhir tempoh (%)
    45:  "kadar_dasar_semalaman_pada_akhir_tempoh",
    
    # Kadar Faedah Purata Institusi Perbankan (%)
    48:  "kadar_faedah_purata_institusi_perbankan_kadar_faedah_deposit_tabungan_bank_perdagangan",
    49:  "kadar_faedah_purata_institusi_perbankan_deposit_tetap_12_bulan",
    50:  "kadar_faedah_purata_institusi_perbankan_kadar_berian_pinjaman_purata_bank_perdagangan",
    
    # Indikator Utama Bursa Malaysia
    53: "fte_bursa_malaysia_klci_fbm",
    54: "indikator_utama_bursa_malaysia_nilai_pasaran",
    55: "indikator_utama_bursa_malaysia_nilai_dagangan",
    56: "indikator_utama_bursa_malaysia_volum_dagangan"
}

# TODO: Map the EXCEL COLUMN NUMBER to the YEAR STRING
COL_MAP = {
    6 : "2023",  
    7 : "2024",  
    8 : "2025"   
}

def populate_jadual_15_0(sheet, hierarchy, report_type):
    print(f"  -> Populating Jadual 15.0 (Pasaran Kewangan) untuk Malaysia_15_0")
    
    # 1. Fetch the Data Payload strictly for Malaysia
    metrics_data = get_metrics_dict("Malaysia", level='malaysia')
    
    if not metrics_data:
            print(f"     [Warning] No data found for Malaysia.")
            return

    # ==========================================
        # DYNAMIC TABLE TITLE MODIFICATION
    # ==========================================
    title_bm = ": Statistik Kewangan Kerajaan, Pasaran Kewangan dan Modal, Malaysia, 2023 - 2025"
    title_en = ": Statistics of Government Finance, Financial and Capital Market, Malaysia, 2023 - 2025"
        
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