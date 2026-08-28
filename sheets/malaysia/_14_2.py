
from src.data_provider import get_metrics_dict
from src.excel_utils import inject_static_table

# ==========================================
# MAPPING CONFIGURATION FOR JADUAL 14.2 (MALAYSIA)
# ==========================================
ROW_MAP = {
    # KDNK pada harga pembeli (RM juta)
    9:  "kdnk_pada_harga_pembeli_jumlah",
    12: "kdnk_pada_harga_pembeli_pampasan_pekerja",
    13: "kdnk_pada_harga_pembeli_lebihan_kendalian_kasar",
    14: "kdnk_pada_harga_pembeli_cukai_tolak_subsidi_ke_atas_pengeluaran_dan_import",

    # KDNK pada harga pembeli - Perubahan Peratusan Tahunan (%)
    16: "kdnk_perubahan_peratusan_tahunan_jumlah",
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

    # Eksport PMKS Barangan dan Perkhidmatan (RM Bilion)
    33: "eksport_pmks_barangan_dan_perkhidmatan",

    # Guna Tenaga PMKS ('000)
    36: "guna_tenaga_pmks",

    # Nilai Ditambah PMKS - Perubahan Peratusan Tahunan (%)
    39: "nilai_ditambah_pmks_pada_harga_malar_2015_perubahan_peratusan_tahunan",

    # Eksport PMKS - Perubahan Peratusan Tahunan (%)
    42: "eksport_pmks_barangan_dan_perkhidmatan_perubahan_peratusan_tahunan",

    # Guna Tenaga PMKS - Perubahan Peratusan Tahunan (%)
    45: "gtpmks_perubahan_peratusan_tahunan",

    # Nilai Ditambah PMKS - Peratus Sumbangan kepada KDNK (%)
    48: "nilai_ditambah_pmks_pada_harga_malar_2015_peratus_sumbangan_kepada_kdnk",

    # Eksport PMKS - Peratus Sumbangan kepada Jumlah Eksport (%)
    51: "eksport_pmks_barangan_dan_perkhidmatan_peratus_sumbangan_kepada_kdnk",

    # Guna Tenaga PMKS - Peratus Sumbangan kepada Guna Tenaga Malaysia (%)
    54: "guna_tenaga_pmks_peratus_sumbangan_kepada_kdnk",

    # Pembentukan Modal Tetap Kasar (PMTK) - RM juta
    58: "kdnk_pembentukan_modal_tetap_kasar",

    # Pembentukan Modal Tetap Kasar (PMTK) - Perubahan Peratusan Tahunan (%)
    61: "pembentukan_modal_tetap_kasar_perubahan_peratusan_tahunan",
}

# Map columns to years (column 8 = 2023, 9 = 2024, 10 = 2025)
COL_MAP = {
    8: "2023",
    9: "2024",
    10: "2025"
}

def populate_jadual_14_2(sheet, hierarchy, report_type):
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