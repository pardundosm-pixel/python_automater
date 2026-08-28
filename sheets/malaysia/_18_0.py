
from src.data_provider import get_metrics_dict
from src.excel_utils import inject_static_table

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