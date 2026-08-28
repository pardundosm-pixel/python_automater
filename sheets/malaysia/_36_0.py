
from src.data_provider import get_metrics_dict
from src.excel_utils import safe_write

# ==========================================
# 1. MAPPING CONFIGURATION FOR JADUAL 36.0
# ==========================================
# This table is strictly locked to 2024
TARGET_YEAR = "2024"

# Map Excel Row Index to a dictionary of Column Indices and their exact DB Metric Names
# E=5 (pengguna_per_kapita), G=7 (kadar_sara_diri), I=9 (kadar_kebergantungan_import)
# IMPORTANT: Replace the dummy string values (e.g., "pengguna_per_kapita_daging_ayam") 
# with the exact metric names used in your fact_metrics_malaysia database.
ROW_MAP = {
    13: {5: "pengguna_per_kapita_daging_ayam", 7: "kadar_sara_diri_daging_ayam", 9: "kadar_kebergantungan_import_daging_ayam"},
    15: {5: "pengguna_per_kapita_telur_ayam_itik",       7: "kadar_sara_diri_telur_ayam_itik",       9: "kadar_kebergantungan_import_telur_ayam_itik"},
    17: {5: "pengguna_per_kapita_kelapa",      7: "kadar_sara_diri_kelapa",      9: "kadar_kebergantungan_import_kelapa"},
    19: {5: "pengguna_per_kapita_daging_babi",        7: "kadar_sara_diri_daging_babi",        9: "kadar_kebergantungan_import_daging_babi"},
    21: {5: "pengguna_per_kapita_durian",      7: "kadar_sara_diri_durian",      9: "kadar_kebergantungan_import_durian"},
    23: {5: "pengguna_per_kapita_bawang_besar",      7: "kadar_sara_diri_bawang_besar",      9: "kadar_kebergantungan_import_bawang_besar"},
    25: {5: "pengguna_per_kapita_pisang",      7: "kadar_sara_diri_pisang",      9: "kadar_kebergantungan_import_pisang"},
    27: {5: "pengguna_per_kapita_nanas",       7: "kadar_sara_diri_nanas",       9: "kadar_kebergantungan_import_nanas"},
    29: {5: "pengguna_per_kapita_kobis_bulat",       7: "kadar_sara_diri_kobis_bulat",       9: "kadar_kebergantungan_import_kobis_bulat"},
    31: {5: "pengguna_per_kapita_daging_lembu_kerbau",       7: "kadar_sara_diri_daging_lembu_kerbau",       9: "kadar_kebergantungan_import_daging_lembu_kerbau"},
    33: {5: "pengguna_per_kapita_bawang_putih",7: "kadar_sara_diri_bawang_putih",9: "kadar_kebergantungan_import_bawang_putih"},
    35: {5: "pengguna_per_kapita_mackerel",    7: "kadar_sara_diri_mackerel",    9: "kadar_kebergantungan_import_mackerel"},
    37: {5: "pengguna_per_kapita_sawi",        7: "kadar_sara_diri_sawi",        9: "kadar_kebergantungan_import_sawi"},
    39: {5: "pengguna_per_kapita_udang",       7: "kadar_sara_diri_udang",       9: "kadar_kebergantungan_import_udang"},
    41: {5: "pengguna_per_kapita_tomato",      7: "kadar_sara_diri_tomato",      9: "kadar_kebergantungan_import_tomato"},
    43: {5: "pengguna_per_kapita_timun",       7: "kadar_sara_diri_timun",       9: "kadar_kebergantungan_import_timun"},
    45: {5: "pengguna_per_kapita_tembikai",    7: "kadar_sara_diri_tembikai",    9: "kadar_kebergantungan_import_tembikai"},
    47: {5: "pengguna_per_kapita_epal",        7: "kadar_sara_diri_epal",        9: "kadar_kebergantungan_import_epal"},
    49: {5: "pengguna_per_kapita_salad",       7: "kadar_sara_diri_salad",       9: "kadar_kebergantungan_import_salad"},
    51: {5: "pengguna_per_kapita_selayang",     7: "kadar_sara_diri_selayang",     9: "kadar_kebergantungan_import_selayang"},
    53: {5: "pengguna_per_kapita_sotong",     7: "kadar_sara_diri_sotong",     9: "kadar_kebergantungan_import_sotong"},
}

# ==========================================
# 2. REPORT INJECTION ENGINE
# ==========================================
def populate_jadual_36_0(sheet, hierarchy, report_type):
    print("  -> Populating Jadual 36.0 (Statistik Pertanian) untuk Malaysia")

    metrics_data = get_metrics_dict("Malaysia", level='malaysia')
    
    if not metrics_data:
        print(f"     [Warning] No data found for Malaysia.")
        return

    year_data = metrics_data.get(TARGET_YEAR, {})
    
    if not year_data:
        print(f"     [Warning] No data found for the year {TARGET_YEAR}.")

    # Titles (Openpyxl syntax)
    sheet["C3"] = ": Statistik Penggunaan Per Kapita, Kadar Sara Diri dan Kadar Kebergantungan Import Item Pertanian,"
    sheet["C4"] = "  Malaysia, 2024"
    sheet["C5"] = ": Statistics of Per Capita Consumption, Self-Sufficiency Ratio and Import Dependency Ratio of"
    sheet["C6"] = "  Agricultural Items, Malaysia, 2024"

    for row_idx, columns_dict in ROW_MAP.items():
        for col_idx, metric_name in columns_dict.items():
            raw_val = year_data.get(metric_name, "n.a")
            
            if raw_val is not None and str(raw_val).strip() not in ["", "n.a", "n.a.", "nan", "NaN"]:
                try: 
                    val = float(raw_val)
                except (ValueError, TypeError): 
                    val = "n.a"
            else:
                val = "n.a"
                
            safe_write(sheet, row_idx, col_idx, val)