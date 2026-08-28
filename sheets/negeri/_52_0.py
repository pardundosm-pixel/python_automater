import pandas as pd
from src.data_provider import get_metrics_dict
from src.excel_utils import safe_write, inject_static_table

# ==========================================
# 1. MAPPING & PAGINATION CONFIGURATION
# ==========================================
# This table has no district loop or pagination - it's a single-page item table
# pulling directly from fact_metrics_negeri for the report's own state.

COL_YEAR = 4          # Column D
COL_JUMLAH = 11        # Column K - Jumlah/Total
COL_BANDAR = 12         # Column L - Bandar/Urban
COL_LUAR_BANDAR = 13     # Column M - Luar bandar/Rural

# The dictionary key (0, 1, 2) represents how many rows down from the item's start row the year sits.
YEAR_OFFSETS = {
    0: "2023",
    1: "2024",
    2: "2025"
}

# Row where each ICT item's block starts, mapped to the metric name suffix
# used in fact_metrics_negeri (e.g. "jumlah_telefon_bimbit", "bandar_telefon_bimbit"...)
ROW_MAP = {
    9: "telefon_bimbit",
    13: "internet",
    17: "komputer",
    21: "siaran_televisyen_berbayar",
    25: "televisyen",
    29: "radio",
    33: "telefon_talian_tetap",
}

TITLE_COORDINATES = {"bm": (1,3), "en": (2,3)}

# ==========================================
# SANITIZER HELPER
# ==========================================
def sanitize_value(val):
    """Safely attempts to convert a value to float, catching dashes and spaces."""
    if pd.notna(val):
        clean_val = str(val).strip()
        if clean_val not in ["", "n.a", "n.a.", "-", "na"]:
            try:
                return float(clean_val)
            except (ValueError, TypeError):
                pass
    return "n.a"

# ==========================================
# REPORT INJECTION ENGINE
# ==========================================
def populate_jadual_52(sheet, hierarchy, report_type):
    state_name = hierarchy.get('state_name', 'Unknown State')
    state_code = hierarchy.get('state_code', '00')

    title_bm = f": Peratusan capaian isi rumah terhadap perkhidmatan dan peralatan ICT mengikut strata, {state_name}, 2023 - 2025"
    title_en = f": Percentage of households with access to ICT services and equipment by strata, {state_name}, 2023 - 2025"

    safe_write(sheet, TITLE_COORDINATES["bm"][0], TITLE_COORDINATES["bm"][1], title_bm)
    safe_write(sheet, TITLE_COORDINATES["en"][0], TITLE_COORDINATES["en"][1], title_en)

    state_metrics = get_metrics_dict(state_code, level='negeri')

    for item_row, item_name in ROW_MAP.items():
        for offset, year in YEAR_OFFSETS.items():
            target_row = item_row + offset
            year_data = state_metrics.get(year, {})

            raw_jumlah = year_data.get(f"jumlah_{item_name}", "n.a")
            raw_bandar = year_data.get(f"bandar_{item_name}", "n.a")
            raw_luar_bandar = year_data.get(f"luar_bandar_{item_name}", "n.a")

            safe_write(sheet, target_row, COL_JUMLAH, sanitize_value(raw_jumlah))
            safe_write(sheet, target_row, COL_BANDAR, sanitize_value(raw_bandar))
            safe_write(sheet, target_row, COL_LUAR_BANDAR, sanitize_value(raw_luar_bandar))