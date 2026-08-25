import pandas as pd
from src.data_provider import get_metrics_dict

# ============================================================
# 1. SHARED CONFIGURATION (change years here only)
# ============================================================
YEARS = ["2023", "2024", "2025"]

# ============================================================
# 2. JADUAL 34.0 – TABLE‑SPECIFIC CONFIGURATION
# ============================================================

# CHANGE HERE: Map each Excel row number to the ICT category.
ROW_MAP = {
    9 : "telefon_bimbit",
    12: "internet",
    14: "komputer",
    17: "siaran_televisyen_berbayar",
    20: "televisyen",
    23: "radio",
    25: "telefon_talian_tetap"
}

# CHANGE HERE: List all column indices for (2023 Total, Urban, Rural),
# then (2024 Total, Urban, Rural), then (2025 Total, Urban, Rural).
# The order must match the years in YEARS.
COLUMN_ORDER = [
    5, 6, 7,     # 2023: Total, Bandar, Luar bandar
    9, 10, 11,   # 2024: Total, Bandar, Luar bandar (column 11 is skipped)
    13, 14, 15   # 2025: Total, Bandar, Luar bandar
]

# Automatically build COL_MAP – no year strings repeated!
STRATA = ["jumlah", "bandar", "luar_bandar"]
COL_MAP = {}
for year_idx, year in enumerate(YEARS):
    for strata_idx, strata in enumerate(STRATA):
        col = COLUMN_ORDER[year_idx * 3 + strata_idx]
        COL_MAP[col] = (year, strata)


# ============================================================
# 3. INJECTION ENGINE
# ============================================================
def populate_jadual_34(sheet, hierarchy, report_type):
    print("  -> Populating Jadual 34.0 (Peratusan capaian ICT)")

    # Titles – year range is auto‑generated from YEARS
    title_bm = f": Peratusan capaian isi rumah terhadap perkhidmatan dan peralatan ICT mengikut strata, Malaysia, {YEARS[0]} - {YEARS[-1]}"
    title_en = f": Percentage of households with access to ICT services and equipment by strata, Malaysia, {YEARS[0]} - {YEARS[-1]}"
    sheet.range("C1").value = title_bm
    sheet.range("C2").value = title_en

    # ==========================================================
    # DATA FETCHING – Malaysia data is stored under code "00" as negeri
    # ==========================================================
    malaysia_data = get_metrics_dict("00", level="negeri")
    if not malaysia_data:
        print("     [Warning] No data found for Malaysia (00).")
        return

    # Inject data
    for row_idx, category in ROW_MAP.items():
        for col_idx, (year, strata) in COL_MAP.items():
            metric_name = f"{strata}_{category}"
            year_data = malaysia_data.get(year, {})
            raw_val = year_data.get(metric_name, "n.a")
            val = "n.a"
            if pd.notna(raw_val) and str(raw_val).strip() not in ("", "n.a", "n.a."):
                try:
                    val = float(raw_val)
                except (ValueError, TypeError):
                    pass
            sheet.range((row_idx, col_idx)).value = val