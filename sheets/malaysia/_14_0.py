# sheets/template_malaysia/_14_0.py
import pandas as pd
from src.data_provider import get_metrics_dict

# ==========================================
# MAPPING CONFIGURATION FOR JADUAL 14.0 (GDP)
# ==========================================
# Map the EXCEL ROW NUMBER to the METRIC NAME in the database
GDP_ROW_MAP = {
    10: "gdp_at_purchasers_prices",
    11: "agriculture",
    12: "mining_and_quarrying",
    13: "manufacturing",
    14: "construction",
    15: "services",
    16: "import_duties",
    
    21: "gdp_growth_rate",
    22: "agri_growth",
    23: "mining_growth",
    24: "manufacturing_growth",
    25: "construction_growth",
    26: "services_growth",
    27: "import_duties_growth",
    
    30: "gdp_at_purchasers_prices_exp", 
    32: "private_consumption",
    34: "government_consumption",
    35: "gross_fixed_capital_formation",
    36: "changes_in_inventories",
    37: "exports_goods_services"
}

# Map the EXCEL COLUMN NUMBER to the YEAR STRING
GDP_COL_MAP = {
    4: "2022",  # Column D
    5: "2023",  # Column E
    6: "2024"   # Column F
}

# ==========================================
# 1. REPORT INJECTION ENGINE
# ==========================================
# FIXED: Added 'report_type' to match the engine.py execution call
def populate_jadual_14(sheet, hierarchy, report_type="malaysia"):
    print(f"  -> Populating Jadual 14.0 (KDNK) for Malaysia")

    title_bm = f"Keluaran Dalam Negeri Kasar (KDNK), Malaysia, 2022 - 2024"
    title_en = f"Gross Domestic Product (GDP), Malaysia, 2022 - 2024"

    sheet.range("C2").value = title_bm
    sheet.range("C3").value = title_en

    # ==========================================
    # 2. FETCH THE DATA PAYLOAD
    # ==========================================
    # Fetch strictly Malaysia-wide data from fact_malaysia
    metrics_data = get_metrics_dict("Malaysia", level='malaysia')
    
    if not metrics_data:
        print(f"     [Warning] No GDP data found for Malaysia.")
        return

    # ==========================================
    # 3. INJECT DATA INTO EXCEL
    # ==========================================
    # Loop over the columns (Years) first
    for col_idx, year in GDP_COL_MAP.items():
        year_data = metrics_data.get(str(year), {})
        
        # Loop over the rows (Metrics) second
        for row_idx, metric_name in GDP_ROW_MAP.items():
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