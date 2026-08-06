import pandas as pd
from src.data_provider import get_metrics_dict

# ==========================================
# 1. MAPPING & PAGINATION CONFIGURATION
# ==========================================
# Define the exact names of the sheets in the template handling this table
PAGINATED_SHEETS = ["42.1", "42.1 (2)", "42.1 (3)"]

# Grid layout definition
START_ROW = 11        # The first row of data on every sheet
ROWS_PER_BLOCK = 4    # 3 years of data + 1 blank separator row
MAX_BLOCKS_PER_SHEET = 6 # E.g., Rows 11, 15, 19, 23, 27, 31

# Column Indices mapping
COL_DISTRICT = 2  # Column B
COL_YEAR = 5      # Column E
COL_VALUE = 8     # Column H (Nilai)
COL_PCT = 9       # Column I (Peratus)

# Offset mapping for the 3 years within a district block
YEAR_OFFSETS = {
    0: "2018",
    1: "2019",
    2: "2020"
}

# ==========================================
# 2. REPORT INJECTION ENGINE
# ==========================================
def populate_jadual_42_1(sheet, hierarchy, report_type):
    # Prevent the engine from running this logic 3 separate times if it loops through all sheets.
    # We will orchestrate all 3 sheets directly from the first sheet's trigger.
    if sheet.name != PAGINATED_SHEETS[0]:
        return

    state_name = hierarchy.get('state_name', 'Unknown State')
    state_code = hierarchy.get('state_code')
    print(f"  -> Populating Jadual 42.1 (KDNK Daerah) across multiple sheets untuk {state_name}")

    wb = sheet.book
    
    # 1. Fetch the districts for this specific state
    # (Assuming a helper function exists in data_provider to list a state's districts)
    from src.data_provider import get_districts_for_state
    districts = get_districts_for_state(state_code) 
    
    if not districts:
        print(f"     [Warning] No districts found for {state_name}.")
        return

    # 2. Setup Title Strings
    title_bm = f": Keluaran Dalam Negeri Kasar (KDNK) mengikut daerah pentadbiran, {state_name}, 2018 - 2020"
    title_en = f": Gross Domestic Product (GDP) by administrative district, {state_name}, 2018 - 2020"
    
    title_bm_samb = title_bm + " (samb.)"
    title_en_samb = title_en + " (cont'd)"

    district_idx = 0
    total_districts = len(districts)

    # 3. Iterate through the designated sheets
    for sheet_idx, sheet_name in enumerate(PAGINATED_SHEETS):
        try:
            current_sheet = wb.sheets[sheet_name]
        except Exception:
            continue # Skip if the template doesn't have this specific continuation sheet

        # Update Titles for the current sheet
        if sheet_idx == 0:
            current_sheet.range("C2").value = title_bm
            current_sheet.range("C3").value = title_en
        else:
            current_sheet.range("C2").value = title_bm_samb
            current_sheet.range("C3").value = title_en_samb

        # Check if we have already plotted all districts. If yes, hide this sheet and skip.
        if district_idx >= total_districts:
            current_sheet.api.Visible = False # Hide unused template sheets from the final Excel
            continue

        # 4. Inject Data into Blocks
        for block_idx in range(MAX_BLOCKS_PER_SHEET):
            base_row = START_ROW + (block_idx * ROWS_PER_BLOCK)
            
            if district_idx < total_districts:
                # --- INJECT DISTRICT DATA ---
                district_info = districts[district_idx]
                dist_code = district_info['code']
                dist_name = district_info['name']
                
                # Write District Name (Only on the first row of the block)
                current_sheet.range((base_row, COL_DISTRICT)).value = dist_name
                
                # Fetch metrics for this specific district
                metrics_data = get_metrics_dict(dist_code, level='daerah')
                
                # Loop through the 3 years
                for offset, year in YEAR_OFFSETS.items():
                    target_row = base_row + offset
                    year_data = metrics_data.get(year, {})
                    
                    # Extract Data
                    val_kdnk = year_data.get("kdnk_harga_malar_nilai", "n.a")
                    val_pct = year_data.get("kdnk_harga_malar_peratus", "n.a")
                    
                    # Sanitize and write Value
                    current_sheet.range((target_row, COL_VALUE)).value = float(val_kdnk) if pd.notna(val_kdnk) and val_kdnk not in ["n.a", ""] else "n.a"
                    
                    # Sanitize and write Percentage
                    current_sheet.range((target_row, COL_PCT)).value = float(val_pct) if pd.notna(val_pct) and val_pct not in ["n.a", ""] else "n.a"
                
                district_idx += 1
                
            else:
                # --- CLEANUP UNUSED DUMMY BLOCKS ---
                # Wipe the District Name, Years, and Values to leave a blank grid
                current_sheet.range((base_row, COL_DISTRICT), (base_row + 2, COL_PCT)).value = None