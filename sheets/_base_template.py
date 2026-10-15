"""
sheets/_base_template.py
========================
Shared mapper factories. Each jadual_N.py calls ONE of:

    populate_column = make_populate_location(ROW_MAP, ROUND_DECIMALS)
    populate_page   = make_populate_year(ROW_MAP, TARGET_YEARS, ROUND_DECIMALS)

The engine discovers whichever function exists and calls it accordingly.

THIS FILE IS NOT DISCOVERED BY THE ENGINE — its name doesn't start with
`jadual_`.
"""
from sheets import write_cell

# ==========================================
# MODE A — COLUMNS ARE LOCATIONS
# ==========================================
def make_populate_location(row_map: dict, round_decimals: int | None):
    """
    Build a `populate_column` for jaduals whose columns are locations.

    ROW_MAP = {row: (metric_name, year_key)}
    Called once per location column, per page.
    """
    def populate_column(ws, col_idx: int, metrics_by_year: dict,
                        row_offset: int = 0, bold: bool = False):
        for template_row, (metric_name, year) in row_map.items():
            actual_row = template_row + row_offset
            year_data  = metrics_by_year.get(str(year), {})
            val        = year_data.get(metric_name, "n.a")
            write_cell(ws, actual_row, col_idx, val,
                       decimals=round_decimals, bold=bold)
    return populate_column


# ==========================================
# MODE B — YEARS SPECIFIED IN CODE (RIGHT-ALIGNED)
# ==========================================
def make_populate_year(row_map: dict, target_years: list, round_decimals: int | None):
    """
    Build a `populate_page` where years are explicitly defined by the mapper.
    The years are injected starting from the rightmost column (COL_L) backwards.
    
    TARGET_YEARS = ["2021", "2022", "2023"]
    Result: COL_L (12) -> 2023, COL_K (11) -> 2022, COL_J (10) -> 2021
    """
    from sheets import write_cell
    from config.settings import COL_L, KATEGORI_ROW

    def populate_page(ws, col_map: dict, state_metrics: dict, row_offset: int = 0):
        # We ignore the engine's parsed `col_map` and build our own 
        # by working backwards from COL_L (column 12).
        for i, year in enumerate(reversed(target_years)):
            col_idx = COL_L - i
            
            # 1. Write the year into the header row dynamically
            header_row = KATEGORI_ROW + row_offset
            write_cell(ws, header_row, col_idx, str(year), bold=True)
            
            # 2. Fetch the state data for this specific year
            year_data = state_metrics.get(str(year), {})
            
            # 3. Populate the rows
            for template_row, metric_name in row_map.items():
                actual_row = template_row + row_offset
                val = year_data.get(metric_name, "n.a")
                write_cell(ws, actual_row, col_idx, val, decimals=round_decimals, bold=False)

    return populate_page