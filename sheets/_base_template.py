"""
sheets/_base_template.py
========================
Shared mapper factories. Each jadual_N.py calls ONE of:

    populate_column = make_populate_location(ROW_MAP, ROUND_DECIMALS)
    populate_page   = make_populate_year(ROW_MAP, ROUND_DECIMALS)

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
# MODE B — COLUMNS ARE YEARS
# ==========================================
def make_populate_year(row_map: dict, round_decimals: int | None):
    """
    Build a `populate_page` for jaduals whose columns are years (state-level
    time series). The whole file is about one state.

    ROW_MAP = {row: metric_name}
    Called ONCE per page, with:
        col_map       = {col_idx: year_str}       (parsed from the header row)
        state_metrics = {year_str: {metric: val}} (state's metrics)
    """
    def populate_page(ws, col_map: dict, state_metrics: dict,
                      row_offset: int = 0):
        for template_row, metric_name in row_map.items():
            actual_row = template_row + row_offset
            for col_idx, year in col_map.items():
                year_data = state_metrics.get(str(year), {})
                val       = year_data.get(metric_name, "n.a")
                write_cell(ws, actual_row, col_idx, val,
                           decimals=round_decimals, bold=False)
    return populate_page