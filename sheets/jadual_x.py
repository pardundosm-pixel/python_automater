"""
sheets/jadual_N.py
==================
Mapper for Jadual N — "TODO: short title of jadual".

Row layout (page 1)
-------------------
TODO: fill in the annotated row layout, same style as jadual_1.py.
This block is documentation only — the actual mapping lives in ROW_MAP.

Example:
   8  <metric A>      2023
   9  <metric A>      2024
  10  <metric A>      2025
  ...

Note on metric keys
-------------------
The tuple's first element must match a value of `kategori_metrik` in the
fact tables. Verify with:
    >>> from src.data_loader import db
    >>> sorted(db.fact_parlimen['kategori_metrik'].unique())
"""
from sheets import write_cell


# ==========================================
# JADUAL IDENTITY
# ==========================================
JADUAL_NUMBER = None            # TODO: replace with the jadual number, e.g. 2

# Decimal places for this jadual.
#   1   →  round to 1 dp, whole numbers stay as int
#   2   →  round to 2 dp
#   None →  no rounding, just integerize whole numbers
ROUND_DECIMALS = None           # TODO: set the per-jadual rounding


# ==========================================
# ROW MAP  (page-1 rows only)
# row → (metric_name, year_key)
# ==========================================
ROW_MAP: dict[int, tuple[str, str]] = {
    # TODO: fill in every data row of the template.
    #       Keys are the row numbers WITHIN PAGE 1 (rows 1..50).
    #       The engine adds `row_offset` automatically for pages 2, 3, …
    #
    # Example shape (delete once filled in):
    #   8:  ("<metric_string_from_db>", "2023"),
    #   9:  ("<metric_string_from_db>", "2024"),
    #   10: ("<metric_string_from_db>", "2025"),
}


# ==========================================
# COLUMN WRITER  (unchanged — copy as-is)
# ==========================================
def populate_column(ws, col_idx: int, metrics_by_year: dict, row_offset: int = 0):
    """
    Write one location's values into `col_idx`.

    Parameters
    ----------
    ws               : openpyxl worksheet
    col_idx          : absolute column index to write into
    metrics_by_year  : {year_str: {metric_name: value}}
    row_offset       : 0 for page 1, +50 for page 2, +100 for page 3, …
    """
    for template_row, (metric_name, year) in ROW_MAP.items():
        actual_row = template_row + row_offset
        year_data  = metrics_by_year.get(str(year), {})
        val        = year_data.get(metric_name, "n.a")
        write_cell(ws, actual_row, col_idx, val, decimals=ROUND_DECIMALS)