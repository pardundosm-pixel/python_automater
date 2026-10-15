"""
sheets/jadual_2.py
==================
Mapper for Jadual 2.

Column layout: Years defined in code, injected from the rightmost column backwards.
Row layout (page 1)
-------------------
  8  Jumlah penduduk
 10  Kepadatan penduduk
 12  Luas kawasan
"""
from sheets._base_template import make_populate_year

JADUAL_NUMBER  = 2
ROUND_DECIMALS = 1

# Define the years to extract. 
# The engine will place the last item ("2023") in COL_L, "2022" in COL_K, etc.
TARGET_YEARS = ["2021", "2022", "2023"]

ROW_MAP: dict[int, str] = {
    8:  "jumlah_penduduk",
    10: "kepadatan_penduduk",
    12: "luas_kawasan",
}

# Use the new factory
populate_page = make_populate_year(ROW_MAP, TARGET_YEARS, ROUND_DECIMALS)