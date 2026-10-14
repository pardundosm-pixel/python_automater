"""
sheets/jadual_2.py
==================
Mapper for Jadual 2 — "<title>".

Column layout: years across the top (state-level time series).
Row layout (page 1)
-------------------
  8  Jumlah penduduk          ← metric only, no year in tuple
 10  Kepadatan penduduk
 12  Luas kawasan
"""
from sheets._base_template import make_populate_year


JADUAL_NUMBER  = 2
ROUND_DECIMALS = 1


ROW_MAP: dict[int, str] = {
    8:  "jumlah_penduduk",
    10: "kepadatan_penduduk",
    12: "luas_kawasan",
}


populate_page = make_populate_year(ROW_MAP, ROUND_DECIMALS)