"""
sheets/jadual_1.py
==================
Mapper for Jadual 1 — "Statistik maklumat asas".

The engine calls `populate_column(ws, col_idx, metrics_by_year, row_offset)`
once per location column. This module only knows the ROW layout of the
jadual (page-1 rows); the engine adds `row_offset` for stacked pages.

Row layout (page 1)
-------------------
    8  Luas kawasan (km²)          2023
    9  Luas kawasan (km²)          2024
   10  Luas kawasan (km²)          2025
   ...
   33  (same)                                                PRU-15

Metric strings must match a value of `kategori_metrik` in the fact tables.
Verify with:
    >>> from src.data_loader import db
    >>> sorted(db.fact_parlimen['kategori_metrik'].unique())
"""
from sheets._base_template import make_populate_location

JADUAL_NUMBER  = "1"        # string — matches filenames like "Jadual 1.xlsx"
ROUND_DECIMALS = 1


# ==========================================
# ROW MAP  (page-1 rows only)
# row → (metric_name, year_key)
# ==========================================
ROW_MAP: dict[int, tuple[str, str]] = {
    # Luas kawasan
    8:  ("luas_kawasan", "2023"),
    9:  ("luas_kawasan", "2024"),
    10: ("luas_kawasan", "2025"),

    # Jumlah penduduk
    12: ("jumlah_penduduk", "2023"),
    13: ("jumlah_penduduk", "2024"),
    14: ("jumlah_penduduk", "2025"),

    # Kepadatan penduduk
    16: ("kepadatan_penduduk", "2023"),
    17: ("kepadatan_penduduk", "2024"),
    18: ("kepadatan_penduduk", "2025"),

    # Jumlah pemilih
    20: ("jumlah_pemilih", "PRU-14"),
    21: ("jumlah_pemilih", "PRU-15"),

    # Jumlah undian oleh pemilih
    23: ("jumlah_undian_oleh_pemilih", "PRU-14"),
    24: ("jumlah_undian_oleh_pemilih", "PRU-15"),

    # Jumlah kertas undi yang ditolak
    26: ("jumlah_kertas_undi_yang_ditolak", "PRU-14"),
    27: ("jumlah_kertas_undi_yang_ditolak", "PRU-15"),

    # Kertas undi dikeluarkan tetapi tidak dimasukkan / tidak dikembalikan
    29: ("kertas_undi_dikeluarkan_tetapi_tidak_dimasukkan_dan_tidak_dikembalikan", "PRU-14"),
    30: ("kertas_undi_dikeluarkan_tetapi_tidak_dimasukkan_dan_tidak_dikembalikan", "PRU-15"),

    # Jumlah kertas undi yang dikeluarkan dalam pengundian
    32: ("jumlah_kertas_undi_yang_dikeluarkan_dalam_pengundian", "PRU-14"),
    33: ("jumlah_kertas_undi_yang_dikeluarkan_dalam_pengundian", "PRU-15"),
}


# ==========================================
# COLUMN WRITER
# ==========================================
populate_column = make_populate_location(ROW_MAP, ROUND_DECIMALS)