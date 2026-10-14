"""
Jadual 1 : Statistik maklumat asas
One file per jadual. Fills a single column (one location) in-place.
"""

JADUAL_NUMBER = 1

# row number  →  (metric name as it appears in `metrics_by_year`, year)
# Only data rows are listed — header/notes rows are untouched.
ROW_MAP = {
    # Luas kawasan (km2)
    8:  ("Luas kawasan",    "2023"),
    9:  ("Luas kawasan",    "2024"),
    10: ("Luas kawasan",    "2025"),
    # Jumlah penduduk ('000)
    12: ("Jumlah penduduk", "2023"),
    13: ("Jumlah penduduk", "2024"),
    14: ("Jumlah penduduk", "2025"),
    # Kepadatan penduduk (per km2)
    16: ("Kepadatan penduduk", "2023"),
    17: ("Kepadatan penduduk", "2024"),
    18: ("Kepadatan penduduk", "2025"),
    # Jumlah pemilih (PRU-14 / PRU-15)
    20: ("Jumlah pemilih",       "PRU-14"),
    21: ("Jumlah pemilih",       "PRU-15"),
    # Jumlah undian oleh pemilih
    23: ("Jumlah undian",        "PRU-14"),
    24: ("Jumlah undian",        "PRU-15"),
    # Jumlah kertas undi yang ditolak
    26: ("Kertas undi ditolak",  "PRU-14"),
    27: ("Kertas undi ditolak",  "PRU-15"),
    # Undi dikeluarkan tidak masuk peti & tidak dikembalikan
    29: ("Undi tidak dikembalikan", "PRU-14"),
    30: ("Undi tidak dikembalikan", "PRU-15"),
    # Jumlah kertas undi yang dikeluarkan dalam pengundian
    32: ("Kertas undi dikeluarkan", "PRU-14"),
    33: ("Kertas undi dikeluarkan", "PRU-15"),
}


def populate_column(ws, col_idx, metrics_by_year):
    """Write one location's values into column `col_idx`."""
    for row, (metric_name, year) in ROW_MAP.items():
        year_data = metrics_by_year.get(str(year), {})
        val = year_data.get(metric_name, "n.a")

        if val in (None, "", "n.a"):
            val = "n.a"
        else:
            try:
                val = float(val)
            except (ValueError, TypeError):
                pass

        ws.cell(row=row, column=col_idx).value = val