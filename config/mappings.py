"""
This file maps Excel Row indices to the exact 'kategori_metrik' 
names found in the normalized database.
"""

# ==========================================
# JADUAL 1.0 MAPPING
# Tuple format: (Metric Name, Year)
# ==========================================
JADUAL_1_ROW_MAP = {
    8:  ("luas_kawasan", "2022"),
    9:  ("luas_kawasan", "2023"),
    10: ("luas_kawasan", "2024"),
    12: ("jumlah_penduduk", "2022"),
    13: ("jumlah_penduduk", "2023"),
    14: ("jumlah_penduduk", "2024"),
    16: ("kepadatan_penduduk", "2022"),
    17: ("kepadatan_penduduk", "2023"),
    18: ("kepadatan_penduduk", "2024"),
    
    # Assuming elections are now stored simply under the specific election year/code
    20: ("jumlah_pemilih", "PRU14"), 
    21: ("jumlah_pemilih", "PRU15"), 
    23: ("jumlah_undian_oleh_pemilih", "PRU14"), 
    24: ("jumlah_undian_oleh_pemilih", "PRU15")
}

# ==========================================
# JADUAL 2.1 & 2.2 MAPPING (Demographics)
# ==========================================
JADUAL_2_ROW_MAP = {
    8:  "jumlah_penduduk",
    12: "warganegara",
    13: "bukan_warganegara",
    15: "lelaki",
    16: "perempuan",
    19: "peratus_warganegara",
    20: "peratus_bukan_warganegara",
    24: "peratus_bumiputera",
    25: "peratus_cina",
    26: "peratus_india",
    27: "peratus_lain_lain",
    31: "umur_0_14",
    33: "umur_15_64",
    35: "umur_65_lebih",
    37: "umur_18_lebih",
    42: "jumlah_nisbah_tanggungan",
    43: "umur_muda",
    44: "umur_tua",
    46: "nisbah_jantina",
    49: "kepadatan_penduduk"
}

# Add X-Axis year column mapping for Jadual 2
JADUAL_2_COL_MAP = {
    6: "2023",  # Column F in Excel
    7: "2024"   # Column G in Excel
}