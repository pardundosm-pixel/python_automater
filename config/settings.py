"""
config/settings.py
==================
Central configuration for the Subnasional report pipeline.

Stage 1 (location_injection_batch.py):
    TEMPLATES_DIR           →  master jadual templates (pure, no data)
    OUTPUT_LOCATIONS_DIR    →  per-state copies with locations injected

Stage 2 (run_factory.py):
    OUTPUT_LOCATIONS_DIR    →  input
    OUTPUT_FINAL_DIR        →  per-state copies with locations + data
"""
import os

# ==========================================
# ROOT
# ==========================================
BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


# ==========================================
# STAGE 1 — TEMPLATE SOURCES
# ==========================================
# Move your templates into a 'templates' folder at the root of your repo
TEMPLATES_DIR = os.path.join(BASE_DIR, 'templates')

# ==========================================
# STAGE 2 — PIPELINE STAGING
# ==========================================
# Move intermediate outputs into a 'data/templates_with_locations' folder
OUTPUT_LOCATIONS_DIR = "/Users/akmalhakim/Downloads/_NBDAC/Subnasional/parlimen_dun/templates_with_locations"

# Stage 2 output: {State}/Jadual N.xlsx  (locations + data filled)
OUTPUT_FINAL_DIR = os.path.join(BASE_DIR, 'output_final')


# ==========================================
# SHARED LAYOUT CONSTANTS
# (must match location_injection_batch.py)
# ==========================================
ROWS_PER_PAGE = 50   # vertical stride between stacked page blocks
KATEGORI_ROW  = 6    # row where the location header lives (within page 1)
COL_B         = 2    # label column ("MAKLUMAT ASAS")
COL_C         = 3    # year column (when present)
COL_D         = 4    # first location column when a year column exists
COL_L         = 12   # rightmost location column


# ==========================================
# LEGACY — only imported by pdf_compiler.py (deferred)
# TODO: remove together with the pdf_compiler.py rewrite
# ==========================================
OUTPUT_DIR = os.path.join(BASE_DIR, 'output')