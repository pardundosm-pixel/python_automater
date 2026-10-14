import os

# ==========================================
# ROOT
# ==========================================
BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

# ==========================================
# TEMPLATE SOURCES
# ==========================================
# Master jadual templates (pure: no locations, no data).
# One .xlsx per jadual. Consumed by Stage 1 (location_injection_batch.py).
TEMPLATES_DIR = "/Users/akmalhakim/Downloads/_NBDAC/Subnasional/template_generator/1_template_automation/output_v2"


# ==========================================
# PIPELINE STAGING
# ==========================================
# Stage 1 output: {state}/Jadual N.xlsx   (locations injected, data empty)
OUTPUT_LOCATIONS_DIR = "/Users/akmalhakim/Downloads/_NBDAC/Subnasional/template_generator/2_location_injector/output_with_locations"

# Stage 2 output: {state}/Jadual N.xlsx   (locations + data)
OUTPUT_FINAL_DIR = os.path.join(BASE_DIR, 'output_final')


# ==========================================
# LEGACY  (kept only so the deferred pdf_compiler.py still imports)
# TODO: remove once pdf_compiler.py is rewritten.
# ==========================================
OUTPUT_DIR = os.path.join(BASE_DIR, 'output')