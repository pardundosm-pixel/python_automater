import os

# Define root directory dynamically
BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

# Data Paths
DATABASE_PATH = os.path.join(BASE_DIR, 'data', 'normalized_database_2026_v2.xlsx')

# ==========================================
# MASTER TEMPLATE CONFIGURATION
# ==========================================
# TODO: Update this string to the exact name of your new combined Excel file
MASTER_TEMPLATE_NAME = 'TEMPLATE_PARLIMEN 2026 AS AT 7.8.2026.xlsx' 
MASTER_TEMPLATE_PATH = os.path.join(BASE_DIR, 'data', MASTER_TEMPLATE_NAME)

# ==========================================
# STAGE 2 — PIPELINE STAGING
# ==========================================
# Move intermediate outputs into a 'data/templates_with_locations' folder
OUTPUT_LOCATIONS_DIR = "/Users/akmalhakim/Downloads/_NBDAC/Subnasional/parlimen_dun/templates_with_locations"

# Output Path
OUTPUT_DIR = os.path.join(BASE_DIR, 'output')
# OUTPUT_DIR = r"C:\Users\rubiah\Desktop\Subnasional 2026\parlimen_dun_python_code\python_automater\output"  # Hardcoded for testing