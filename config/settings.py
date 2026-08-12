import os

# Define root directory dynamically
BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

# Data Paths
DATABASE_PATH = os.path.join(BASE_DIR, 'data', 'normalized_database_2026_v1.xlsx')

# ==========================================
# MASTER TEMPLATE CONFIGURATION
# ==========================================
# TODO: Update this string to the exact name of your new combined Excel file
MASTER_TEMPLATE_NAME = 'TEMPLATE_PARLIMEN 2026 AS AT 7.8.2026.xlsx' 
MASTER_TEMPLATE_PATH = os.path.join(BASE_DIR, 'data', MASTER_TEMPLATE_NAME)

# We keep the dictionary structure so engine.py routing doesn't break, 
# but we point all profiles to the exact same master file.
TEMPLATE_PATHS = {
    "parlimen_dun": MASTER_TEMPLATE_PATH,
    "malaysia": MASTER_TEMPLATE_PATH,
    "negeri": MASTER_TEMPLATE_PATH,
}

# Output Path
OUTPUT_DIR = os.path.join(BASE_DIR, 'output')
# OUTPUT_DIR = r"C:\Users\rubiah\Desktop\Subnasional 2026\parlimen_dun_python_code\python_automater\output"  # Hardcoded for testing