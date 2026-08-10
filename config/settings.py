import os

# Define root directory dynamically
BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

# Data Paths
DATABASE_PATH = os.path.join(BASE_DIR, 'data', 'normalized_database_2026_v1.xlsx')

# ==========================================
# MULTI-TEMPLATE CONFIGURATION
# ==========================================
# Map your template profiles to their dynamic file paths
TEMPLATE_PATHS = {
    "parlimen_dun": os.path.join(BASE_DIR, 'data', 'TEMPLATE_PARLIMEN 2026 AS AT 7.8.2026.xlsx'),
    "malaysia": os.path.join(BASE_DIR, 'data', 'TEMPLATE_PARLIMEN 2025 AS AT 22.10.2025_MALAYSIA.xlsx'),
    "negeri": os.path.join(BASE_DIR, 'data', 'TEMPLATE_PARLIMEN 2025 AS AT 22.10.2025_NEGERI.xlsx'),
}

# Output Path
OUTPUT_DIR = os.path.join(BASE_DIR, 'output')