# data_loader.py (added progress print)
import pandas as pd
from config.settings import DATABASE_PATH

class Database:
    """Singleton class to load and hold the database in memory."""
    def __init__(self):
        print(f"Loading database into memory from:\n  {DATABASE_PATH}")
        self.all_sheets = pd.read_excel(DATABASE_PATH, sheet_name=None)
        
        # 1. Standard Geographic Dimensions
        self.dim_geo = self.all_sheets.get('dim_geografi')
        self.dim_daerah = self.all_sheets.get('dim_daerah')
        
        # 2. Isolated Domain Dimensions (PDRM & JKM)
        self.dim_pdrm = self.all_sheets.get('dim_daerah_pdrm')
        self.dim_jkm = self.all_sheets.get('dim_cawangan_jkm')
        
        # 3. Standard Fact Tables
        self.fact_dun = self.all_sheets.get('fact_metrics_dun')
        self.fact_parlimen = self.all_sheets.get('fact_metrics_parlimen')
        self.fact_negeri = self.all_sheets.get('fact_metrics_negeri')
        self.fact_daerah = self.all_sheets.get('fact_metrics_daerah')
        self.fact_malaysia = self.all_sheets.get('fact_metrics_malaysia')
        
        # 4. NEW: Isolated Domain Fact Tables
        self.fact_pdrm = self.all_sheets.get('fact_metrics_daerah_pdrm')
        self.fact_jkm = self.all_sheets.get('fact_metrics_cawangan_jkm')
        
        print("✅ Database loaded successfully.\n")

# Initialize the database instance to be imported by other modules
db = Database()