import pandas as pd
from config.settings import DATABASE_PATH

class Database:
    """Singleton class to load and hold the database in memory."""
    def __init__(self):
        print(f"Loading database into memory from:\n  {DATABASE_PATH}")
        self.all_sheets = pd.read_excel(DATABASE_PATH, sheet_name=None)
        
        # Assign tables to easily accessible attributes
        self.dim_geo = self.all_sheets.get('dim_geografi')
        self.fact_dun = self.all_sheets.get('fact_metrics_dun')
        self.fact_parlimen = self.all_sheets.get('fact_metrics_parlimen')
        self.fact_negeri = self.all_sheets.get('fact_metrics_negeri')
        self.fact_malaysia = self.all_sheets.get('fact_metrics_malaysia')
        
        print("✅ Database loaded successfully.\n")

# Initialize the database instance to be imported by other modules
db = Database()