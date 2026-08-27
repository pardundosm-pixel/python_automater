import os
import pandas as pd
from config.settings import BASE_DIR

# Point this directly to your new Parquet folder
DATABASE_DIR = os.path.join(BASE_DIR, 'data', '_normalized_db_parquet')

class Database:
    """Singleton class to load and hold the Parquet database in memory."""
    def __init__(self):
        print(f"Loading Parquet database into memory from:\n  {DATABASE_DIR}")
        
        def load_table(table_name):
            file_path = os.path.join(DATABASE_DIR, f"{table_name}.parquet")
            if os.path.exists(file_path):
                return pd.read_parquet(file_path)
            
            print(f"Warning: '{table_name}.parquet' not found.")
            return pd.DataFrame() # Return empty DataFrame to prevent crashes

        # 1. Standard Geographic Dimensions
        self.dim_geo = load_table('dim_geografi')
        self.dim_daerah = load_table('dim_daerah')
        
        # 2. Isolated Domain Dimensions
        self.dim_pdrm = load_table('dim_daerah_pdrm')
        self.dim_jkm = load_table('dim_cawangan_jkm')
        self.dim_meteorologi = load_table('dim_meteorologi')
        
        # 3. Standard Fact Tables
        self.fact_dun = load_table('fact_metrics_dun')
        self.fact_parlimen = load_table('fact_metrics_parlimen')
        self.fact_negeri = load_table('fact_metrics_negeri')
        self.fact_daerah = load_table('fact_metrics_daerah')
        self.fact_malaysia = load_table('fact_metrics_malaysia')
        
        # 4. Isolated Domain Fact Tables
        self.fact_pdrm = load_table('fact_metrics_daerah_pdrm')
        self.fact_jkm = load_table('fact_metrics_cawangan_jkm')
        self.fact_meteorologi = load_table('fact_metrics_meteorologi')
        
        print("Parquet Database loaded successfully.\n")

# Expose a single instance to be imported across the app
db = Database()