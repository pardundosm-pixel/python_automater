"""
src/data_loader.py
==================
Singleton loader.

Cache strategy
--------------
1. If `data/_db_cache/` exists and is NEWER than the source `.xlsx`,
   load Parquet files from the cache (fast, ~1–2 s).
2. Otherwise, read the `.xlsx`, write each sheet to
   `data/_db_cache/<sheet>.parquet`, and use the freshly-read sheets.

To force a rebuild after the source `.xlsx` changes:
    rm -rf data/_db_cache/

The cache is invalidated automatically whenever the `.xlsx` is newer
than the cache folder's mtime.
"""
import os
import pandas as pd

from config.settings import DATABASE_PATH, BASE_DIR


# ==========================================
# CACHE LOCATION
# ==========================================
CACHE_DIR = os.path.join(BASE_DIR, 'data', '_db_cache')


# ==========================================
# CACHE HELPERS
# ==========================================
def _cache_is_fresh() -> bool:
    """True if the cache exists AND is newer than the source .xlsx."""
    if not os.path.isdir(CACHE_DIR):
        return False

    # Any .parquet file in the cache dir?
    if not any(f.endswith(".parquet") for f in os.listdir(CACHE_DIR)):
        return False

    try:
        return os.path.getmtime(CACHE_DIR) > os.path.getmtime(DATABASE_PATH)
    except OSError:
        return False


def _load_from_excel() -> dict:
    """Read every sheet from the source .xlsx and write a Parquet cache."""
    print(f"Reading Excel (first run — building cache):\n  {DATABASE_PATH}")
    sheets = pd.read_excel(DATABASE_PATH, sheet_name=None)

    os.makedirs(CACHE_DIR, exist_ok=True)

    # Clean stale parquet files from a previous cache
    for fname in os.listdir(CACHE_DIR):
        if fname.endswith(".parquet"):
            try:
                os.remove(os.path.join(CACHE_DIR, fname))
            except OSError:
                pass

    written = 0
    for name, df in sheets.items():
        path = os.path.join(CACHE_DIR, f"{name}.parquet")
        try:
            df.to_parquet(path, index=False)
            written += 1
        except Exception as first_err:
            # Fallback: pyarrow can't infer a numeric type because the
            # column mixes numbers with placeholders ("n.a.", "-", …).
            # Stringify every object-dtype column and retry.
            try:
                df2 = df.copy()
                for col in df2.columns:
                    if df2[col].dtype == "object":
                        df2[col] = df2[col].astype("string")   # pandas nullable str
                df2.to_parquet(path, index=False)
                written += 1
                print(f"  [info] Cached '{name}' using string-fallback "
                      f"(original error: {first_err})")
            except Exception as second_err:
                print(f"  [warn] Could not cache sheet '{name}': {second_err}")

    # Touch the cache dir so its mtime reflects the conversion time
    os.utime(CACHE_DIR, None)

    print(f"✅ Cached {written} sheet(s) → {CACHE_DIR}\n")
    return sheets


def _load_from_cache() -> dict:
    """Read every .parquet file in the cache dir into a dict of DataFrames."""
    print(f"Loading cached database from:\n  {CACHE_DIR}")
    sheets = {}
    for fname in os.listdir(CACHE_DIR):
        if fname.endswith(".parquet"):
            name = fname[:-8]  # strip ".parquet"
            sheets[name] = pd.read_parquet(os.path.join(CACHE_DIR, fname))
    print(f"✅ Loaded {len(sheets)} sheet(s) from cache\n")
    return sheets


# ==========================================
# DATABASE SINGLETON
# ==========================================
class Database:
    """Singleton class to load and hold the database in memory."""

    def __init__(self):
        self.all_sheets = (
            _load_from_cache() if _cache_is_fresh()
            else _load_from_excel()
        )

        # --------------------------------------------------
        # 1. Standard Geographic Dimensions
        # --------------------------------------------------
        self.dim_geo    = self.all_sheets.get('dim_geografi')
        self.dim_daerah = self.all_sheets.get('dim_daerah')

        # --------------------------------------------------
        # 2. Isolated Domain Dimensions
        # --------------------------------------------------
        self.dim_pdrm        = self.all_sheets.get('dim_daerah_pdrm')
        self.dim_jkm         = self.all_sheets.get('dim_cawangan_jkm')
        self.dim_meteorologi = self.all_sheets.get('dim_meteorologi')

        # --------------------------------------------------
        # 3. Standard Fact Tables
        # --------------------------------------------------
        self.fact_dun      = self.all_sheets.get('fact_metrics_dun')
        self.fact_parlimen = self.all_sheets.get('fact_metrics_parlimen')
        self.fact_negeri   = self.all_sheets.get('fact_metrics_negeri')
        self.fact_daerah   = self.all_sheets.get('fact_metrics_daerah')
        self.fact_malaysia = self.all_sheets.get('fact_metrics_malaysia')

        # --------------------------------------------------
        # 4. Isolated Domain Fact Tables
        # --------------------------------------------------
        self.fact_pdrm        = self.all_sheets.get('fact_metrics_daerah_pdrm')
        self.fact_jkm         = self.all_sheets.get('fact_metrics_cawangan_jkm')
        self.fact_meteorologi = self.all_sheets.get('fact_metrics_meteorologi')


# Singleton instance — imported by data_provider.py and run_factory.py
db = Database()