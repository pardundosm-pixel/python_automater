# data_provider.py (with query caching)
import pandas as pd
from functools import lru_cache
from src.data_loader import db

@lru_cache(maxsize=128)
def get_location_hierarchy(parlimen_code: str):
    """Fetches Parliament details and its nested DUNs from dim_geografi."""
    print(f"[DATA] Fetching location hierarchy for parl_code: {parlimen_code}")
    geo = db.dim_geo
    
    subset = geo[geo['kod_parlimen'] == parlimen_code]
    
    if subset.empty:
        print(f"[DATA] No location found for parl_code: {parlimen_code}")
        return None
        
    parl_name = subset.iloc[0]['nama_parlimen']
    state_code = subset.iloc[0]['kod_negeri']
    state_name = subset.iloc[0]['nama_negeri']
    
    duns_raw = subset[['kod_dun', 'nama_dun']].drop_duplicates().to_dict('records')
    
    duns = [{'code': d['kod_dun'], 'name': d['nama_dun']} 
            for d in duns_raw if str(d['kod_dun']).lower() != 'n.a.']
    
    print(f"[DATA] Found {len(duns)} DUNs for {parl_name} ({state_name})")
    return {
        'parl_code': parlimen_code,
        'parl_name': parl_name,
        'state_code': state_code,
        'state_name': state_name,
        'duns': duns
    }

@lru_cache(maxsize=256)
def get_metrics_dict(location_code: str, level: str, parent_code: str = None):
    """
    Fetches metrics for a location and formats them into a nested dictionary by year.
    If level is 'dun', parent_code (kod_parlimen) MUST be provided to prevent state collisions.
    """
    print(f"[DATA] Fetching metrics for location_code: {location_code}, level: {level}, parent_code: {parent_code}")

    def _norm(series):
        # Normalize to string regardless of whether the source column was
        # loaded as str, int, or float (e.g. 1 / 1.0 / "01").
        s = series.astype(str).str.strip()
        return s.str.replace(r'\.0$', '', regex=True)

    def _norm_negeri_code(value):
        # State codes are always 2-digit, zero-padded ("01".."16").
        v = str(value).strip()
        if v.endswith('.0'):
            v = v[:-2]
        return v.zfill(2)

    def _norm_daerah_code(value):
        # State and District codes are usually 2-digit, zero-padded ("01", "02").
        v = str(value).strip()
        if v.endswith('.0'):
            v = v[:-2]
        return v.zfill(2)

    location_code_norm = str(location_code).strip()
    if location_code_norm.endswith('.0'):
        location_code_norm = location_code_norm[:-2]

    if level == 'parlimen':
        df = db.fact_parlimen
        mask = _norm(df['kod_parlimen']) == location_code_norm

    elif level == 'dun':
        df = db.fact_dun
        mask = _norm(df['kod_dun']) == location_code_norm
        if parent_code:
            mask = mask & (_norm(df['kod_parlimen']) == str(parent_code).strip())

    elif level == 'negeri':
        df = db.fact_negeri
        target = _norm_negeri_code(location_code)
        mask = df['kod_negeri'].apply(_norm_negeri_code) == target

    elif level == 'daerah':
        df = db.fact_daerah
        
        # --- NEW: Handle State Total inside fact_daerah (empty kod_daerah) ---
        if location_code == "STATE_TOTAL":
            # Catch all pandas variations of an empty/null cell
            mask = df['kod_daerah'].isna() | \
                    (df['kod_daerah'].astype(str).str.strip() == '') | \
                    (df['kod_daerah'].astype(str).str.lower() == 'nan') | \
                    (df['kod_daerah'].astype(str).str.lower() == 'n.a.')
        else:
            target = _norm_daerah_code(location_code)
            mask = df['kod_daerah'].apply(_norm_daerah_code) == target
            
        if parent_code:
            parent_target = _norm_negeri_code(parent_code)
            mask = mask & (df['kod_negeri'].apply(_norm_negeri_code) == parent_target)

    elif level == 'malaysia':
        df = db.fact_malaysia
        mask = _norm(df['lokasi']) == location_code_norm

    elif level == 'pdrm':
        df = db.fact_pdrm
        target = _norm_daerah_code(location_code)
        
        # Safely handle column naming variations in the database
        if 'kod_daerah_pdrm' in df.columns:
            mask = df['kod_daerah_pdrm'].apply(_norm_daerah_code) == target
        elif 'kod_daerah' in df.columns:
            mask = df['kod_daerah'].apply(_norm_daerah_code) == target
        else:
            print("[DATA] Error: Could not find PDRM district code column in fact_pdrm")
            return {}
            
        # Add safety lock for State Code to prevent collisions
        if parent_code:
            parent_target = _norm_negeri_code(parent_code)
            mask = mask & (df['kod_negeri'].apply(_norm_negeri_code) == parent_target)
            
    elif level == 'jkm':
        df = db.fact_jkm
        
        # --- NEW: Handle State Total inside fact_jkm (empty kod_cawangan_jkm) ---
        if location_code == "STATE_TOTAL":
            # Catch all pandas variations of an empty/null cell
            mask = df['kod_cawangan_jkm'].isna() | \
                (df['kod_cawangan_jkm'].astype(str).str.strip() == '') | \
                (df['kod_cawangan_jkm'].astype(str).str.lower() == 'nan') | \
                (df['kod_cawangan_jkm'].astype(str).str.lower() == 'n.a.')
        else:
            # Use _norm_daerah_code to guarantee "2" becomes "02"
            target = _norm_daerah_code(location_code)
            mask = df['kod_cawangan_jkm'].apply(_norm_daerah_code) == target
            
        # Add safety lock for State Code to prevent collisions
        if parent_code:
            parent_target = _norm_negeri_code(parent_code)
            mask = mask & (df['kod_negeri'].apply(_norm_negeri_code) == parent_target)

    elif level == 'meteorologi':
        dim_df = db.dim_meteorologi
        fact_df = db.fact_meteorologi
        
        # 1. Search for the station by its name (stesen) case-insensitively
        target_name = str(location_code).strip().lower()
        dim_mask = dim_df['stesen'].astype(str).str.strip().str.lower() == target_name
        
        # 2. Lock the search to the specific state to prevent cross-state collisions
        if parent_code:
            target_state = _norm_negeri_code(parent_code)
            dim_mask = dim_mask & (dim_df['kod_negeri'].apply(_norm_negeri_code) == target_state)
            
        dim_subset = dim_df[dim_mask]
        
        if dim_subset.empty:
            print(f"[DATA] Error: Station '{location_code}' not found in dim_meteorologi for state '{parent_code}'.")
            return {}
            
        # 3. Extract the exact kod_stesen and kod_negeri from the dimension table
        # FIX: Manually normalize the single scalar string instead of using the Pandas _norm() function
        raw_stesen = str(dim_subset.iloc[0]['kod_stesen']).strip()
        target_kod_stesen = raw_stesen[:-2] if raw_stesen.endswith('.0') else raw_stesen
        
        target_kod_negeri = _norm_negeri_code(dim_subset.iloc[0]['kod_negeri'])
        
        # 4. Build the final mask for the fact table using BOTH codes
        # We can still use _norm() here because fact_df['kod_stesen'] is a full Pandas Series
        mask = (_norm(fact_df['kod_stesen']) == target_kod_stesen) & \
                (fact_df['kod_negeri'].apply(_norm_negeri_code) == target_kod_negeri)
            
        df = fact_df  # Assign back to df so the rest of the function can process it
    
    else:
        print(f"[DATA] Unknown level '{level}' - returning empty")
        return {}

    subset = df[mask]
    if subset.empty:
        print(f"[DATA] No metrics found for {location_code} at level {level}")
        return {}

    metrics_by_year = {}
    for tahun, group in subset.groupby('tahun'):
        tahun_str = str(tahun).strip() if pd.notnull(tahun) else 'unknown'
        if tahun_str.endswith(".0"):
            tahun_str = tahun_str[:-2]

        metrics_by_year[tahun_str] = dict(zip(group['kategori_metrik'], group['nilai']))

    print(f"[DATA] Retrieved metrics for {len(metrics_by_year)} year(s) for {location_code}")
    return metrics_by_year

@lru_cache(maxsize=128)
def get_dun_hierarchy(dun_code: str, parent_parl_code: str = None):
    """Fetches DUN details and its parent Parliament safely using strict matching."""
    print(f"[DATA] Fetching DUN hierarchy for dun_code: {dun_code}, parent_parl_code: {parent_parl_code}")
    geo = db.dim_geo
    
    mask = geo['kod_dun'] == dun_code
    
    # Strictly lock the query to the parent parliament if provided
    if parent_parl_code:
        mask = mask & (geo['kod_parlimen'] == parent_parl_code)
        
    subset = geo[mask]
    
    if subset.empty:
        print(f"[DATA] No DUN found for dun_code: {dun_code} with parent {parent_parl_code}")
        return None
    
    result = {
        'dun_code': dun_code,
        'dun_name': subset.iloc[0]['nama_dun'],
        'state_code': subset.iloc[0]['kod_negeri'],
        'state_name': subset.iloc[0]['nama_negeri'],
        'parent_parl_code': subset.iloc[0]['kod_parlimen'],
        'parent_parl_name': subset.iloc[0]['nama_parlimen']
    }
    print(f"[DATA] Found DUN: {result['dun_name']} under {result['parent_parl_name']}")
    return result

@lru_cache(maxsize=64)
def get_negeri_hierarchy(state_code: str):
    """Fetches State details and its nested Districts from the dim_daerah sheet."""
    print(f"[DATA] Fetching location hierarchy for state_code: {state_code}")
    
    clean_target = str(state_code).zfill(2)
    
    # 1. Fetch State Name from dim_geo (Standard practice for consistency)
    geo = db.dim_geo
    clean_geo_codes = geo['kod_negeri'].astype(str).str.replace(r'\.0$', '', regex=True).str.zfill(2)
    geo_subset = geo[clean_geo_codes == clean_target]
    
    if geo_subset.empty:
        print(f"[DATA] No location found for state_code: {state_code} in dim_geografi")
        return None
        
    state_name = geo_subset.iloc[0]['nama_negeri']
    
    # 2. Extract Districts safely from the dedicated dim_daerah sheet
    dim_daerah = db.dim_daerah
    if dim_daerah is not None and not dim_daerah.empty:
        
        # Normalize the state codes in dim_daerah just in case
        daerah_mask = dim_daerah['kod_negeri'].astype(str).str.replace(r'\.0$', '', regex=True).str.zfill(2) == clean_target
        daerah_subset = dim_daerah[daerah_mask]
        
        # Extract the district codes and names using the new official district columns
        districts_raw = daerah_subset[['kod_daerah_pentadbiran', 'daerah_pentadbiran']].drop_duplicates().dropna().to_dict('records')
        
        # Filter out empty or "n.a." values
        districts = [{'code': d['kod_daerah_pentadbiran'], 'name': d['daerah_pentadbiran']} 
                        for d in districts_raw if str(d['kod_daerah_pentadbiran']).lower() not in ['n.a.', 'n.a', 'na']]
    else:
        print("[DATA] Warning: 'dim_daerah' sheet not found in the database. Returning empty district list.")
        districts = []
    
    print(f"[DATA] Found {len(districts)} Daerahs for Negeri {state_name}")
    return {
        'state_code': clean_target,
        'state_name': state_name,
        'districts': districts
    }

# The following functions are not cached in the provided example, but you may add caching if needed.
# For consistency with the example, we leave them as is.
def get_pdrm_hierarchy(state_code: str):
    """Fetches unique Police Districts (Daerah PDRM) for a given state from the isolated dim_daerah_pdrm table."""
    print(f"[DATA] Fetching PDRM hierarchy for state_code: {state_code}")
    dim_pdrm = db.dim_pdrm
    
    if dim_pdrm is None or dim_pdrm.empty:
        print("[DATA] Warning: 'dim_daerah_pdrm' sheet not found.")
        return []
        
    clean_target = str(state_code).zfill(2)
    pdrm_mask = dim_pdrm['kod_negeri'].astype(str).str.replace(r'\.0$', '', regex=True).str.zfill(2) == clean_target
    pdrm_subset = dim_pdrm[pdrm_mask]
    
    # Extract unique PDRM districts
    pdrm_raw = pdrm_subset[['kod_daerah_pdrm', 'daerah_pdrm']].drop_duplicates().dropna().to_dict('records')
    
    pdrm_districts = [{'code': d['kod_daerah_pdrm'], 'name': d['daerah_pdrm']} 
                        for d in pdrm_raw if str(d['kod_daerah_pdrm']).lower() not in ['n.a.', 'n.a', 'na']]

    print(f"[DATA] Found {len(pdrm_districts)} PDRM Districts.")
    return pdrm_districts

def get_jkm_hierarchy(state_code: str):
    """Fetches unique Social Welfare Branches (Cawangan JKM) for a given state from the isolated dim_cawangan_jkm table."""
    print(f"[DATA] Fetching JKM hierarchy for state_code: {state_code}")
    dim_jkm = db.dim_jkm
    
    if dim_jkm is None or dim_jkm.empty:
        print("[DATA] Warning: 'dim_cawangan_jkm' sheet not found.")
        return []
        
    clean_target = str(state_code).zfill(2)
    jkm_mask = dim_jkm['kod_negeri'].astype(str).str.replace(r'\.0$', '', regex=True).str.zfill(2) == clean_target
    jkm_subset = dim_jkm[jkm_mask]
    
    # Extract unique JKM branches
    jkm_raw = jkm_subset[['kod_cawangan_jkm', 'cawangan_jkm']].drop_duplicates().dropna().to_dict('records')
    
    jkm_branches = [{'code': d['kod_cawangan_jkm'], 'name': d['cawangan_jkm']} 
                    for d in jkm_raw if str(d['kod_cawangan_jkm']).lower() not in ['n.a.', 'n.a', 'na']]
                    
    print(f"[DATA] Found {len(jkm_branches)} JKM Branches.")
    return jkm_branches

def get_meteorologi_hierarchy(state_code: str):
    """Fetches unique Meteorological Stations for a given state from the isolated dim_meteorologi table."""
    print(f"[DATA] Fetching Meteorologi hierarchy for state_code: {state_code}")
    dim_met = db.dim_meteorologi
    
    if dim_met is None or dim_met.empty:
        print("[DATA] Warning: 'dim_meteorologi' sheet not found.")
        return []
        
    clean_target = str(state_code).zfill(2)
    met_mask = dim_met['kod_negeri'].astype(str).str.replace(r'\.0$', '', regex=True).str.zfill(2) == clean_target
    met_subset = dim_met[met_mask]
    
    # Extract unique stations along with their administrative district
    met_raw = met_subset[['kod_stesen', 'stesen', 'kod_daerah_pentadbiran', 'daerah_pentadbiran']].drop_duplicates().dropna(subset=['stesen']).to_dict('records')
    
    met_stations = [
        {
            'code': d['kod_stesen'], 
            'name': d['stesen'],
            'daerah_code': d.get('kod_daerah_pentadbiran', 'n.a.'),
            'daerah_name': d.get('daerah_pentadbiran', 'n.a.')
        } 
        for d in met_raw if str(d['kod_stesen']).lower() not in ['n.a.', 'n.a', 'na']
    ]

    print(f"[DATA] Found {len(met_stations)} Meteorologi Stations.")
    return met_stations