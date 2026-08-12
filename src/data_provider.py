# data_provider.py (added logging prints)
import pandas as pd
from src.data_loader import db

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

    elif level == 'malaysia':
        df = db.fact_malaysia
        mask = _norm(df['lokasi']) == location_code_norm
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
        
        # Extract the district codes and names
        districts_raw = daerah_subset[['kod_daerah', 'nama_daerah']].drop_duplicates().dropna().to_dict('records')
        
        # Filter out empty or "n.a." values
        districts = [{'code': d['kod_daerah'], 'name': d['nama_daerah']} 
                        for d in districts_raw if str(d['kod_daerah']).lower() not in ['n.a.', 'n.a', 'na']]
    else:
        print("[DATA] Warning: 'dim_daerah' sheet not found in the database. Returning empty district list.")
        districts = []
    
    print(f"[DATA] Found {len(districts)} Daerahs for Negeri {state_name}")
    return {
        'state_code': clean_target,
        'state_name': state_name,
        'districts': districts
    }