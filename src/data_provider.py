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
    if level == 'parlimen':
        df = db.fact_parlimen
        mask = df['kod_parlimen'] == location_code
        
    elif level == 'dun':
        df = db.fact_dun
        mask = df['kod_dun'] == location_code
        if parent_code:
            mask = mask & (df['kod_parlimen'] == parent_code)
            
    elif level == 'negeri':
        df = db.fact_negeri
        mask = df['kod_negeri'] == location_code
        
    elif level == 'malaysia':
        df = db.fact_malaysia
        mask = df['lokasi'] == location_code
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