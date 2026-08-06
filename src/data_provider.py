import pandas as pd
from src.data_loader import db

def get_location_hierarchy(parlimen_code: str):
    """Fetches Parliament details and its nested DUNs from dim_geografi."""
    geo = db.dim_geo
    
    subset = geo[geo['kod_parlimen'] == parlimen_code]
    
    if subset.empty:
        return None
        
    parl_name = subset.iloc[0]['nama_parlimen']
    state_code = subset.iloc[0]['kod_negeri']
    state_name = subset.iloc[0]['nama_negeri']
    
    duns_raw = subset[['kod_dun', 'nama_dun']].drop_duplicates().to_dict('records')
    
    duns = [{'code': d['kod_dun'], 'name': d['nama_dun']} 
            for d in duns_raw if str(d['kod_dun']).lower() != 'n.a.']
            
    return {
        'parl_code': parlimen_code,
        'parl_name': parl_name,
        'state_code': state_code,
        'state_name': state_name,
        'duns': duns
    }

def get_dun_hierarchy(dun_code: str, parent_parl_code: str = None):
    """Fetches DUN details and its parent Parliament safely using strict matching."""
    geo = db.dim_geo
    
    mask = geo['kod_dun'] == dun_code
    
    # Strictly lock the query to the parent parliament if provided
    if parent_parl_code:
        mask = mask & (geo['kod_parlimen'] == parent_parl_code)
        
    subset = geo[mask]
    
    if subset.empty:
        return None
        
    return {
        'dun_code': dun_code,
        'dun_name': subset.iloc[0]['nama_dun'],
        'state_code': subset.iloc[0]['kod_negeri'],
        'state_name': subset.iloc[0]['nama_negeri'],
        'parent_parl_code': subset.iloc[0]['kod_parlimen'],
        'parent_parl_name': subset.iloc[0]['nama_parlimen']
    }

def get_metrics_dict(location_code: str, level: str, parent_code: str = None):
    """
    Fetches metrics for a location and formats them into a nested dictionary by year.
    If level is 'dun', parent_code (kod_parlimen) MUST be provided to prevent state collisions.
    """
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
        return {}
        
    subset = df[mask]
    if subset.empty:
        return {}
        
    metrics_by_year = {}
    for tahun, group in subset.groupby('tahun'):
        tahun_str = str(tahun).strip() if pd.notnull(tahun) else 'unknown'
        if tahun_str.endswith(".0"):
            tahun_str = tahun_str[:-2]
        
        metrics_by_year[tahun_str] = dict(zip(group['kategori_metrik'], group['nilai']))
        
    return metrics_by_year