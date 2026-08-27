import os
import time
from concurrent.futures import ProcessPoolExecutor, as_completed
from tqdm import tqdm
from src.engine import generate_report
from src.data_loader import db

# Use total cores minus 2 (to keep Windows OS responsive)
MAX_CORES = max(1, os.cpu_count() - 2)
TARGET_TEMPLATE = "parlimen_dun"

def build_flat_queue():
    """Generates a flat list of every single location task needed."""
    tasks = []
    geo_df = db.dim_geo
    valid_states = geo_df['kod_negeri'].dropna().astype(str).str.replace(r'\.0$', '', regex=True).str.zfill(2).unique()
    
    for state_code in valid_states:
        state_mask = geo_df['kod_negeri'].astype(str).str.replace(r'\.0$', '', regex=True).str.zfill(2) == state_code
        state_geo = geo_df[state_mask]
        
        # 1. Parliaments
        for parl in state_geo['kod_parlimen'].dropna().unique():
            tasks.append({'location_code': parl, 'report_type': 'parlimen', 'parent_code': None})
            
        # 2. DUNs
        duns_subset = state_geo[state_geo['kod_dun'].str.lower() != 'n.a.'][['kod_dun', 'kod_parlimen']].dropna().drop_duplicates()
        for _, row in duns_subset.iterrows():
            tasks.append({'location_code': row['kod_dun'], 'report_type': 'dun', 'parent_code': row['kod_parlimen']})
            
    return tasks

def worker_process(task):
    """Executes a single report generation directly in memory via openpyxl."""
    try:
        generate_report(
            location_code=task['location_code'],
            report_type=task['report_type'],
            parent_code=task['parent_code'],
            template_key=TARGET_TEMPLATE
        )
        return True, task['location_code'], None
    except Exception as e:
        return False, task['location_code'], str(e)

if __name__ == "__main__":
    start_time = time.time()
    print("=============================================")
    print(f"🚀 BARE-METAL HIGH-CONCURRENCY FACTORY")
    print(f"⚙️  Utilizing {MAX_CORES} CPU Cores.")
    print("=============================================\n")
    
    master_queue = build_flat_queue()
    success_count = 0
    fail_list = []
    
    # Unleash the CPU
    with ProcessPoolExecutor(max_workers=MAX_CORES) as executor:
        futures = {executor.submit(worker_process, task): task for task in master_queue}
        
        for future in tqdm(as_completed(futures), total=len(master_queue), desc="Generating Reports", unit="file"):
            success, loc_code, error_msg = future.result()
            if success:
                success_count += 1
            else:
                fail_list.append(f"{loc_code}: {error_msg}")

    elapsed_minutes = (time.time() - start_time) / 60
    print("\n=============================================")
    print(f"🏁 FACTORY RUN COMPLETE in {elapsed_minutes:.2f} minutes!")
    print(f"✅ Successfully Generated: {success_count} files")