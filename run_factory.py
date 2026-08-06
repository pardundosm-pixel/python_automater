import os
import xlwings as xw
from concurrent.futures import ProcessPoolExecutor, as_completed
from src.engine import generate_report
from src.data_loader import db

# ==========================================
# FACTORY CONFIGURATION
# ==========================================
MAX_PARALLEL_WORKERS = 5  # Safe sweet spot for COM stability
TARGET_TEMPLATE = "parlimen_dun" 
# For testing safely, we will only run 2 states instead of all 16
# Change this back to [str(i).zfill(2) for i in range(1, 17)] for full production
TEST_STATES = ["01", "02"] # Johor and Melaka


def get_parliaments_for_state(state_code: str):
    """Fetches all unique Parliament codes for a specific State."""
    geo_df = db.dim_geo
    clean_db_codes = geo_df['kod_negeri'].astype(str).str.replace(r'\.0$', '', regex=True).str.zfill(2)
    clean_target = str(state_code).zfill(2)
    mask = clean_db_codes == clean_target
    return geo_df[mask]['kod_parlimen'].dropna().unique().tolist()


def get_duns_for_state(state_code: str):
    """Fetches all unique DUNs for a specific State paired with parent Parliament."""
    geo_df = db.dim_geo
    clean_db_codes = geo_df['kod_negeri'].astype(str).str.replace(r'\.0$', '', regex=True).str.zfill(2)
    clean_target = str(state_code).zfill(2)
    
    mask = (clean_db_codes == clean_target) & (geo_df['kod_dun'].str.lower() != 'n.a.')
    subset = geo_df[mask][['kod_dun', 'kod_parlimen']].dropna().drop_duplicates()
    return subset.to_dict('records')


# ==========================================
# THE ISOLATED WORKER FUNCTION
# ==========================================
def process_single_state(state_code, template_key):
    """Runs entirely in its own parallel background process."""
    target_parlimen = get_parliaments_for_state(state_code)
    target_duns = get_duns_for_state(state_code)
    
    if not target_parlimen:
        return f"[State {state_code}] Skipped: No locations found."

    # Spin up a DEDICATED Excel App for this specific worker
    app = xw.App(visible=False)
    app.screen_updating = False
    app.display_alerts = False
    
    success_count = 0
    
    try:
        # Generate Parliament Reports
        for parl in target_parlimen:
            generate_report(
                location_code=parl, 
                report_type='parlimen', 
                excel_app=app,
                template_key=template_key
            )
            success_count += 1
            
        # Generate DUN Reports
        for dun_info in target_duns:
            generate_report(
                location_code=dun_info['kod_dun'], 
                report_type='dun', 
                excel_app=app, 
                parent_code=dun_info['kod_parlimen'],
                template_key=template_key
            )
            success_count += 1
            
        return f"✅ [State {state_code}] Successfully generated {success_count} reports."
        
    except Exception as e:
        return f"❌ [State {state_code}] FAILED with error: {e}"
        
    finally:
        # GUARANTEE the hidden Excel app is killed
        app.quit()


# ==========================================
# MAIN DISPATCHER
# ==========================================
if __name__ == "__main__":
    print(f"=============================================")
    print(f"Starting Multi-Process Factory for: {TARGET_TEMPLATE}")
    print(f"Spinning up parallel Excel workers...")
    print(f"=============================================\n")
    
    with ProcessPoolExecutor(max_workers=MAX_PARALLEL_WORKERS) as executor:
        
        # Submit the test states to the pool
        futures = {executor.submit(process_single_state, state, TARGET_TEMPLATE): state for state in TEST_STATES}
        
        for future in as_completed(futures):
            result_message = future.result()
            print(result_message)
            
    print("\n=============================================")
    print("Batch Processing Complete! All workers shut down safely.")
    print("=============================================")