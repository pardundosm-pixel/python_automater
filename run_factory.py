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
def build_state_code_map():
    """{'Johor': '01', 'Melaka': '02', ...} from dim_geo."""
    from src.data_loader import db   # lazy: keeps worker startup light
    geo = db.dim_geo

    m = {}
    for _, row in geo[["kod_negeri", "nama_negeri"]].drop_duplicates().iterrows():
        code = str(row["kod_negeri"]).replace(".0", "").zfill(2)
        name = str(row["nama_negeri"]).strip()
        m[name] = code
    return m


# ==========================================
# CLI
# ==========================================
def parse_args():
    p = argparse.ArgumentParser(
        description="Stage-2 data injection factory (openpyxl).",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=__doc__,
    )
    p.add_argument(
        "--state", nargs="+", default=None, metavar="NAME",
        help="Target state folder name(s), e.g. --state Johor Melaka. "
             "Default: every folder found in OUTPUT_LOCATIONS_DIR.",
    )
    p.add_argument(
        "--jadual", nargs="+", type=str, default=None, metavar="N",
        help="Target jadual number(s), e.g. --jadual 1 5 2.1. "
             "Default: every jadual found in each state folder.",
    )
    p.add_argument(
        "--cores", type=int, default=MAX_CORES,
        help=f"Max parallel workers (default: {MAX_CORES}).",
    )
    p.add_argument(
        "--dry-run", action="store_true",
        help="List tasks but do not execute.",
    )
    return p.parse_args()


# ==========================================
# TASK QUEUE
# ==========================================
def build_flat_queue(target_states=None, target_jaduals=None):
    """Return one task per (state folder, jadual file)."""
    if not os.path.isdir(OUTPUT_LOCATIONS_DIR):
        print(f"❌ Input dir not found: {OUTPUT_LOCATIONS_DIR}")
        print("   Did Stage 1 (location_injection_batch.py) run first?")
        return []

    state_map = build_state_code_map()
    tasks = []

    available_states = sorted(
        f for f in os.listdir(OUTPUT_LOCATIONS_DIR)
        if os.path.isdir(os.path.join(OUTPUT_LOCATIONS_DIR, f))
    )

    if target_states:
        requested = set(target_states)
        missing = requested - set(available_states)
        if missing:
            print(f"⚠️  States not found in {OUTPUT_LOCATIONS_DIR}: "
                  f"{', '.join(sorted(missing))}")
        states_to_run = [s for s in available_states if s in requested]
    else:
        states_to_run = available_states

    for state_folder in states_to_run:
        state_dir = os.path.join(OUTPUT_LOCATIONS_DIR, state_folder)
        state_code = state_map.get(state_folder)
        if state_code is None:
            print(f"  [warn] No state code for folder '{state_folder}', skipping.")
            continue

        jadual_files = sorted(
            f for f in os.listdir(state_dir)
            if f.endswith(".xlsx") and not f.startswith("~$")
        )

        for jadual_file in jadual_files:
            if target_jaduals is not None:
                num = extract_jadual_number(jadual_file)
                if num is None or num not in target_jaduals:
                    continue

            tasks.append({
                "state_name":  state_folder,
                "state_code":  state_code,
                "jadual_file": jadual_file,
                "input_path":  os.path.join(state_dir, jadual_file),
                "output_path": os.path.join(OUTPUT_FINAL_DIR, state_folder, jadual_file),
            })

    return tasks


# ==========================================
# WORKER
# ==========================================
def worker(task):
    try:
        success_count = 0
        
        # ------------------------------------------------
        # ROUTE 1: NEGERI REPORT (1 State = 1 Report)
        # ------------------------------------------------
        if template_key == "negeri":
            generate_report(
                location_code=state_code, 
                report_type='negeri', 
                excel_app=app,
                template_key=template_key
            )
            return f"✅ [State {state_code}] Successfully generated 1 Negeri Report."
            
        # ------------------------------------------------
        # ROUTE 2: PARLIMEN & DUN REPORTS (1 State = Many Reports)
        # ------------------------------------------------
        elif template_key == "parlimen_dun":
            target_parlimen = get_parliaments_for_state(state_code)
            target_duns = get_duns_for_state(state_code)
            
            # Generate all Parliament Reports for this state
            for parl in target_parlimen:
                generate_report(
                    location_code=parl, 
                    report_type='parlimen', 
                    excel_app=app,
                    template_key=template_key
                )
                success_count += 1
                
            # Generate all DUN Reports for this state
            for dun_info in target_duns:
                generate_report(
                    location_code=dun_info['kod_dun'], 
                    report_type='dun', 
                    excel_app=app, 
                    parent_code=dun_info['kod_parlimen'],
                    template_key=template_key
                )
                success_count += 1
                
            return f"✅ [State {state_code}] Successfully generated {success_count} Parlimen & DUN reports."
            
        else:
            return f"❌ [State {state_code}] FAILED: Unknown template key '{template_key}'"
            
    except Exception as e:
        return f"❌ [State {state_code}] FAILED with error: {e}"
        
    finally:
        # GUARANTEE the hidden Excel app is killed
        try:
            app.quit()
        except:
            pass


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