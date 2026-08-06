import xlwings as xw
from src.engine import generate_report
from src.data_loader import db

# ==========================================
# CONFIGURATION
# ==========================================
TEST_MODE = True
TARGET_TEMPLATE = "malaysia"  # NEW: Specifies which template profile to process
TARGET_STATE = "01"
TARGET_PARLIMEN = "P.143"
TARGET_DUN = "N.07"

def get_parliaments_for_state(state_code: str):
    """Fetches all unique Parliament codes for a specific State."""
    geo_df = db.dim_geo
    clean_db_codes = geo_df['kod_negeri'].astype(str).str.replace(r'\.0$', '', regex=True).str.zfill(2)
    clean_target = str(state_code).zfill(2)
    mask = clean_db_codes == clean_target
    return geo_df[mask]['kod_parlimen'].dropna().unique().tolist()

def get_duns_for_state(state_code: str):
    """
    Fetches all unique DUNs for a specific State and pairs them with their 
    parent Parliament code to guarantee uniqueness.
    """
    geo_df = db.dim_geo
    clean_db_codes = geo_df['kod_negeri'].astype(str).str.replace(r'\.0$', '', regex=True).str.zfill(2)
    clean_target = str(state_code).zfill(2)
    
    # Filter by state and exclude "n.a." DUN codes
    mask = (clean_db_codes == clean_target) & (geo_df['kod_dun'].astype(str).str.lower() != 'n.a.')
    
    # Extract both columns and drop duplicates
    duns_df = geo_df[mask][['kod_dun', 'kod_parlimen']].drop_duplicates()
    
    # Return as a list of dictionaries: [{'kod_dun': 'N.07', 'kod_parlimen': 'P.143'}, ...]
    return duns_df.to_dict('records')

if __name__ == "__main__":
    print("Starting Global Excel Instance...")
    app = xw.App(visible=False)
    app.screen_updating = False 
    app.display_alerts = False
    
    try:
        if TEST_MODE:
            print("\n=============================================")
            print(f"[TEST MODE ACTIVE] Executing surgical run...")
            print("=============================================")
            # 1. Run target Parliament
            generate_report(
                location_code=TARGET_PARLIMEN, 
                report_type='parlimen', 
                excel_app=app,
                template_key=TARGET_TEMPLATE
            )
            # 2. Run target DUN (Passing the Parent Parliament for strict matching)
            generate_report(
                location_code=TARGET_DUN, 
                report_type='dun', 
                excel_app=app, 
                parent_code=TARGET_PARLIMEN,
                template_key=TARGET_TEMPLATE
            )
            
        else:
            print(f"\n[PRODUCTION MODE] Fetching hierarchy for State Code '{TARGET_STATE}'...")
            target_parlimen = get_parliaments_for_state(TARGET_STATE)
            target_duns = get_duns_for_state(TARGET_STATE)
            
            if not target_parlimen:
                print(f"Error: No locations found for state code '{TARGET_STATE}'.")
            else:
                print(f"Found {len(target_parlimen)} Parliaments and {len(target_duns)} DUNs.")
                
                print("\n--- Starting Phase A: Parliament Reports ---")
                for parl in target_parlimen:
                    generate_report(
                        location_code=parl, 
                        report_type='parlimen', 
                        excel_app=app,
                        template_key=TARGET_TEMPLATE
                    )
                    
                print("\n--- Starting Phase B: DUN Reports ---")
                for dun_info in target_duns:
                    # Pass both the DUN code and its parent Parliament code
                    generate_report(
                        location_code=dun_info['kod_dun'], 
                        report_type='dun', 
                        excel_app=app, 
                        parent_code=dun_info['kod_parlimen'],
                        template_key=TARGET_TEMPLATE
                    )
                    
    except Exception as e:
        print(f"\nBatch generation interrupted: {e}")
    finally:
        print("\nShutting down Excel engine...")
        app.quit()
        print("Batch Processing Complete!")