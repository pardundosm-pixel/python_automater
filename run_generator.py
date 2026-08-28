import argparse
from src.engine import generate_report

def main():
    parser = argparse.ArgumentParser(description="Surgical Test Generator for Subnasional Reports")
    
    parser.add_argument(
        '--template', 
        type=str, 
        required=True, 
        choices=['malaysia', 'negeri', 'parlimen_dun'],
        help="Which template profile to run."
    )
    parser.add_argument('--state', type=str, help="State Code (e.g., '01' for Johor)")
    parser.add_argument('--parlimen', type=str, help="Parliament Code (e.g., 'P.143')")
    parser.add_argument('--dun', type=str, help="DUN Code (e.g., 'N.07')")
    parser.add_argument('--sheets', nargs='+', help="Specific sheets to run (e.g., '1.0' '2.1' '4.0')")

    args = parser.parse_args()

    print("\n=============================================")
    print(f"🚀 SURGICAL TEST INITIATED (Openpyxl Engine)")
    print("=============================================")

    try:
        # --- ROUTE 1: MALAYSIA ---
        if args.template == "malaysia":
            print("--- Testing Phase D: Malaysia Report ---")
            generate_report(
                location_code="00", 
                report_type='malaysia', 
                template_key=args.template, 
                allowed_sheets=args.sheets
            )

        # --- ROUTE 2: NEGERI ---
        elif args.template == "negeri":
            if not args.state:
                raise ValueError("You must provide --state when testing the 'negeri' template.")
            print(f"--- Testing Phase C: Negeri Report (State {args.state}) ---")
            generate_report(
                location_code=args.state, 
                report_type='negeri', 
                template_key=args.template, 
                allowed_sheets=args.sheets
            )

        # --- ROUTE 3: PARLIMEN & DUN ---
        elif args.template == "parlimen_dun":
            if not args.parlimen:
                raise ValueError("You must provide --parlimen when testing the 'parlimen_dun' template.")
                
            # Run target Parliament
            print(f"--- Testing Phase A: Parliament Report ({args.parlimen}) ---")
            generate_report(
                location_code=args.parlimen, 
                report_type='parlimen', 
                template_key=args.template, 
                allowed_sheets=args.sheets
            )
            
            # Run target DUN (if provided)
            if args.dun:
                print(f"\n--- Testing Phase B: DUN Report ({args.dun}) ---")
                generate_report(
                    location_code=args.dun, 
                    report_type='dun', 
                    parent_code=args.parlimen, 
                    template_key=args.template,
                    allowed_sheets=args.sheets
                )
            else:
                print("\n[!] No --dun provided. Skipping Phase B.")

    except Exception as e:
        print(f"\n❌ Surgical run failed: {e}")
    finally:
        print("\n🏁 Test Execution Complete!")

if __name__ == "__main__":
    main()