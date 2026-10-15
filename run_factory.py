"""
run_factory.py
==============
Data injection factory — Stage 2 of the pipeline.

Reads Stage-1 output from OUTPUT_LOCATIONS_DIR/{State}/Jadual N.xlsx,
injects metric values into each location column, writes to
OUTPUT_FINAL_DIR/{State}/Jadual N.xlsx.

Examples
--------
# Run EVERYTHING (all states, all jaduals)
python3 run_factory.py

# Only Johor, all jaduals
python3 run_factory.py --state Johor

# Johor + Melaka, all jaduals
python3 run_factory.py --state Johor Melaka

# Only Jadual 1, all states
python3 run_factory.py --jadual 1

# Johor, only Jadual 1 and 5
python3 run_factory.py --state Johor --jadual 1 2

# Preview the task list without executing
python3 run_factory.py --dry-run
"""

import os
import sys
import time
import argparse
from collections import Counter
from concurrent.futures import ProcessPoolExecutor, as_completed
from tqdm import tqdm

from config.settings import OUTPUT_LOCATIONS_DIR, OUTPUT_FINAL_DIR
from src.engine import generate_jadual_for_state, extract_jadual_number

MAX_CORES = max(1, os.cpu_count() - 2)


# ==========================================
# STATE LOOKUP  (folder name → 2-digit code)
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
        "--jadual", nargs="+", type=int, default=None, metavar="N",
        help="Target jadual number(s), e.g. --jadual 1 5 13. "
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
        msg = generate_jadual_for_state(
            state_name      = task["state_name"],
            state_code      = task["state_code"],
            jadual_filename = task["jadual_file"],
            input_path      = task["input_path"],
            output_path     = task["output_path"],
        )
        return msg.startswith("✅"), msg
    except Exception as e:
        return False, f"❌ CRASH {task['state_name']}/{task['jadual_file']}: {e}"


# ==========================================
# MAIN
# ==========================================
if __name__ == "__main__":
    args = parse_args()
    start = time.time()

    print("=" * 60)
    print("🚀 DATA INJECTION FACTORY (Stage 2)")
    print("=" * 60)
    print(f"  States  : {args.state if args.state else 'ALL'}")
    print(f"  Jaduals : {args.jadual if args.jadual else 'ALL'}")
    print(f"  Cores   : {args.cores}")
    print(f"  Input   : {OUTPUT_LOCATIONS_DIR}")
    print(f"  Output  : {OUTPUT_FINAL_DIR}")
    print("=" * 60 + "\n")

    queue = build_flat_queue(
        target_states=args.state,
        target_jaduals=args.jadual,
    )

    if not queue:
        print("⚠️  No tasks match the given filters.")
        sys.exit(0)

    per_state = Counter(t["state_name"] for t in queue)
    print(f"📁 {len(queue)} task(s) queued across {len(per_state)} state(s):")
    for state, n in sorted(per_state.items()):
        print(f"     {state:<24} {n} jadual(s)")
    print()

    if args.dry_run:
        print("🔎 --dry-run: task list only, nothing executed.")
        for t in queue[:20]:
            print(f"     {t['state_name']}/{t['jadual_file']}")
        if len(queue) > 20:
            print(f"     ... and {len(queue) - 20} more")
        sys.exit(0)

    success, fail = 0, 0
    failed_tasks = []

    with ProcessPoolExecutor(max_workers=args.cores) as ex:
        futures = {ex.submit(worker, t): t for t in queue}
        for fut in tqdm(as_completed(futures), total=len(queue),
                        desc="Injecting data", unit="file"):
            ok, msg = fut.result()
            if ok:
                success += 1
            else:
                fail += 1
                failed_tasks.append(msg)
                tqdm.write(msg)

    print("\n" + "=" * 60)
    print(f"🏁 Done in {(time.time() - start) / 60:.2f} min")
    print(f"✅ {success} succeeded   ❌ {fail} failed")
    if failed_tasks:
        print("\nFailures:")
        for m in failed_tasks[:20]:
            print(f"  {m}")
        if len(failed_tasks) > 20:
            print(f"  ... and {len(failed_tasks) - 20} more")
    print("=" * 60)