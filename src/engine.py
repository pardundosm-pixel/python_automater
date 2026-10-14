"""
src/engine.py
=============
Stage-2 orchestrator: inject metric values into one jadual file for one state.

Design
------
* Plugin registry keyed by JADUAL NUMBER (not sheet-name prefix).
* Each `sheets/jadual_N.py` must expose:
      JADUAL_NUMBER   : int
      populate_column : (ws, col_idx, metrics_by_year) -> None
* Engine locates every location column by parsing the header row of the
  Stage-1 output file. Mappers are location-agnostic: they only know the
  row layout of *their* jadual.
"""
import os
import re
import pkgutil
import importlib
import openpyxl

from config.settings import OUTPUT_FINAL_DIR  # noqa: F401  (re-exported for convenience)
from src.data_provider import get_metrics_dict


# ==========================================
# 1. PLUGIN REGISTRY
# ==========================================
import sheets  # requires sheets/__init__.py

MASTER_REGISTRY: dict[int, callable] = {}

print("[ENGINE] Discovering jadual mappers...")
for _, module_name, is_pkg in pkgutil.iter_modules(sheets.__path__):
    if is_pkg or not module_name.startswith("jadual_"):
        continue
    mod = importlib.import_module(f"sheets.{module_name}")
    if not hasattr(mod, "JADUAL_NUMBER") or not hasattr(mod, "populate_column"):
        print(f"  [skip] {module_name}: missing JADUAL_NUMBER or populate_column")
        continue
    MASTER_REGISTRY[int(mod.JADUAL_NUMBER)] = mod.populate_column
    print(f"  loaded sheets/{module_name}.py  →  Jadual {mod.JADUAL_NUMBER}")

print(f"[ENGINE] {len(MASTER_REGISTRY)} jadual mapper(s) registered.\n")


# ==========================================
# 2. LAYOUT CONSTANTS  (mirror location_injection_batch.py)
# ==========================================
KATEGORI_ROW = 6          # header row where locations live
COL_B        = 2          # "MAKLUMAT ASAS" label column
COL_C        = 3          # year column (or first location col if no year)
COL_D        = 4          # first location col when a year column exists
COL_L        = 12         # rightmost location column


# ==========================================
# 3. HEADER PARSER
# ==========================================
def detect_location_start_col(ws, header_row: int = KATEGORI_ROW) -> int:
    """
    If the C-column header contains 'Tahun'/'Year', the jadual has a year
    column and locations start at D. Otherwise they start at C.
    """
    c_val = ws.cell(row=header_row, column=COL_C).value
    has_year = c_val is not None and "tahun" in str(c_val).lower()
    return COL_D if has_year else COL_C


def parse_header_locations(ws, header_row: int = KATEGORI_ROW,
                           start_col: int | None = None,
                           end_col: int = COL_L):
    """
    Walk the header row and return an ordered list of dicts:
        {col, kind, kod, nama, parent_parlimen}

    kind ∈ {'negeri', 'parlimen', 'dun'}
    parent_parlimen is filled for DUNs from the last-seen parliament to the left.
    """
    if start_col is None:
        start_col = detect_location_start_col(ws, header_row)

    locations = []
    current_parlimen = None

    for col in range(start_col, end_col + 1):
        val = ws.cell(row=header_row, column=col).value
        if val is None:
            continue
        text = str(val).strip()
        if not text:
            continue

        # "P.140\nSegamat"  ->  kod, nama
        if "\n" in text:
            kod, nama = text.split("\n", 1)
            kod, nama = kod.strip(), nama.strip()
        else:
            kod, nama = None, text

        if kod and kod.upper().startswith("P."):
            current_parlimen = kod
            locations.append({
                "col": col, "kind": "parlimen",
                "kod": kod, "nama": nama, "parent_parlimen": None,
            })
        elif kod and kod.upper().startswith("N."):
            locations.append({
                "col": col, "kind": "dun",
                "kod": kod, "nama": nama, "parent_parlimen": current_parlimen,
            })
        else:
            # Fallback: no prefix ⇒ state cell ("Johor")
            locations.append({
                "col": col, "kind": "negeri",
                "kod": None, "nama": nama, "parent_parlimen": None,
            })

    return locations


def extract_jadual_number(filename: str) -> int | None:
    """'Jadual 7.xlsx' → 7.  Accepts 'Jadual 7.xlsx', 'Jadual_7.xlsx', etc."""
    m = re.search(r"Jadual[\s_]+(\d+)", filename, re.IGNORECASE)
    return int(m.group(1)) if m else None


# ==========================================
# 4. METRIC LOOKUP
# ==========================================
def fetch_metrics_for_location(loc: dict, state_code: str) -> dict:
    """Return {year: {metric: value}} for one location, or {} on failure."""
    kind = loc["kind"]
    try:
        if kind == "parlimen":
            return get_metrics_dict(loc["kod"], "parlimen")
        if kind == "dun":
            return get_metrics_dict(loc["kod"], "dun",
                                    parent_code=loc["parent_parlimen"])
        if kind == "negeri":
            return get_metrics_dict(state_code, "negeri")
    except Exception as e:
        print(f"  [warn] metric fetch failed for {loc}: {e}")
    return {}


# ==========================================
# 5. PER-FILE ORCHESTRATOR
# ==========================================
def generate_jadual_for_state(state_name: str,
                              state_code: str,
                              jadual_filename: str,
                              input_path: str,
                              output_path: str) -> str:
    """
    Load one Stage-1 jadual file, populate every location column,
    and save to output_path.  Returns a status string starting with
    ✅ (success) or ❌ (failure).
    """
    jadual_num = extract_jadual_number(jadual_filename)
    if jadual_num is None:
        return f"❌ Cannot parse jadual number from '{jadual_filename}'"

    mapper = MASTER_REGISTRY.get(jadual_num)
    if mapper is None:
        return f"❌ No mapper registered for Jadual {jadual_num}"

    wb = None
    try:
        wb = openpyxl.load_workbook(input_path)
        ws = wb.active  # one sheet per jadual file

        locations = parse_header_locations(ws)
        if not locations:
            return f"❌ No locations found in header row of {jadual_filename}"

        # Populate every non-state column
        populated = 0
        for loc in locations:
            if loc["kind"] == "negeri":
                continue  # state cell is a label only — no data column
            metrics = fetch_metrics_for_location(loc, state_code)
            mapper(ws, loc["col"], metrics)
            populated += 1

        os.makedirs(os.path.dirname(output_path), exist_ok=True)
        wb.save(output_path)
        wb.close()
        wb = None

        return (f"✅ {state_name}/{jadual_filename}  "
                f"({populated} location column(s) populated)")

    except Exception as e:
        return f"❌ FAILED {state_name}/{jadual_filename} | {e}"
    finally:
        if wb is not None:
            try:
                wb.close()
            except Exception:
                pass