"""
src/engine.py
=============
Stage-2 orchestrator: inject metric values into one jadual file for one state.

Two mapper modes
----------------
The engine discovers mappers by scanning `sheets/jadual_*.py`. Each mapper
declares exactly ONE of:

    populate_column = make_populate_location(ROW_MAP, ROUND_DECIMALS)
        → columns are LOCATIONS (P.140, N.01, …)
        → engine calls it once per location column, per page

    populate_page   = make_populate_year(ROW_MAP, TARGET_YEARS, ROUND_DECIMALS)
        → columns are YEARS (right-aligned from COL_L)
        → engine calls it once per page, with the state's metrics

Both modes share:
    JADUAL_NUMBER : str   (required — string, so "1", "2", "2.1" all work)

Page handling
-------------
A jadual file may contain MULTIPLE stacked page blocks (Stage 1 packs a
different set of locations on each page). Each page's header row is parsed
independently, so column layouts may differ per page. Row layout is
identical on every page → mappers declare page-1 rows only, and the engine
adds `row_offset` for pages 2, 3, …

Files whose name doesn't start with `jadual_` (e.g. `_base_template.py`)
are never scanned.
"""
import os
import re
import pkgutil
import importlib
import openpyxl

from config.settings import (
    ROWS_PER_PAGE, KATEGORI_ROW,
    COL_B, COL_C, COL_D, COL_L,
)
from src.data_provider import get_metrics_dict


# ==========================================
# 1. PLUGIN REGISTRY  (dual-mode)
# ==========================================
import sheets  # requires sheets/__init__.py

LOCATION_REGISTRY: dict[str, callable] = {}   # jadual → populate_column
YEAR_REGISTRY:     dict[str, callable] = {}   # jadual → populate_page

print("[ENGINE] Discovering jadual mappers...")
for _, module_name, is_pkg in pkgutil.iter_modules(sheets.__path__):
    if is_pkg or not module_name.startswith("jadual_"):
        continue

    mod = importlib.import_module(f"sheets.{module_name}")

    # Skip placeholder / unfinished mappers
    if not hasattr(mod, "JADUAL_NUMBER"):
        print(f"  [skip] {module_name}: JADUAL_NUMBER not defined")
        continue
    if mod.JADUAL_NUMBER is None:
        print(f"  [skip] {module_name}: JADUAL_NUMBER is None (not configured)")
        continue

    # Always normalise to string so filenames like "Jadual 2.1.xlsx" match
    num = str(mod.JADUAL_NUMBER)

    if hasattr(mod, "populate_column"):
        LOCATION_REGISTRY[num] = mod.populate_column
        print(f"  loaded sheets/{module_name}.py  →  Jadual {num}  [location-mode]")
    elif hasattr(mod, "populate_page"):
        YEAR_REGISTRY[num] = mod.populate_page
        print(f"  loaded sheets/{module_name}.py  →  Jadual {num}  [year-mode]")
    else:
        print(f"  [skip] {module_name}: no populate_column or populate_page")

print(f"[ENGINE] {len(LOCATION_REGISTRY)} location-mapper(s), "
      f"{len(YEAR_REGISTRY)} year-mapper(s) registered.\n")


# ==========================================
# 2. HEADER / PAGE HELPERS
# ==========================================
def detect_location_start_col(ws, header_row: int = KATEGORI_ROW) -> int:
    """
    For location-mode jaduals: if column C's header contains 'Tahun'/'Year',
    the jadual has a year column and locations start at column D.
    Otherwise locations start at C.
    """
    c_val = ws.cell(row=header_row, column=COL_C).value
    has_year = c_val is not None and "tahun" in str(c_val).lower()
    return COL_D if has_year else COL_C


def parse_header_locations(ws,
                           header_row: int = KATEGORI_ROW,
                           start_col: int | None = None,
                           end_col: int = COL_L) -> list[dict]:
    """
    Walk one page's header row and return ordered location dicts:
        {col, kind, kod, nama, parent_parlimen}
    kind ∈ {'negeri', 'parlimen', 'dun'}
    """
    if start_col is None:
        start_col = detect_location_start_col(ws, header_row)

    locations: list[dict] = []
    current_parlimen: str | None = None

    for col in range(start_col, end_col + 1):
        val = ws.cell(row=header_row, column=col).value
        if val is None:
            continue
        text = str(val).strip()
        if not text:
            continue

        # "P.140\nSegamat"  →  kod="P.140", nama="Segamat"
        if "\n" in text:
            kod, nama = text.split("\n", 1)
            kod, nama = kod.strip(), nama.strip()
        else:
            kod, nama = None, text

        if kod and kod.upper().startswith("P."):
            current_parlimen = kod
            locations.append({
                "col": col, "kind": "parlimen",
                "kod": kod, "nama": nama,
                "parent_parlimen": None,
            })
        elif kod and kod.upper().startswith("N."):
            locations.append({
                "col": col, "kind": "dun",
                "kod": kod, "nama": nama,
                "parent_parlimen": current_parlimen,
            })
        else:
            # Fallback: no prefix ⇒ state cell ("Johor")
            locations.append({
                "col": col, "kind": "negeri",
                "kod": None, "nama": nama,
                "parent_parlimen": None,
            })

    return locations


def detect_page_header_rows(ws) -> list[int]:
    """
    Return absolute row numbers of every page-header in the sheet.

    Page 1 header is always at KATEGORI_ROW. Later pages land at
    KATEGORI_ROW + N * ROWS_PER_PAGE (per Stage-1's stacking rule).
    Stop as soon as a candidate row has an empty column-B label.
    """
    headers = [KATEGORI_ROW]
    page = 1
    while True:
        test_row = KATEGORI_ROW + page * ROWS_PER_PAGE
        if test_row > ws.max_row:
            break
        val = ws.cell(row=test_row, column=COL_B).value
        if val is None or not str(val).strip():
            break
        headers.append(test_row)
        page += 1
    return headers


def extract_jadual_number(filename: str) -> str | None:
    """'Jadual 7.xlsx' → '7'. 'Jadual 2.1.xlsx' → '2.1'."""
    m = re.search(r"Jadual[\s_\-]+(\d+(?:\.\d+)?)", filename, re.IGNORECASE)
    return m.group(1) if m else None


# ==========================================
# 3. METRIC LOOKUP
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
# 4. PER-FILE ORCHESTRATOR
# ==========================================
def generate_jadual_for_state(state_name: str,
                              state_code: str,
                              jadual_filename: str,
                              input_path: str,
                              output_path: str) -> str:
    """
    Load one Stage-1 jadual file, populate it across every page,
    and save to output_path.

    Returns a status string starting with ✅ (success) or ❌ (failure).
    """
    from sheets import apply_bilingual_rich_formatting  # avoid top-level import cycle

    jadual_num = extract_jadual_number(jadual_filename)
    if jadual_num is None:
        return f"❌ Cannot parse jadual number from '{jadual_filename}'"

    in_location_mode = jadual_num in LOCATION_REGISTRY
    in_year_mode     = jadual_num in YEAR_REGISTRY
    if not (in_location_mode or in_year_mode):
        return f"❌ No mapper registered for Jadual {jadual_num}"

    wb = None
    try:
        wb = openpyxl.load_workbook(input_path)
        ws = wb.active  # one sheet per jadual file

        page_headers = detect_page_header_rows(ws)
        if not page_headers:
            return f"❌ No page header found in {jadual_filename}"

        total_populated = 0

        for page_header_row in page_headers:
            row_offset = page_header_row - KATEGORI_ROW     # 0, 50, 100, …

            # ---- Dispatch by mode ----
            if in_location_mode:
                mapper = LOCATION_REGISTRY[jadual_num]
                locations = parse_header_locations(ws, header_row=page_header_row)

                for loc in locations:
                    if loc["kind"] == "negeri":
                        continue  # state cell is a label only — no data column

                    metrics = fetch_metrics_for_location(loc, state_code)
                    mapper(ws, loc["col"], metrics,
                           row_offset=row_offset,
                           bold=(loc["kind"] == "parlimen"))
                    total_populated += 1

            else:  # year-mode
                mapper = YEAR_REGISTRY[jadual_num]
                state_metrics = get_metrics_dict(state_code, "negeri")
                written = mapper(ws, state_metrics, row_offset=row_offset)
                total_populated += (written or 0)

            # ---- Re-apply bilingual formatting on the label cells ----
            apply_bilingual_rich_formatting(ws, page_header_row, COL_B)
            # Only location-mode jaduals have a "Tahun"/"Year" column in C.
            # Year-mode jaduals use column C as a data column.
            if in_location_mode and detect_location_start_col(ws, page_header_row) == COL_D:
                apply_bilingual_rich_formatting(ws, page_header_row, COL_C)

        os.makedirs(os.path.dirname(output_path), exist_ok=True)
        wb.save(output_path)
        wb.close()
        wb = None

        mode_tag = "location" if in_location_mode else "year"
        return (f"✅ {state_name}/{jadual_filename}  "
                f"[{mode_tag}] {total_populated} cell-group(s) across "
                f"{len(page_headers)} page(s)")

    except Exception as e:
        return f"❌ FAILED {state_name}/{jadual_filename} | {e}"
    finally:
        if wb is not None:
            try:
                wb.close()
            except Exception:
                pass