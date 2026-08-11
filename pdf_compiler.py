import os
import logging
import xlwings as xw
from concurrent.futures import ProcessPoolExecutor, as_completed
from config.settings import OUTPUT_DIR, BASE_DIR
from pypdf import PdfWriter

# Mute pypdf warnings (prevents terminal spam from misaligned external PDFs)
logging.getLogger("pypdf").setLevel(logging.ERROR)

# ==========================================
# CONFIGURATION
# ==========================================
MAX_PDF_WORKERS = 4  
SEPARATOR_DIR = os.path.join(BASE_DIR, 'separator')
BLANK_PAGE_PATH = os.path.join(SEPARATOR_DIR, '00_BLANK_PAGE.pdf')

# ==========================================
# PHASE 1: THE BLUEPRINTS
# ==========================================
PARLIMEN_BLUEPRINT = [
    {"type": "pdf", "filename": "1 maklumat asas.pdf"},
    {"type": "excel_chunk", "sheets": ["1.0_Maklumat_Asas"]},
    
    {"type": "pdf", "filename": "2 penduduk.pdf"},
    {"type": "excel_chunk", "sheets": [
        "2.0_Penduduk_Malaysia", 
        "2.1_Penduduk_Negeri", 
        "2.2_Penduduk_Parlimen",
        "2.3_Penduduk"
    ]},
    
    {"type": "pdf", "filename": "3 perumahan.pdf"},
    {"type": "excel_chunk", "sheets": ["3.0_Perumahan"]},

    {"type": "pdf", "filename": "4 guna tenaga.pdf"},
    {"type": "excel_chunk", "sheets": ["4.0_Guna_Tenaga"]},
    
    {"type": "pdf", "filename": "5 pendapatan.pdf"},
    {"type": "excel_chunk", "sheets": ["5.0_Pendapatan"]},

    {"type": "pdf", "filename": "6 pendidikan.pdf"},
    {"type": "excel_chunk", "sheets": ["6.0_Pendidikan_SK", "6.1_Pendidikan_Swasta"]},

    {"type": "pdf", "filename": "7 kesihatan.pdf"},
    {"type": "excel_chunk", "sheets": ["7.0_Kesihatan"]},

    {"type": "pdf", "filename": "8 kemiskinan.pdf"},
    {"type": "excel_chunk", "sheets": [
        "8.0_Jantina",
        "8.1_Etnik",
        "8.2_Agama",
        "8.3_Umur"
    ]},

    {"type": "pdf", "filename": "9 keselamatan.pdf"},
    {"type": "excel_chunk", "sheets": ["9.0_Keselamatan_Awam"]},

    {"type": "pdf", "filename": "10 internet.pdf"},
    {"type": "excel_chunk", "sheets": ["10.0_IMS"]},

    {"type": "pdf", "filename": "11 kemudahan.pdf"},
    {"type": "excel_chunk", "sheets": ["11.0_Air_Elektrik_Sampah"]},

    {"type": "pdf", "filename": "11.5 ekonomi.pdf"},

    {"type": "pdf", "filename": "12 pertubuhan.pdf"},
    {"type": "excel_chunk", "sheets": [
        "12.0_Pertubuhan",
        "12.1_Prtubuhn_Perkhdmtn", 
        "12.1_Prtubuhn_Perkhdmtn_(samb.)",
        "12.2_Kemudahan_Asas"
    ]},

    {"type": "pdf", "filename": "13 statistik-statistik lain.pdf"},
    {"type": "excel_chunk", "sheets": [
        "13.0_Fasiliti_awam",
        "13.0_Fasiliti_awam_samb", 
        "13.1_Statistik_Perkhid._lain",
        "13.2_Koperasi"
    ]}
]

MALAYSIA_BLUEPRINT = [
    {"type": "pdf", "filename": "14 malaysia.pdf"},
    {"type": "excel_chunk", "sheets": [
        "14.0", 
        "14.1", 
        "20.0", 
        "31.0"
    ]}
]

NEGERI_BLUEPRINT = [
    {"type": "pdf", "filename": "15 negeri.pdf"},
]


# ==========================================
# PHASE 2: THE CRAWLER & RESUME FEATURE
# ==========================================
def get_unconverted_files(base_directory):
    pending_files = []
    for root, dirs, files in os.walk(base_directory):
        for file in files:
            if file.endswith('.xlsx') and not file.startswith('~$'):
                excel_path = os.path.join(root, file)
                pdf_path = os.path.splitext(excel_path)[0] + '.pdf'
                if not os.path.exists(pdf_path):
                    pending_files.append(excel_path)
    return pending_files


# ==========================================
# PHASE 3: THE ISOLATED CHUNK ENGINE
# ==========================================
def convert_single_file_to_pdf(excel_relative_path):
    excel_abs_path = os.path.abspath(excel_relative_path)
    base_name = os.path.splitext(os.path.basename(excel_abs_path))[0]
    final_pdf_path = os.path.splitext(excel_abs_path)[0] + '.pdf'
    
    if "Parlimen" in excel_abs_path or "DUN" in excel_abs_path:
        blueprint = PARLIMEN_BLUEPRINT
    elif "Negeri" in excel_abs_path:
        blueprint = NEGERI_BLUEPRINT
    else:
        blueprint = MALAYSIA_BLUEPRINT
        
    app = xw.App(visible=False)
    app.screen_updating = False
    app.display_alerts = False
    
    temp_chunks = []
    merger = PdfWriter()
    
    try:
        wb = app.books.open(excel_abs_path)
        
        for i, block in enumerate(blueprint):
            if block["type"] == "pdf":
                sep_path = os.path.join(SEPARATOR_DIR, block["filename"])
                if os.path.exists(sep_path):
                    merger.append(sep_path)
                    
                    if os.path.exists(BLANK_PAGE_PATH):
                        merger.append(BLANK_PAGE_PATH)
                else:
                    print(f"  -> [Warning] Missing Separator Asset: {block['filename']}")
            
            elif block["type"] == "excel_chunk":
                valid_sheets = [s for s in block["sheets"] if s in [sht.name for sht in wb.sheets]]
                
                if valid_sheets:
                    wb.api.Worksheets(valid_sheets).Select()
                    chunk_path = os.path.join(os.path.dirname(excel_abs_path), f"temp_chunk_{i}_{base_name}.pdf")
                    app.api.ActiveSheet.ExportAsFixedFormat(0, chunk_path)
                    
                    temp_chunks.append(chunk_path)
                    merger.append(chunk_path)
                    
        merger.write(final_pdf_path)
        merger.close()
        
        return f"✅ PDF Assembled: {base_name}.pdf"
        
    except Exception as e:
        return f"❌ FAILED to convert {base_name} | Error: {e}"
        
    finally:
        try:
            wb.close()
        except:
            pass
        app.quit()
        
        for chunk in temp_chunks:
            if os.path.exists(chunk):
                try:
                    os.remove(chunk)
                except:
                    pass


# ==========================================
# PHASE 4: THE MULTI-PROCESS DISPATCHER
# ==========================================
if __name__ == "__main__":
    print("=============================================")
    print("🔍 Scanning Output Directory for pending PDFs...")
    
    target_dir = os.path.abspath(OUTPUT_DIR)
    files_to_convert = get_unconverted_files(target_dir)
    
    if not files_to_convert:
        print("✅ No pending Excel files found. All PDFs are up to date!")
        print("=============================================")
        exit()
        
    print(f"📁 Found {len(files_to_convert)} files needing conversion.")
    print(f"🚀 Spinning up {MAX_PDF_WORKERS} parallel PDF rendering engines...")
    print("=============================================\n")
    
    success_count = 0
    fail_count = 0
    
    with ProcessPoolExecutor(max_workers=MAX_PDF_WORKERS) as executor:
        futures = {executor.submit(convert_single_file_to_pdf, file_path): file_path for file_path in files_to_convert}
        
        for future in as_completed(futures):
            result = future.result()
            print(result)
            if "✅" in result:
                success_count += 1
            else:
                fail_count += 1

    print("\n=============================================")
    print(f"🏁 PDF Batch Conversion Complete!")
    print(f"   - Successfully Assembled: {success_count}")
    print(f"   - Failed: {fail_count}")
    print("=============================================")