import os
import xlwings as xw
from concurrent.futures import ProcessPoolExecutor, as_completed
from config.settings import OUTPUT_DIR

# ==========================================
# CONFIGURATION
# ==========================================
MAX_PDF_WORKERS = 4  # Keep between 3 to 5 to prevent Excel Print Spooler crashes

# ==========================================
# PHASE 2: THE CRAWLER & RESUME FEATURE
# ==========================================
def get_unconverted_files(base_directory):
    """
    Scans the output directory for .xlsx files. 
    Checks if a .pdf version already exists. If not, adds it to the queue.
    """
    pending_files = []
    
    for root, dirs, files in os.walk(base_directory):
        for file in files:
            # Only target Excel files and ignore hidden temporary files (like ~$Jadual.xlsx)
            if file.endswith('.xlsx') and not file.startswith('~$'):
                excel_path = os.path.join(root, file)
                pdf_path = os.path.splitext(excel_path)[0] + '.pdf'
                
                # The "Resume" Logic: Only add to queue if the PDF doesn't exist yet
                if not os.path.exists(pdf_path):
                    pending_files.append(excel_path)
                    
    return pending_files

# ==========================================
# PHASE 3: THE ISOLATED CONVERTER ENGINE
# ==========================================
def convert_single_file_to_pdf(excel_relative_path):
    """
    Opens a single Excel file in a hidden instance and renders the PDF.
    Runs entirely in its own isolated background process.
    """
    # Windows COM and Mac AppleScript REQUIRE absolute paths to work safely
    excel_abs_path = os.path.abspath(excel_relative_path)
    pdf_abs_path = os.path.splitext(excel_abs_path)[0] + '.pdf'
    
    file_name = os.path.basename(excel_abs_path)
    
    # Spin up an isolated Excel instance just for this file
    app = xw.App(visible=False)
    app.screen_updating = False
    app.display_alerts = False
    
    try:
        # Open the workbook
        wb = app.books.open(excel_abs_path)
        
        # Native xlwings PDF export (Wraps ExportAsFixedFormat on Windows)
        wb.to_pdf(path=pdf_abs_path)
        
        return f"✅ PDF Created: {file_name}"
        
    except Exception as e:
        return f"❌ FAILED to convert {file_name} | Error: {e}"
        
    finally:
        # GUARANTEE the workbook is closed without saving (prevents prompt hang)
        try:
            wb.close()
        except:
            pass
        # GUARANTEE the app is killed
        app.quit()


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
    
    # Feed the queue into the Multi-Process Factory
    with ProcessPoolExecutor(max_workers=MAX_PDF_WORKERS) as executor:
        
        # Submit all pending files to the workers
        futures = {executor.submit(convert_single_file_to_pdf, file_path): file_path for file_path in files_to_convert}
        
        # Monitor the queue as files finish rendering
        for future in as_completed(futures):
            result = future.result()
            print(result)
            
            if "✅" in result:
                success_count += 1
            else:
                fail_count += 1

    print("\n=============================================")
    print(f"🏁 PDF Batch Conversion Complete!")
    print(f"   - Successfully Converted: {success_count}")
    print(f"   - Failed: {fail_count}")
    print("=============================================")