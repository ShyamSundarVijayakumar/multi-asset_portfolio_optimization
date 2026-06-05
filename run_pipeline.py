import os
import sys
import time
import subprocess
import shutil
from pathlib import Path
from nbconvert.preprocessors import ExecutePreprocessor
import nbformat
from dotenv import load_dotenv
load_dotenv()

# Setup root directory paths
ROOT_DIR = Path(__file__).resolve().parent
NOTEBOOK_DIR = ROOT_DIR / "notebooks"
CHART_DIR = ROOT_DIR / "data" / "charts"

def run_ocr_batch_file():
    """Triggers the Start_OCR.bat file in the root directory and waits for completion."""
    ocr_path = ROOT_DIR / "Start_OCR.bat"
    print("=" * 60)
    print(f"[STEP 1/4] EXECUTING OCR PROCESSOR: {ocr_path.name}")
    print("=" * 60)
    
    if not ocr_path.exists():
        print(f"[CRITICAL ERROR] Batch file not found at: {ocr_path}")
        sys.exit(1)
        
    try:
        result = subprocess.run(
            str(ocr_path), 
            shell=True, 
            check=True, 
            cwd=str(ROOT_DIR)
        )
        print(f"[SUCCESS] OCR Process finished with exit code {result.returncode}.\n")
    except subprocess.CalledProcessError as e:
        print(f"[CRITICAL ERROR] Start_OCR.bat failed during execution: {e}")
        sys.exit(1)

def run_notebook(notebook_filename, etf_choice="q"):
    """Executes a notebook file programmatically from top to bottom."""
    notebook_path = NOTEBOOK_DIR / notebook_filename
    print(f"[RUNNING] Executing cells in {notebook_filename}...")
    
    start_time = time.time()
    try:
        with open(notebook_path, "r", encoding="utf-8") as f:
            nb = nbformat.read(f, as_version=4)
            
        os.environ["ETF_UPDATE_CHOICE"] = etf_choice
        
        ep = ExecutePreprocessor(timeout=600, kernel_name='python3')
        ep.preprocess(nb, {'metadata': {'path': str(NOTEBOOK_DIR)}})
            
        elapsed = time.time() - start_time
        print(f"[SUCCESS] Finished {notebook_filename} in {elapsed:.2f} seconds.\n")
        
    except Exception as e:
        print(f"[CRITICAL ERROR] Failed executing {notebook_filename}: {e}")
        sys.exit(1)

def open_dashboard_in_vivaldi():
    """Launches Vivaldi directly using the path specified in your .env configuration."""
    dashboard_path = CHART_DIR / "Master_Dashboard.html"
    print("=" * 60)
    print("[STEP 4/4] LAUNCHING INTERACTIVE MASTER DASHBOARD")
    print("=" * 60)
    
    if not dashboard_path.exists():
        print(f"[CRITICAL ERROR] Could not find the dashboard file at: {dashboard_path}")
        sys.exit(1)

    # Grab the exact path string from your environment setup
    config_path = os.getenv("vivaldi_path")
    
    if not config_path:
        print("[CRITICAL ERROR] vivaldi_path is not set in your environment variables (.env).")
        sys.exit(1)

    vivaldi_exe = Path(config_path)
            
    if vivaldi_exe.exists():
        print(f"[LAUNCH] Opening dashboard using Vivaldi: {vivaldi_exe}")
        try:
            # Launch Vivaldi asynchronously so it does not block the terminal
            subprocess.Popen([str(vivaldi_exe), str(dashboard_path)])
            print("[SUCCESS] Dashboard window spawned successfully.")
        except Exception as e:
            print(f"[CRITICAL ERROR] Failed to execute Vivaldi browser binary: {e}")
            sys.exit(1)
    else:
        print(f"[CRITICAL ERROR] Vivaldi executable not found at specified config path: {vivaldi_exe}")
        sys.exit(1)

def main():   
    print("=" * 60)
    print("PIPELINE CONFIGURATION")
    print("=" * 60)
       
    print("Have any new image based pdf files been added from Source 1?")
    print("Type 'y' to run Start_OCR.bat. or Type 'n' to skip OCR processing.")
    run_ocr_choice = input("Your choice (y/n): ").strip().lower()
    print("=" * 60 + "\n")

    # Conditional Step: Trigger the OCR Windows Batch File based on user input   
    print("[STEP 1/4]")
    if run_ocr_choice == 'y':
        run_ocr_batch_file()
    else:
        print("[SKIPPED] No new files from Source 1. Skipping OCR batch execution.\n")
    
    print("[STEP 2/4]")
    run_notebook("01_data_extraction.ipynb")

    # Ask for ETF update choices
    print("Which ETF holdings have you updated? (Enter numbers, e.g., 2,4)")
    print("Type 'q' to skip updates (use existing tables). or Enter 'all' to run all updates.")
    user_choice = input("Your choice: ").strip().lower()
    print("-" * 60)

    print("[STEP 3/4]")
    run_notebook("02_portfolio_tracker.ipynb", user_choice)

    open_dashboard_in_vivaldi()
    
    print("\n" + "=" * 60)
    print("SUCCESS: Pipeline automation loop finalized perfectly.")
    print("=" * 60)

if __name__ == "__main__":
    main()