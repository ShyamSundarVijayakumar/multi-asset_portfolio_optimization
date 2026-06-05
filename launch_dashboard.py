import os
import sys
import time
from pathlib import Path
import nbformat
from nbconvert.preprocessors import ExecutePreprocessor
from dotenv import load_dotenv

# Load environment variables for Vivaldi Path
load_dotenv()

# Setup root directory paths
ROOT_DIR = Path(__file__).resolve().parent
NOTEBOOK_DIR = ROOT_DIR / "notebooks"
CHART_DIR = ROOT_DIR / "data" / "charts"

def run_notebook_fast(notebook_filename):
    """Executes a notebook in the background, forcing 'q' for no ETF updates."""
    notebook_path = NOTEBOOK_DIR / notebook_filename
    print(f"[FAST RUN] Refreshing dashboard components via {notebook_filename}...")
    
    start_time = time.time()
    try:
        with open(notebook_path, "r", encoding="utf-8") as f:
            nb = nbformat.read(f, as_version=4)
        
        # Force "q" so it completely skips manual prompts and downloads
        os.environ["ETF_UPDATE_CHOICE"] = "q"
        
        ep = ExecutePreprocessor(timeout=300, kernel_name='python3')
        ep.preprocess(nb, {'metadata': {'path': str(NOTEBOOK_DIR)}})
            
        elapsed = time.time() - start_time
        print(f"[SUCCESS] Dashboard data refreshed in {elapsed:.2f} seconds.\n")
        
    except Exception as e:
        print(f"[CRITICAL ERROR] Failed updating dashboard: {e}")
        input("\nPress Enter to exit...")
        sys.exit(1)

def open_dashboard_in_vivaldi():
    """Launches Vivaldi using your .env path configuration."""
    dashboard_path = CHART_DIR / "Master_Dashboard.html"
    
    if not dashboard_path.exists():
        print(f"[CRITICAL ERROR] Dashboard file not found at: {dashboard_path}")
        input("\nPress Enter to exit...")
        sys.exit(1)

    config_path = os.getenv("vivaldi_path")
    if not config_path:
        print("[CRITICAL ERROR] vivaldi_path is not set in your .env file.")
        input("\nPress Enter to exit...")
        sys.exit(1)

    vivaldi_exe = Path(config_path)
    if vivaldi_exe.exists():
        print(f"[LAUNCH] Opening Master Dashboard in Vivaldi...")
        import subprocess
        subprocess.Popen([str(vivaldi_exe), str(dashboard_path)])
    else:
        print(f"[CRITICAL ERROR] Vivaldi executable not found at: {vivaldi_exe}")
        input("\nPress Enter to exit...")
        sys.exit(1)

def main():
    print("=" * 60)
    print("   FAST LAUNCH: REFRESHING & OPENING DASHBOARD   ")
    print("=" * 60)
    
    run_notebook_fast("02_portfolio_tracker.ipynb")
    open_dashboard_in_vivaldi()
    
    print("[FINISHED] Dashboard launched successfully.")
    time.sleep(2)
    
if __name__ == "__main__":
    main()