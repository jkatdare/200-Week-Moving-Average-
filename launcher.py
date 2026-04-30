
import io
import sys
import shutil
import zipfile
import urllib.request
import subprocess
from pathlib import Path

REPO_ZIP = "https://github.com/jkatdare/200-Week-Moving-Average-/archive/refs/heads/main.zip"
FOLDER   = "200-Week-Moving-Average"

DESKTOP  = Path.home() / "Desktop"
REPO_DIR = DESKTOP / FOLDER


def run(cmd, cwd=None):
    print(f"\n>>> {' '.join(cmd)}")
    subprocess.run(cmd, cwd=cwd, check=True)


def download_and_extract():
    print(f"Downloading repo ZIP from GitHub...")

    
    if REPO_DIR.exists():
        shutil.rmtree(REPO_DIR)

  
    with urllib.request.urlopen(REPO_ZIP) as response:
        zip_data = response.read()

    
    DESKTOP.mkdir(exist_ok=True)
    with zipfile.ZipFile(io.BytesIO(zip_data)) as z:
        z.extractall(DESKTOP)

  
    extracted = next(DESKTOP.glob("200-Week-Moving-Average*-main"))
    extracted.rename(REPO_DIR)
    print(f"Extracted to {REPO_DIR}")


def main():
    download_and_extract()

    print("\nInstalling dependencies...")
    run([sys.executable, "-m", "pip", "install", "--quiet",
         "yfinance", "pandas", "requests", "beautifulsoup4",
         "openpyxl", "matplotlib"])

    run([sys.executable, "download_tickers.py"], cwd=REPO_DIR)
    run([sys.executable, "main.py"], cwd=REPO_DIR)

    print("\nAll done!")


if __name__ == "__main__":
    main()
