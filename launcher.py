"""
200 WMA Launcher
================
Copy this file into your IDE and run it. It will:
  1. Create a folder on your Desktop
  2. Clone (or update) the repo from GitHub
  3. Install required dependencies
  4. Download the latest tickers
  5. Run the 200 WMA analysis
"""

import os
import sys
import subprocess
from pathlib import Path

REPO_URL = "https://github.com/jkatdare/200-Week-Moving-Average-.git"
FOLDER   = "200-Week-Moving-Average"

# ── Locate the Desktop folder regardless of OS / language ────────────────────
DESKTOP  = Path.home() / "Desktop"
REPO_DIR = DESKTOP / FOLDER


def run(cmd, cwd=None):
    """Run a shell command and stream its output."""
    print(f"\n>>> {' '.join(cmd)}")
    subprocess.run(cmd, cwd=cwd, check=True)


def main():
    # 1. Clone or update
    if (REPO_DIR / ".git").exists():
        print(f"Updating existing repo at {REPO_DIR}")
        run(["git", "pull"], cwd=REPO_DIR)
    else:
        print(f"Cloning fresh repo to {REPO_DIR}")
        DESKTOP.mkdir(exist_ok=True)
        run(["git", "clone", REPO_URL, str(REPO_DIR)])

    # 2. Install dependencies
    print("\nInstalling dependencies...")
    run([sys.executable, "-m", "pip", "install", "--quiet",
         "yfinance", "pandas", "requests", "beautifulsoup4",
         "openpyxl", "matplotlib"])

    # 3. Download tickers
    run([sys.executable, "download_tickers.py"], cwd=REPO_DIR)

    # 4. Run main analysis
    run([sys.executable, "main.py"], cwd=REPO_DIR)

    print("\nAll done!")


if __name__ == "__main__":
    main()
