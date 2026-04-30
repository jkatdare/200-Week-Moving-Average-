@echo off
setlocal

set REPO_URL=https://github.com/jkatdare/200-Week-Moving-Average-.git
set REPO_DIR=%USERPROFILE%\Desktop\200-Week-Moving-Average

:: ── Clone or pull ─────────────────────────────────────────────────────────────
if exist "%REPO_DIR%" (
    echo Updating from GitHub...
    cd /d "%REPO_DIR%"
    git pull
) else (
    echo Cloning repo from GitHub...
    git clone %REPO_URL% "%REPO_DIR%"
    cd /d "%REPO_DIR%"
)

:: ── Install dependencies ───────────────────────────────────────────────────────
echo Installing dependencies...
pip install yfinance pandas requests beautifulsoup4 openpyxl matplotlib --quiet

:: ── Download tickers ──────────────────────────────────────────────────────────
echo Downloading tickers...
python download_tickers.py

:: ── Run main ──────────────────────────────────────────────────────────────────
echo Running 200 WMA analysis...
python main.py

echo.
echo All done!
pause
