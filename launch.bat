@echo off
setlocal

set REPO_URL=https://github.com/jkatdare/200-Week-Moving-Average-.git
set REPO_DIR=%HOMEDRIVE%%HOMEPATH%\Desktop\200-Week-Moving-Average

:: ── Clone or pull ─────────────────────────────────────────────────────────────
if exist "%REPO_DIR%\.git" (
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
python "%REPO_DIR%\download_tickers.py"

:: ── Run main ──────────────────────────────────────────────────────────────────
echo Running 200 WMA analysis...
python "%REPO_DIR%\main.py"

echo.
echo All done!
pause
