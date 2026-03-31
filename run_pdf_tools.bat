@echo off
chcp 65001 >nul
cd /d "%~dp0"

python --version >nul 2>&1
if errorlevel 1 (
    echo Python not found. Install Python 3.10+ from python.org and check "Add to PATH".
    pause
    exit /b 1
)

python -c "import fitz" 2>nul
if errorlevel 1 (
    echo Installing PyMuPDF...
    python -m pip install -r "%~dp0requirements.txt"
    if errorlevel 1 (
        echo Install failed. Run: pip install -r requirements.txt
        pause
        exit /b 1
    )
)

python "%~dp0pdf_splitter_app.py"
if errorlevel 1 pause
