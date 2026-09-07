@echo off
setlocal
cd /d "%~dp0.."
python -m PyInstaller --noconfirm --clean --onefile --windowed --name MTR_Excel --paths . --add-data "mtr_data.json.gz;." --exclude-module fitz --exclude-module pymupdf --exclude-module mtr_core --exclude-module MTR_Obezlichivatel excel\MTR_Excel.py
exit /b %errorlevel%
