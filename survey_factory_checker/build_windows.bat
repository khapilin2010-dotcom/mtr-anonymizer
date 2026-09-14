@echo off
setlocal
cd /d "%~dp0"
py -m pip install --upgrade pip
py -m pip install -r requirements.txt pyinstaller
py -m PyInstaller --noconfirm --clean --onefile --windowed --name SurveyFactoryChecker --collect-all fitz --collect-all pymupdf SurveyFactoryChecker.py
if errorlevel 1 exit /b 1
copy /Y "..\MTR_База_обезличивания_v14_BASELINE.xlsx" "dist\MTR_База_обезличивания_v14_BASELINE.xlsx" >nul 2>nul
powershell -NoProfile -Command "Compress-Archive -Path 'dist\SurveyFactoryChecker.exe','dist\MTR_База_обезличивания_v14_BASELINE.xlsx' -DestinationPath 'dist\SurveyFactoryChecker_Windows.zip' -Force"
echo READY: dist\SurveyFactoryChecker_Windows.zip
endlocal
