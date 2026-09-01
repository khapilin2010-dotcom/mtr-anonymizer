@echo off
chcp 65001 >nul
cd /d "%~dp0"
where py >nul 2>nul
if %errorlevel%==0 (
  set PY=py
) else (
  set PY=python
)
if not exist ".venv_build\Scripts\python.exe" (
  %PY% -m venv .venv_build || goto :error
)
.venv_build\Scripts\python.exe -m pip install --upgrade pip
.venv_build\Scripts\python.exe -m pip install -r requirements.txt pyinstaller || goto :error
.venv_build\Scripts\python.exe -m PyInstaller --noconfirm --clean --onefile --windowed ^
  --name MTR_Obezlichivatel ^
  --add-data "mtr_data.json.gz;." ^
  --add-data "tessdata;tessdata" ^
  --collect-all pymupdf ^
  MTR_Obezlichivatel.py || goto :error
echo.
echo ГОТОВО: dist\MTR_Obezlichivatel.exe
explorer dist
goto :eof
:error
echo.
echo Ошибка сборки.
pause
