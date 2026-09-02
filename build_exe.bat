@echo off
chcp 65001 >nul
cd /d "%~dp0"
if defined CI (
  set PY=python
) else (
  where py >nul 2>nul
  if errorlevel 1 (set PY=python) else (set PY=py)
)
if not exist ".venv_build\Scripts\python.exe" (
  %PY% -m venv .venv_build || goto :error
)
.venv_build\Scripts\python.exe -m pip install --upgrade pip
.venv_build\Scripts\python.exe -m pip install -r requirements.txt pyinstaller || goto :error
.venv_build\Scripts\python.exe -m PyInstaller --noconfirm --clean --onefile --windowed ^
  --name MTR_Obezlichivatel ^
  --add-data "mtr_data.json.gz;." ^
  --collect-all pymupdf ^
  MTR_Obezlichivatel.py || goto :error
echo.
echo ГОТОВО: dist\MTR_Obezlichivatel.exe
if not defined CI explorer dist
exit /b 0
:error
echo.
echo Ошибка сборки.
if not defined CI pause
exit /b 1
