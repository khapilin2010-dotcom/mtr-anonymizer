@echo off
chcp 65001 >nul
cd /d "%~dp0"

rem Use the Python selected by the environment/CI. This avoids py.exe silently
rem choosing another installed version (for example 3.14 on GitHub runners).
set "PY=python"

if not exist ".venv_build\Scripts\python.exe" (
  %PY% -m venv .venv_build || goto :error
)

.venv_build\Scripts\python.exe -m pip install --upgrade pip || goto :error
.venv_build\Scripts\python.exe -m pip install -r requirements.txt pyinstaller || goto :error
.venv_build\Scripts\python.exe -m PyInstaller --noconfirm --clean --onefile --windowed ^
  --name MTR_Obezlichivatel ^
  --add-data "mtr_data.json.gz;." ^
  --collect-all pymupdf ^
  MTR_Obezlichivatel.py || goto :error

echo.
echo ГОТОВО: dist\MTR_Obezlichivatel.exe

rem Explorer is useful only for a local interactive build. On GitHub Actions it
rem returns a non-zero code and falsely marks a successful EXE build as failed.
if defined GITHUB_ACTIONS exit /b 0
explorer dist >nul 2>nul
exit /b 0

:error
echo.
echo Ошибка сборки.
exit /b 1
