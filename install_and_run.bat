@echo off
chcp 65001 >nul
cd /d "%~dp0"
where py >nul 2>nul
if %errorlevel%==0 (
  set PY=py
) else (
  set PY=python
)
if not exist ".venv\Scripts\python.exe" (
  %PY% -m venv .venv || goto :error
)
.venv\Scripts\python.exe -m pip install --upgrade pip
.venv\Scripts\python.exe -m pip install -r requirements.txt || goto :error
.venv\Scripts\python.exe MTR_Obezlichivatel.py
goto :eof
:error
echo.
echo Не удалось подготовить окружение. Проверьте, что Python 3.10+ установлен и доступен в PATH.
pause
