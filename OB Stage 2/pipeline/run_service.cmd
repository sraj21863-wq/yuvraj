@echo off
REM Start the Stage 2 data service and open the dashboard.
setlocal
cd /d "%~dp0"
where py >nul 2>&1 && (set PY=py -3) || (set PY=python)
%PY% -c "import openpyxl" 2>nul || %PY% -m pip install -r requirements.txt
start "" http://127.0.0.1:8787/
%PY% serve.py
pause
