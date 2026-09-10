@echo off
REM Rebuild the offline snapshot (dashboard/data.js) from the workbook.
setlocal
cd /d "%~dp0"
where py >nul 2>&1 && (set PY=py -3) || (set PY=python)
%PY% -c "import openpyxl" 2>nul || %PY% -m pip install -r requirements.txt
%PY% ob_pipeline.py
pause
