@echo off
REM Freeze the filled live block into history and lay a fresh one underneath.
REM Run this AFTER you have saved the workbook in Excel, and BEFORE you paste
REM the next export over the top of the last one.
setlocal
cd /d "%~dp0"
where py >nul 2>&1 && (set PY=py -3) || (set PY=python)
echo.
echo This is what would happen:
%PY% roll_forward.py --dry-run
echo.
set /p GO="Go ahead? (y/N) "
if /i not "%GO%"=="y" goto :done
%PY% roll_forward.py
:done
pause
