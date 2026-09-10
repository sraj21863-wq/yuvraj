@echo off
REM ---------------------------------------------------------------------------
REM  Start the data service and open the dashboard.
REM ---------------------------------------------------------------------------
setlocal
cd /d "%~dp0"

REM Find a working Python. On Windows 10/11 a bare "python" is often a Microsoft
REM Store stub that is not Python at all, so test that it actually runs before
REM trusting it -- otherwise the Store opens and nothing explains why.
set PY=
where py >nul 2>&1 && set PY=py -3
if not defined PY (
  where python >nul 2>&1 && set PY=python
)
if defined PY (
  %PY% -c "import sys" >nul 2>&1 || set PY=
)
if not defined PY goto nopython

%PY% -c "import openpyxl" >nul 2>&1 || (
  echo Installing the one library this needs, please wait...
  %PY% -m pip install --quiet -r requirements.txt || goto pipfailed
)

start "" http://127.0.0.1:8787/
%PY% serve.py
goto done

:nopython
echo.
echo   Python is not installed on this computer, or Windows cannot find it.
echo.
echo   The quickest fix, in a PowerShell window:
echo.
echo       winget install -e --id Python.Python.3.12
echo.
echo   Or download it from  https://www.python.org/downloads/
echo   and tick "Add python.exe to PATH" on the installer's FIRST screen.
echo   That tick box is the whole reason this message appears.
echo.
echo   Close this window, install Python, then run this file again.
echo.
echo   You do not need Python just to LOOK at the dashboard -- open
echo   dashboard\Order_Booking_Dashboard_standalone.html in any browser.
echo.
pause
goto :eof

:pipfailed
echo.
echo   Could not install openpyxl. If this machine is behind a proxy, run:
echo       %PY% -m pip install openpyxl
echo   and read the error it prints.
echo.
pause
goto :eof

:done
pause
