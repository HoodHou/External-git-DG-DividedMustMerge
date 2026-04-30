@echo off
setlocal
cd /d "%~dp0"

where python >nul 2>nul
if not errorlevel 1 (
    python app.py
    set "ERR=%errorlevel%"
    if not "%ERR%"=="0" pause
    exit /b %ERR%
)

where py >nul 2>nul
if not errorlevel 1 (
    py -3 app.py
    set "ERR=%errorlevel%"
    if not "%ERR%"=="0" pause
    exit /b %ERR%
)

echo Python not found. Please install Python 3 or add it to PATH.
pause
exit /b 1
