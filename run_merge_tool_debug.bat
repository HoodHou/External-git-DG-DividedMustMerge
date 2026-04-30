@echo off
setlocal
cd /d "%~dp0"

where python >nul 2>nul
if not errorlevel 1 (
    python app.py
    pause
    exit /b %errorlevel%
)

where py >nul 2>nul
if not errorlevel 1 (
    py -3 app.py
    pause
    exit /b %errorlevel%
)

echo Python not found. Please install Python 3 or add it to PATH.
pause
exit /b 1
