@echo off
setlocal EnableExtensions
chcp 65001 >nul 2>nul
cd /d "%~dp0"

set "PYTHON_EXE="
set "PYTHON_ARGS="

where python >nul 2>nul
if not errorlevel 1 (
    set "PYTHON_EXE=python"
)

if not defined PYTHON_EXE (
    where py >nul 2>nul
    if not errorlevel 1 (
        set "PYTHON_EXE=py"
        set "PYTHON_ARGS=-3"
    )
)

if not defined PYTHON_EXE (
    echo Python not found. Please install Python 3 or add it to PATH.
    pause
    exit /b 1
)

echo Building FenJiuBiHe...
"%PYTHON_EXE%" %PYTHON_ARGS% -m PyInstaller --noconfirm --clean xml_merge_tool.spec
set "ERR=%errorlevel%"

if not "%ERR%"=="0" (
    echo.
    echo Build failed with exit code %ERR%.
    pause
    exit /b %ERR%
)

"%PYTHON_EXE%" %PYTHON_ARGS% -c "from pathlib import Path; from table_merge_tool.version import APP_VERSION; targets=list(Path('dist').glob('*/*.exe')); assert targets, 'built exe not found'; [((exe.parent/'APP_VERSION.txt').write_text(APP_VERSION, encoding='utf-8')) for exe in targets]; print('APP_VERSION=' + APP_VERSION)"
set "ERR=%errorlevel%"

if not "%ERR%"=="0" (
    echo.
    echo Failed to write APP_VERSION.txt with exit code %ERR%.
    pause
    exit /b %ERR%
)

echo.
echo Build completed:
for /r "%CD%\dist" %%F in (*.exe) do echo %%F
for /f "delims=" %%F in ('dir /b /s "%CD%\dist\APP_VERSION.txt" 2^>nul') do echo %%F
exit /b 0
