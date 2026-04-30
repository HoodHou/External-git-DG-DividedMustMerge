@echo off
setlocal EnableExtensions
chcp 65001 >nul 2>nul
cd /d "%~dp0"

call "%~dp0build_exe.bat"
set "ERR=%errorlevel%"
if not "%ERR%"=="0" (
    exit /b %ERR%
)

set "RELEASE_DIR=%CD%\release"
set "APP_DIR=%CD%\dist\分久必合"
set "ZIP_PATH=%RELEASE_DIR%\FenJiuBiHe.zip"
set "VERSION_FILE=%APP_DIR%\APP_VERSION.txt"

if not exist "%APP_DIR%\分久必合.exe" (
    echo Built application not found: "%APP_DIR%\分久必合.exe"
    exit /b 1
)

if not exist "%RELEASE_DIR%" mkdir "%RELEASE_DIR%"
if exist "%ZIP_PATH%" del /f /q "%ZIP_PATH%"

powershell -NoProfile -ExecutionPolicy Bypass -Command ^
  "Compress-Archive -Path '%APP_DIR%\*' -DestinationPath '%ZIP_PATH%' -Force"
set "ERR=%errorlevel%"
if not "%ERR%"=="0" (
    echo Failed to create release zip.
    exit /b %ERR%
)

copy /Y "%VERSION_FILE%" "%RELEASE_DIR%\APP_VERSION.txt" >nul

echo.
echo Release assets ready:
echo %ZIP_PATH%
echo %RELEASE_DIR%\APP_VERSION.txt
exit /b 0
