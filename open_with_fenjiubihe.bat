@echo off
setlocal
cd /d "%~dp0"
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0open_with_fenjiubihe.ps1" %*
exit /b %errorlevel%
