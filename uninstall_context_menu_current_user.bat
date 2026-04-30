@echo off
setlocal
cd /d "%~dp0"
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0uninstall_context_menu_current_user.ps1"
pause
exit /b %errorlevel%
