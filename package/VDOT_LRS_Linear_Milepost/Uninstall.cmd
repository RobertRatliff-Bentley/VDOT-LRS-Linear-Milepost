@echo off
rem Removes the VDOT LRS Linear Milepost tool for the current user.
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0Install.ps1" -Uninstall %*
echo.
pause
