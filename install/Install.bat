@echo off
rem Double-click installer for QuickPhrase.
rem Runs Install.ps1 with the execution policy bypassed for this process only.
setlocal
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0Install.ps1" %*
echo.
pause
