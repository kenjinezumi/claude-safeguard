@echo off
setlocal
"%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe" -NoProfile -ExecutionPolicy Bypass -File "%~dp0Repair.ps1"
set "safeguard_exit=%errorlevel%"
pause
exit /b %safeguard_exit%
