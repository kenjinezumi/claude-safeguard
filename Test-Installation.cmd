@echo off
"%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe" -NoProfile -ExecutionPolicy Bypass -File "%~dp0Test-Installation.ps1"
set "result=%errorlevel%"
pause
exit /b %result%
