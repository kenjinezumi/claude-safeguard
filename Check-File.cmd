@echo off
setlocal
if not exist "%LOCALAPPDATA%\claude-safeguard\CheckFile.ps1" (
    echo Extract the new ZIP and run Install.cmd or Repair.cmd first.
    pause
    exit /b 1
)
"%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe" -NoProfile -STA -ExecutionPolicy Bypass -File "%LOCALAPPDATA%\claude-safeguard\CheckFile.ps1"
set "safeguard_exit=%errorlevel%"
if not "%safeguard_exit%"=="0" pause
exit /b %safeguard_exit%
