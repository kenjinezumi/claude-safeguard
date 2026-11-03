@echo off
setlocal
if not exist "%LOCALAPPDATA%\claude-safeguard\Launch.ps1" (
    echo claude-safeguard is not installed. Extract the whole ZIP and run Install.cmd first.
    pause
    exit /b 1
)
"%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe" -NoProfile -ExecutionPolicy Bypass -File "%LOCALAPPDATA%\claude-safeguard\Launch.ps1"
set "safeguard_exit=%errorlevel%"
if not "%safeguard_exit%"=="0" pause
exit /b %safeguard_exit%
