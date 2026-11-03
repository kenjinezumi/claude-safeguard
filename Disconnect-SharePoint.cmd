@echo off
setlocal
if not exist "%LOCALAPPDATA%\claude-safeguard\ConnectSharePoint.ps1" (
    echo Extract the new ZIP and run Install.cmd or Repair.cmd first.
    pause
    exit /b 1
)
"%LOCALAPPDATA%\claude-safeguard\runtime\Scripts\python.exe" "%LOCALAPPDATA%\claude-safeguard\sharepoint.py" --sign-out
set "safeguard_exit=%errorlevel%"
pause
exit /b %safeguard_exit%
