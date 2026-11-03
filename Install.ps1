$ErrorActionPreference = 'Stop'
try {
    $destination = Join-Path $env:LOCALAPPDATA 'claude-safeguard'
    if (Test-Path -LiteralPath $destination) {
        Write-Host 'Existing installation found. Applying the package repair with backups.'
        & (Join-Path $PSScriptRoot 'Repair.ps1')
        exit $LASTEXITCODE
    }
    . (Join-Path $PSScriptRoot 'ClaudeRuntime.ps1')
    Install-SafeguardClaude
    $python = Join-Path $env:LOCALAPPDATA 'Programs\Python\Python313\python.exe'
    if (-not (Test-Path -LiteralPath $python)) {
        if (-not (Get-Command winget -ErrorAction SilentlyContinue)) { throw 'Windows App Installer (winget) is unavailable. IT must enable it for automatic Python setup.' }
        Write-Host 'Installing Python for local document conversion...'
        & winget install --id Python.Python.3.13 --exact --source winget --scope user --accept-source-agreements --accept-package-agreements
        if (-not (Test-Path -LiteralPath $python)) { throw 'Python installation did not complete. Contact IT.' }
    }
    New-Item -ItemType Directory -Path $destination | Out-Null
    foreach ($name in @('guard.ps1','CLAUDE.md','Launch.ps1','ClaudeRuntime.ps1','README.txt','START-HERE.txt','VALIDATION.txt','Prepare.ps1','FileCheck.ps1','CheckFile.ps1','convert.py','spreadsheet.py','finance.py','folder_access.py','sharepoint.py','ConnectSharePoint.ps1','SHAREPOINT-SETUP.txt','SelectFolder.ps1','ImportWorkbook.ps1','requirements.txt','requirements-sharepoint.txt','runtime_check.py','configuration.py')) { Copy-Item -LiteralPath (Join-Path $PSScriptRoot $name) -Destination $destination }
    New-Item -ItemType Directory -Path (Join-Path $destination 'Approved-Data'),(Join-Path $destination '.claude') | Out-Null
    Copy-Item -LiteralPath (Join-Path $PSScriptRoot 'sample.csv') -Destination (Join-Path $destination 'Approved-Data')
    & $python -m venv (Join-Path $destination 'runtime')
    if ($LASTEXITCODE -ne 0) { throw 'Local Python environment creation failed.' }
    New-Item -ItemType Directory -Path (Join-Path $destination 'Workbooks') | Out-Null
    Install-SafeguardDependencies -Destination $destination
    Update-SafeguardConfiguration -Destination $destination -Fresh
    $shell = New-Object -ComObject WScript.Shell
    $shortcut = $shell.CreateShortcut((Join-Path ([Environment]::GetFolderPath('Desktop')) 'claude-safeguard.lnk'))
    $shortcut.TargetPath = Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'
    $shortcut.Arguments = '-NoProfile -ExecutionPolicy Bypass -File "' + (Join-Path $destination 'Launch.ps1') + '" -PauseOnError'
    $shortcut.WorkingDirectory = $destination
    $shortcut.Save()
    $prepareShortcut = $shell.CreateShortcut((Join-Path ([Environment]::GetFolderPath('Desktop')) 'claude-safeguard - Prepare file.lnk'))
    $prepareShortcut.TargetPath = $shortcut.TargetPath
    $prepareShortcut.Arguments = '-NoProfile -STA -ExecutionPolicy Bypass -File "' + (Join-Path $destination 'Prepare.ps1') + '"'
    $prepareShortcut.WorkingDirectory = $destination
    $prepareShortcut.Save()
    $bookShortcut = $shell.CreateShortcut((Join-Path ([Environment]::GetFolderPath('Desktop')) 'claude-safeguard - Import workbook.lnk'))
    $bookShortcut.TargetPath = $shortcut.TargetPath
    $bookShortcut.Arguments = '-NoProfile -STA -ExecutionPolicy Bypass -File "' + (Join-Path $destination 'ImportWorkbook.ps1') + '"'
    $bookShortcut.WorkingDirectory = $destination
    $bookShortcut.Save()
    $folderShortcut = $shell.CreateShortcut((Join-Path ([Environment]::GetFolderPath('Desktop')) 'claude-safeguard - Select working folder.lnk'))
    $folderShortcut.TargetPath = $shortcut.TargetPath
    $folderShortcut.Arguments = '-NoProfile -STA -ExecutionPolicy Bypass -File "' + (Join-Path $destination 'SelectFolder.ps1') + '"'
    $folderShortcut.WorkingDirectory = $destination
    $folderShortcut.Save()
    Write-Host 'Installed. Open claude-safeguard on your desktop. Read START-HERE.txt before using data.'
    Invoke-Item (Join-Path $destination 'START-HERE.txt')
    exit 0
} catch { Write-Host ('Installation stopped: ' + $_.Exception.Message); exit 1 }
