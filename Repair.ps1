$ErrorActionPreference = 'Stop'
try {
    $destination = Join-Path $env:LOCALAPPDATA 'claude-safeguard'
    if ([IO.Path]::GetFullPath($PSScriptRoot).TrimEnd('\') -eq [IO.Path]::GetFullPath($destination).TrimEnd('\')) {
        throw 'Run Repair.cmd from the newly downloaded and extracted ZIP, outside the installed folder.'
    }
    # Back up package files and settings; preserve user data and unrelated settings.
    foreach ($file in @('guard.ps1','CLAUDE.md','.claude/settings.json','.mcp.json',
                        'runtime/Scripts/python.exe','spreadsheet.py','finance.py',
                        'folder_access.py','Prepare.ps1','SelectFolder.ps1','ImportWorkbook.ps1')) {
        if (-not (Test-Path -LiteralPath (Join-Path $destination $file) -PathType Leaf)) {
            throw "Existing installation is incomplete (missing $file). Ask IT to complete setup while preserving Workbooks, Approved-Data and approved-folder.json."
        }
    }
    $updates = @('Launch.ps1','ClaudeRuntime.ps1','guard.ps1','Prepare.ps1','FileCheck.ps1','CheckFile.ps1','spreadsheet.py','finance.py','folder_access.py','convert.py','ImportWorkbook.ps1','sharepoint.py','ConnectSharePoint.ps1','SHAREPOINT-SETUP.txt','CLAUDE.md','START-HERE.txt','README.txt','VALIDATION.txt','requirements.txt','requirements-sharepoint.txt','runtime_check.py','configuration.py')
    foreach ($name in $updates) {
        if (-not (Test-Path -LiteralPath (Join-Path $PSScriptRoot $name) -PathType Leaf)) {
            throw 'Repair files missing. Extract the whole new ZIP and run Repair.cmd again.'
        }
    }
    . (Join-Path $PSScriptRoot 'ClaudeRuntime.ps1')
    Install-SafeguardClaude
    $backup = Join-Path $destination ('Launcher-Backups\' + (Get-Date -Format 'yyyyMMdd-HHmmss') + '-' + [guid]::NewGuid().ToString('N'))
    New-Item -ItemType Directory -Path $backup -Force | Out-Null
    $oldFiles = @()
    $backupFiles = $updates + @('.claude/settings.json','.mcp.json')
    foreach ($name in $backupFiles) {
        $existing = Join-Path $destination $name
        if (Test-Path -LiteralPath $existing -PathType Leaf) {
            New-Item -ItemType Directory -Path (Split-Path -Parent (Join-Path $backup $name)) -Force | Out-Null
            Copy-Item -LiteralPath $existing -Destination (Join-Path $backup $name)
            $oldFiles += $name
        }
    }
    try {
        foreach ($name in $updates) {
            Copy-Item -LiteralPath (Join-Path $PSScriptRoot $name) -Destination (Join-Path $destination $name) -Force
        }
        Install-SafeguardDependencies -Destination $destination
        Update-SafeguardConfiguration -Destination $destination
        Set-SafeguardLaunchShortcut -Destination $destination
    } catch {
        # File/settings rollback does not undo pip changes; rerun Repair.cmd after fixing the error.
        foreach ($name in $backupFiles) {
            if ($oldFiles -contains $name) {
                Copy-Item -LiteralPath (Join-Path $backup $name) -Destination (Join-Path $destination $name) -Force
            } else {
                Remove-Item -LiteralPath (Join-Path $destination $name) -Force -ErrorAction SilentlyContinue
            }
        }
        throw
    }
    Write-Host 'Package repaired. Open claude-safeguard on your desktop, or double-click Start.cmd here.'
    Write-Host 'Your workbooks and approved folder were preserved. Guard settings and dependencies were updated. Previous files/settings are in Launcher-Backups.'
    exit 0
} catch {
    Write-Host ('Repair stopped: ' + $_.Exception.Message) -ForegroundColor Red
    exit 1
}
