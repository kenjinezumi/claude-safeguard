$ErrorActionPreference = 'Stop'
try {
    Add-Type -AssemblyName System.Windows.Forms
    $python = Join-Path $PSScriptRoot 'runtime\Scripts\python.exe'
    if (-not (Test-Path -LiteralPath $python -PathType Leaf)) { throw 'Run Install.cmd or Repair.cmd from the new ZIP first.' }
    Write-Host 'Preparing Microsoft sign-in components...'
    & $python -m pip install --disable-pip-version-check -r (Join-Path $PSScriptRoot 'requirements-sharepoint.txt')
    if ($LASTEXITCODE -ne 0) { throw 'Microsoft sign-in dependencies could not be installed. Check connectivity with IT.' }
    & $python (Join-Path $PSScriptRoot 'runtime_check.py') --sharepoint --imports-only
    if ($LASTEXITCODE -ne 0) { throw 'Runtime check failed. Close Claude and run Repair.cmd.' }
    $config = Join-Path $env:LOCALAPPDATA 'claude-safeguard-auth\sharepoint-settings.json'
    $choose = $true
    if (Test-Path -LiteralPath $config -PathType Leaf) {
        $answer = [System.Windows.Forms.MessageBox]::Show('Use your saved SharePoint settings? Choose No to select a new settings file from IT.','SharePoint settings','YesNoCancel','Question')
        if ($answer -eq 'Cancel') { exit 0 }
        $choose = $answer -eq 'No'
    }
    if ($choose) {
        $dialog = New-Object System.Windows.Forms.OpenFileDialog
        $dialog.Title = 'Select sharepoint-settings.json supplied by your Microsoft 365 administrator'
        $dialog.Filter = 'SharePoint settings|*.json'
        if ($dialog.ShowDialog() -ne 'OK') { exit 0 }
        $config = $dialog.FileName
    }
    $answer = [System.Windows.Forms.MessageBox]::Show('The configured SharePoint folder must contain only material approved for Claude. Claude will see scanned filenames and imported content. Pattern checks can miss sensitive data. Continue with these IT-approved settings?','Approve SharePoint source','YesNo','Warning')
    if ($answer -ne 'Yes') { exit 0 }
    if ($choose) {
        & $python (Join-Path $PSScriptRoot 'sharepoint.py') --configure $config
        if ($LASTEXITCODE -ne 0) { throw 'Settings were not accepted. Ask IT to check the settings file.' }
    }
    Write-Host 'Sign in in the Microsoft browser window. Never paste passwords or tokens into Claude.'
    & $python (Join-Path $PSScriptRoot 'sharepoint.py') --sign-in
    if ($LASTEXITCODE -ne 0) { throw 'SharePoint connection stopped. Use the SP- code above when asking IT for help.' }
    exit 0
} catch {
    Write-Host ('SharePoint setup stopped: ' + $_.Exception.Message) -ForegroundColor Red
    exit 1
}
