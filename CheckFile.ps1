$ErrorActionPreference = 'Stop'
try {
    Add-Type -AssemblyName System.Windows.Forms
    . (Join-Path $PSScriptRoot 'FileCheck.ps1')
    $dialog = New-Object System.Windows.Forms.OpenFileDialog
    $dialog.Title = 'Check a prepared local file - nothing is uploaded'
    $dialog.InitialDirectory = Join-Path $PSScriptRoot 'Approved-Data'
    $dialog.Filter = 'Prepared text|*.txt;*.csv|All files|*.*'
    if ($dialog.ShowDialog() -ne 'OK') { exit 0 }
    $result = Invoke-SafeguardFileCheck -Path $dialog.FileName
    [void][System.Windows.Forms.MessageBox]::Show($result.reason,'claude-safeguard local file check')
} catch {
    Write-Host 'Local diagnostic unavailable. Report SG-DIAGNOSTIC to your data lead. No file was uploaded.'
    exit 1
}
