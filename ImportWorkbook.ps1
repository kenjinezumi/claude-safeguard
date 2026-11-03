$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Windows.Forms
$dest = $null
$approved = $false
try {
    $dialog = New-Object System.Windows.Forms.OpenFileDialog
    $dialog.Title = 'Select a reviewed, approved de-identified workbook'
    $dialog.Filter = 'Excel workbook|*.xlsx'
    if ($dialog.ShowDialog() -ne 'OK') { exit }
    if ((Get-Item -LiteralPath $dialog.FileName).Length -gt 25MB) { throw 'Too large' }
    if ([System.Windows.Forms.MessageBox]::Show('Have you reviewed ALL sheets, including hidden sheets, comments and metadata, and confirmed this copy may be processed by Claude? The scanner cannot find every identifier.','Approve workbook','YesNo','Warning') -ne 'Yes') { exit }
    $name = [guid]::NewGuid().ToString('N') + '.xlsx'
    $dest = Join-Path (Join-Path $PSScriptRoot 'Workbooks') $name
    Copy-Item -LiteralPath $dialog.FileName -Destination $dest
    & (Join-Path $PSScriptRoot 'runtime\Scripts\python.exe') (Join-Path $PSScriptRoot 'spreadsheet.py') --check $name
    if ($LASTEXITCODE -ne 0) { Remove-Item -LiteralPath $dest -Force; throw 'Scan failed' }
    $approved = $true
    [System.Windows.Forms.MessageBox]::Show('Copy imported. In Claude, ask: List my workbooks using the spreadsheet tool. Your original is untouched.','Imported') | Out-Null
    Invoke-Item (Join-Path $PSScriptRoot 'Workbooks')
} catch {
    if ($dest -and -not $approved) { Remove-Item -LiteralPath $dest -Force -ErrorAction SilentlyContinue }
    [System.Windows.Forms.MessageBox]::Show('Import stopped. Review the workbook locally for sensitive content, unsupported features or size limits.','Import stopped') | Out-Null
}
