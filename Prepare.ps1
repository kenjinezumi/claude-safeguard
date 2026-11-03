$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Windows.Forms
$tempText = $null
$preparedPath = $null
$preparedOK = $false
$stage = 'SG-PREPARE-UI'
try {
    . (Join-Path $PSScriptRoot 'FileCheck.ps1')
    $dialog = New-Object System.Windows.Forms.OpenFileDialog
    $dialog.Title = 'Choose a locally approved file to extract and review'
    $dialog.Filter = 'Supported documents|*.pdf;*.docx;*.xlsx;*.pptx;*.odt;*.ods;*.odp;*.txt;*.csv;*.tsv;*.json;*.xml;*.html;*.htm;*.md;*.log|All files|*.*'
    if ($dialog.ShowDialog() -ne 'OK') { exit 0 }
    $tempText = Join-Path ([IO.Path]::GetTempPath()) (([guid]::NewGuid().ToString()) + '.txt')
    $stage = 'SG-PREPARE-CONVERT'
    & (Join-Path $PSScriptRoot 'runtime\Scripts\python.exe') (Join-Path $PSScriptRoot 'convert.py') $dialog.FileName $tempText
    if ($LASTEXITCODE -ne 0) { throw 'Conversion stopped. Unsupported, encrypted, oversized, empty or unreadable content. Use a clean export or ask your data lead.' }
    $stage = 'SG-PREPARE-REVIEW'
    Write-Host 'Review this extraction locally. Remove identifiers and sensitive details. Save and CLOSE Notepad.'
    Start-Process notepad.exe -ArgumentList ('"' + $tempText + '"') -Wait
    $answer = [System.Windows.Forms.MessageBox]::Show('Have you reviewed the extracted text and confirmed it is approved for sending to Claude? A pattern scan is not anonymisation.','Approve prepared copy','YesNo','Warning')
    if ($answer -ne 'Yes') { exit 0 }
    $name = 'prepared-' + [guid]::NewGuid().ToString('N').Substring(0,8) + '.txt'
    $stage = 'SG-PREPARE-SAVE'
    $preparedPath = Join-Path (Join-Path $PSScriptRoot 'Approved-Data') $name
    Copy-Item -LiteralPath $tempText -Destination $preparedPath
    $result = Invoke-SafeguardFileCheck -Path $preparedPath
    if ($result.status -ne 'passed') {
        [void][System.Windows.Forms.MessageBox]::Show($result.reason,'Preparation stopped')
        exit 1
    }
    $preparedOK = $true
    [System.Windows.Forms.MessageBox]::Show(('Prepared and checked locally. No content was uploaded. Open claude-safeguard and ask: Summarize ' + $name + ' in Approved-Data. The hook will scan it before reading.'),'Prepared') | Out-Null
} catch {
    [System.Windows.Forms.MessageBox]::Show(('Preparation stopped [' + $stage + ']. No file was sent to Claude. Check the local source format, access and size; report this code if needed.'),'Preparation stopped') | Out-Null
} finally {
    if ($preparedPath -and -not $preparedOK -and (Test-Path -LiteralPath $preparedPath)) { Remove-Item -LiteralPath $preparedPath -Force }
    if ($tempText -and (Test-Path -LiteralPath $tempText)) { Remove-Item -LiteralPath $tempText -Force }
}
