$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Windows.Forms
try {
    $picker = New-Object System.Windows.Forms.FolderBrowserDialog
    $picker.Description = 'Choose a local folder containing only files approved for Claude. Subfolders are included.'
    if ($picker.ShowDialog() -ne 'OK') { exit }
    $selected = $picker.SelectedPath
    if ($selected.StartsWith('\\')) { throw 'Network folders are not supported in this pilot.' }
    if ([System.Windows.Forms.MessageBox]::Show('Claude will be able to list names in this folder and all subfolders, and import supported workbooks. Select only an approved working folder. Continue?','Approve working folder','YesNo','Warning') -ne 'Yes') { exit }
    @{root=$selected} | ConvertTo-Json | Set-Content -LiteralPath (Join-Path $PSScriptRoot 'approved-folder.json') -Encoding UTF8
    [System.Windows.Forms.MessageBox]::Show('Folder saved. Ask Claude: Browse my approved folder. Source files will not be edited.','Folder selected') | Out-Null
} catch {
    [System.Windows.Forms.MessageBox]::Show('Could not select the folder. Use a local folder and contact IT if needed.','Selection stopped') | Out-Null
}
