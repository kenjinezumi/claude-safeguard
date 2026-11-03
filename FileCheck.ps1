# Local UI helper. Starts the same guard in diagnostic mode; no Claude process.
function Invoke-SafeguardFileCheck {
    param([Parameter(Mandatory=$true)][string]$Path)
    $guardPath = Join-Path $PSScriptRoot 'guard.ps1'
    $powershell = Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'
    $info = New-Object System.Diagnostics.ProcessStartInfo
    $info.FileName = $powershell
    $info.Arguments = '-NoProfile -NonInteractive -ExecutionPolicy Bypass -File "' + $guardPath + '" -CheckFile "' + $Path + '"'
    $info.UseShellExecute = $false
    $info.CreateNoWindow = $true
    $info.RedirectStandardOutput = $true
    $info.RedirectStandardError = $true
    $info.StandardOutputEncoding = New-Object System.Text.UTF8Encoding($false)
    $process = New-Object System.Diagnostics.Process
    $process.StartInfo = $info
    try {
        [void]$process.Start()
        $stdoutTask = $process.StandardOutput.ReadToEndAsync()
        $stderrTask = $process.StandardError.ReadToEndAsync()
        if (-not $process.WaitForExit(20000)) {
            $process.Kill()
            throw 'Local check timed out.'
        }
        $result = $stdoutTask.GetAwaiter().GetResult() | ConvertFrom-Json
        [void]$stderrTask.GetAwaiter().GetResult()
        if ($result.status -eq 'passed' -and $process.ExitCode -eq 0) { return $result }
        if ($result.status -eq 'blocked' -and $process.ExitCode -eq 2) { return $result }
        throw 'Unexpected local check response.'
    } catch {
        return [pscustomobject]@{status='blocked';code='SG-DIAGNOSTIC';reason='The local file check could not run. Close Claude and report SG-DIAGNOSTIC to your data lead.'}
    } finally { $process.Dispose() }
}
