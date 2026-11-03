param([switch]$PauseOnError)
$ErrorActionPreference = 'Stop'
try {
    Set-Location -LiteralPath $PSScriptRoot
    foreach ($file in @('CLAUDE.md','guard.ps1','.claude/settings.json','.mcp.json',
                        'ClaudeRuntime.ps1','runtime/Scripts/python.exe','spreadsheet.py',
                        'finance.py','folder_access.py','runtime_check.py','configuration.py','requirements.txt')) {
        if (-not (Test-Path -LiteralPath (Join-Path $PSScriptRoot $file) -PathType Leaf)) {
            throw "Required installation file missing: $file. Run Install.cmd from the extracted package; if repair stops, ask IT."
        }
    }
    . (Join-Path $PSScriptRoot 'ClaudeRuntime.ps1')
    $claudeExecutable = Find-SafeguardClaude
    if (-not $claudeExecutable) {
        throw 'Claude Code could not be found. Close this window and run Install.cmd from the extracted package to install or repair it.'
    }
    Assert-SafeguardClaude -Executable $claudeExecutable
    Assert-SafeguardRuntime -Destination $PSScriptRoot
    Assert-SafeguardConfiguration -Destination $PSScriptRoot
    Write-Host 'claude-safeguard privacy PILOT: public/synthetic/approved de-identified data only.'
    Write-Host 'Use fresh sessions; do not attach files, use @mentions, enable connectors or disable hooks.'
    Write-Host 'First launch: accept this trusted workspace; use /hooks and /mcp to check the setup.'
    & $claudeExecutable --permission-mode default
    if ($LASTEXITCODE -ne 0) { throw "Claude Code stopped with exit code $LASTEXITCODE. See its message above." }
    exit 0
} catch {
    Write-Host ('Unable to start claude-safeguard: ' + $_.Exception.Message) -ForegroundColor Red
    if ($PauseOnError) { [void](Read-Host 'Press Enter to close') }
    exit 1
}
