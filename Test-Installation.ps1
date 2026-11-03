$ErrorActionPreference = 'Stop'
try {
    $destination = Join-Path $env:LOCALAPPDATA 'claude-safeguard'
    $python = Join-Path $destination 'runtime\Scripts\python.exe'
    if (-not (Test-Path -LiteralPath $python)) { throw 'Run Install.cmd first.' }
    . (Join-Path $destination 'ClaudeRuntime.ps1')
    $claude = Find-SafeguardClaude
    if (-not $claude) { throw 'Claude Code was not found. Run Install.cmd.' }
    Assert-SafeguardClaude -Executable $claude
    Assert-SafeguardConfiguration -Destination $destination
    & $python (Join-Path $destination 'runtime_check.py')
    if ($LASTEXITCODE -ne 0) { throw 'Installed runtime/MCP check failed.' }
    $env:SAFEGUARD_TEST_PWSH = Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'
    & $python -m unittest discover -s (Join-Path $PSScriptRoot 'tests') -v
    if ($LASTEXITCODE -ne 0) { throw 'Package regression tests failed. Keep this window open for IT.' }
    Write-Host 'PASS: installed local runtime and offline package checks. Skipped platform-specific tests are listed above.'
    Write-Host 'This does not test Claude hooks in a live session, Microsoft sign-in, tenant grants or Excel rendering.'
    Write-Host 'Complete the synthetic rehearsal in START-HERE.txt before a demonstration.'
    exit 0
} catch {
    Write-Host ('Check stopped: ' + $_.Exception.Message) -ForegroundColor Red
    exit 1
}
