# Shared by setup, repair and launch. No persistent PATH or profile changes.
function Update-SafeguardProcessPath {
    $parts = @($env:PATH, [Environment]::GetEnvironmentVariable('PATH', 'User'),
        [Environment]::GetEnvironmentVariable('PATH', 'Machine'))
    # Preserve process-only entries as well as newer persisted entries.
    $env:PATH = ($parts | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }) -join [IO.Path]::PathSeparator
}

function Find-SafeguardClaude {
    if (-not [string]::IsNullOrWhiteSpace($env:USERPROFILE)) {
        $nativePath = Join-Path $env:USERPROFILE '.local\bin\claude.exe'
        if (Test-Path -LiteralPath $nativePath -PathType Leaf) {
            return (Get-Item -LiteralPath $nativePath).FullName
        }
    }
    Update-SafeguardProcessPath
    # Application excludes aliases/functions and PowerShell profile shims.
    foreach ($name in @('claude.exe', 'claude.cmd')) {
        $commands = @(Get-Command $name -CommandType Application -ErrorAction SilentlyContinue)
        foreach ($candidate in $commands) {
            if (Test-Path -LiteralPath $candidate.Path -PathType Leaf) { return $candidate.Path }
        }
    }
    return $null
}

function Assert-SafeguardClaude {
    param([Parameter(Mandatory=$true)][string]$Executable, [switch]$Update)
    $versionText = (& $Executable --version | Out-String)
    if ($LASTEXITCODE -ne 0) {
        throw "Claude Code version check failed (exit code $LASTEXITCODE). Ask IT to repair Claude Code."
    }
    if ($versionText -notmatch '(?<![0-9])([0-9]+\.[0-9]+\.[0-9]+)') { throw 'Claude Code returned an unrecognised version.' }
    $version = [version]$Matches[1]
    if ($version -lt [version]'2.1.295') {
        if ($Update) {
            Write-Host 'Updating Claude Code to support blocking failed safeguards...'
            & $Executable update
            if ($LASTEXITCODE -ne 0) { throw 'Claude Code update failed. Ask IT to update to v2.1.295 or later.' }
            Assert-SafeguardClaude -Executable $Executable
            return
        }
        throw 'Claude Code v2.1.295 or later is required. Close Claude and run Install.cmd to update it.'
    }
    Write-Host ('Claude Code version checked: ' + $version)
}

function Install-SafeguardClaude {
    $executable = Find-SafeguardClaude
    if (-not $executable) {
        Write-Host 'Installing Claude Code from Anthropic...'
        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
        $installer = Join-Path ([IO.Path]::GetTempPath()) (([guid]::NewGuid().ToString()) + '.ps1')
        $powershell = Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'
        try {
            Invoke-WebRequest -UseBasicParsing -Uri 'https://claude.ai/install.ps1' -OutFile $installer
            & $powershell -NoProfile -ExecutionPolicy Bypass -File $installer
            if ($LASTEXITCODE -ne 0) { throw 'Official Claude Code installation failed. See the error above or contact IT.' }
        } finally { Remove-Item -LiteralPath $installer -Force -ErrorAction SilentlyContinue }
        $executable = Find-SafeguardClaude
        if (-not $executable) {
            throw 'Claude Code executable is still missing after installation. See the installer output above or contact IT.'
        }
    }
    # A successful download or installer exit is not sufficient: run the binary.
    Assert-SafeguardClaude -Executable $executable -Update
}

function Set-SafeguardLaunchShortcut {
    param([Parameter(Mandatory=$true)][string]$Destination)
    $shell = New-Object -ComObject WScript.Shell
    $shortcut = $shell.CreateShortcut((Join-Path ([Environment]::GetFolderPath('Desktop')) 'claude-safeguard.lnk'))
    $shortcut.TargetPath = Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'
    $shortcut.Arguments = '-NoProfile -ExecutionPolicy Bypass -File "' + (Join-Path $Destination 'Launch.ps1') + '" -PauseOnError'
    $shortcut.WorkingDirectory = $Destination
    $shortcut.Save()
}

function Install-SafeguardDependencies {
    param([Parameter(Mandatory=$true)][string]$Destination)
    $python = Join-Path $Destination 'runtime\Scripts\python.exe'
    & $python -m pip install --disable-pip-version-check -r (Join-Path $Destination 'requirements.txt')
    if ($LASTEXITCODE -ne 0) { throw 'Dependency installation failed. Check the error above with IT and rerun Repair.cmd.' }
    & $python (Join-Path $Destination 'runtime_check.py')
    if ($LASTEXITCODE -ne 0) { throw 'Local runtime/MCP check failed. Setup has not completed.' }
}

function Assert-SafeguardRuntime {
    param([Parameter(Mandatory=$true)][string]$Destination)
    & (Join-Path $Destination 'runtime\Scripts\python.exe') (Join-Path $Destination 'runtime_check.py') --imports-only
    if ($LASTEXITCODE -ne 0) { throw 'Local runtime is incomplete. Close Claude and run Repair.cmd.' }
}

function Update-SafeguardConfiguration {
    param([Parameter(Mandatory=$true)][string]$Destination, [switch]$Fresh)
    $mode = '--repair'
    if ($Fresh) { $mode = '--install' }
    $powershell = Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'
    & (Join-Path $Destination 'runtime\Scripts\python.exe') (Join-Path $Destination 'configuration.py') $mode $powershell
    if ($LASTEXITCODE -ne 0) { throw 'Settings update failed. Ask IT to review the existing settings.' }
}

function Assert-SafeguardConfiguration {
    param([Parameter(Mandatory=$true)][string]$Destination)
    $powershell = Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'
    & (Join-Path $Destination 'runtime\Scripts\python.exe') (Join-Path $Destination 'configuration.py') --check $powershell
    if ($LASTEXITCODE -ne 0) { throw 'Required safeguard settings are missing or changed. Close Claude and run Repair.cmd.' }
}
