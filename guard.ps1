# Local command hook; no network calls and no source-content logging.
# CheckFile is an optional local diagnostic. It uses exactly the Read checks.
param([string]$CheckFile)
$ErrorActionPreference = 'Stop'
$script:eventName = ''
$script:stage = 'INPUT'
$script:localCheck = $PSBoundParameters.ContainsKey('CheckFile')
function Block([string]$code, [string]$message) {
    $reason = "claude-safeguard [$code]: $message"
    if ($script:localCheck) {
        @{status='blocked';code=$code;reason=$reason} | ConvertTo-Json -Compress
        exit 2
    }
    if ($script:eventName -eq 'PreToolUse') {
        @{hookSpecificOutput=@{hookEventName='PreToolUse';permissionDecision='deny';permissionDecisionReason=$reason}} | ConvertTo-Json -Depth 8 -Compress
        exit 0
    }
    if ($script:eventName -eq 'UserPromptSubmit') {
        @{decision='block';reason=$reason;hookSpecificOutput=@{hookEventName='UserPromptSubmit';suppressOriginalPrompt=$true}} | ConvertTo-Json -Depth 8 -Compress
        exit 0
    }
    # SessionStart cannot block a session; a failure is shown to the user.
    [Console]::Error.WriteLine($reason)
    exit 2
}
function CheckText([string]$text) {
    $patterns = @(
        '(?i)[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}',
        '(?<!\w)\+[1-9][0-9 ()-]{9,}[0-9](?!\w)',
        '(?i)\b(?:password|passwd|api[_ -]?key|secret|access[_ -]?token)\s*[:=]',
        '(?i)-----BEGIN .*PRIVATE KEY-----',
        '(?i)\b(?:passport|national insurance|date of birth|full name|email address|phone number|home address|source identity|safeguarding)\b',
        '(?i)\b(?:confidential|restricted|strictly private)\b'
    )
    foreach ($pattern in $patterns) {
        if ([regex]::IsMatch($text,$pattern)) {
            Block 'SG-SENSITIVE' 'A sensitive-data pattern matched. Review the source locally; do not paste it or disable the guard. This can include policy labels, not just personal data.'
        }
    }
}
try {
    [Console]::InputEncoding = New-Object System.Text.UTF8Encoding($false)
    [Console]::OutputEncoding = New-Object System.Text.UTF8Encoding($false)
    $root = [IO.Path]::GetFullPath($PSScriptRoot)
    if ($script:localCheck) {
        $payload = [pscustomobject]@{hook_event_name='PreToolUse';tool_name='Read';cwd=$root;tool_input=@{file_path=$CheckFile}}
    } else {
        $payload = [Console]::In.ReadToEnd() | ConvertFrom-Json
    }
    $script:eventName = [string]$payload.hook_event_name
    if (-not (Test-Path -LiteralPath (Join-Path $root 'CLAUDE.md') -PathType Leaf)) {
        Block 'SG-POLICY' 'The workspace policy is missing. Close Claude and repair the package.'
    }
    switch ($script:eventName) {
        'SessionStart' {
            $script:stage='POLICY'
            @{hookSpecificOutput=@{hookEventName='SessionStart';additionalContext=[IO.File]::ReadAllText((Join-Path $root 'CLAUDE.md'))}} | ConvertTo-Json -Depth 8 -Compress
            exit 0
        }
        'UserPromptSubmit' {
            $script:stage='PROMPT'
            $promptText = [string]$payload.prompt
            if ($promptText.Length -gt 12000) { Block 'SG-PROMPT-SIZE' 'Use a shorter instruction and an approved CSV/TXT file.' }
            if ($promptText -match '@|(?m)^\s*/|[A-Za-z]:\\|\\\\') { Block 'SG-PROMPT-PATH' 'Direct file mentions, slash commands and absolute paths in prompts are restricted. Refer to the exact prepared filename in Approved-Data.' }
            CheckText $promptText
            exit 0
        }
        'PreToolUse' {
            if ($payload.tool_name -eq 'mcp__claude_safeguard__spreadsheet') { exit 0 }
            if ($payload.tool_name -ne 'Read') { Block 'SG-TOOL' 'Only scanned Read and the bundled local tool are permitted. For SharePoint, use its scoped sharepoint_browse/sharepoint_import operations after local setup.' }
            $script:stage='PATH'
            $requested = [string]$payload.tool_input.file_path
            if ([string]::IsNullOrWhiteSpace($requested)) { Block 'SG-PATH' 'No file path was supplied. Use the exact filename shown by Prepare file.' }
            if ($requested.StartsWith('\\') -or $requested.StartsWith('//')) { Block 'SG-PATH' 'Network paths are unsupported. Use Prepare file on a reviewed local copy.' }
            if (-not [IO.Path]::IsPathRooted($requested)) {
                if ([string]::IsNullOrWhiteSpace([string]$payload.cwd)) { Block 'SG-PATH' 'The working directory is missing. Restart using the safeguard shortcut.' }
                $requested = Join-Path ([string]$payload.cwd) $requested
            }
            $full = [IO.Path]::GetFullPath($requested)
            $allowedRoot = [IO.Path]::GetFullPath((Join-Path $root 'Approved-Data'))
            $allowed = $allowedRoot.TrimEnd([IO.Path]::DirectorySeparatorChar, [IO.Path]::AltDirectorySeparatorChar) + [IO.Path]::DirectorySeparatorChar
            if (-not $full.StartsWith($allowed,[StringComparison]::OrdinalIgnoreCase)) { Block 'SG-FOLDER' 'Text must be in the installed safeguard Approved-Data folder. A folder selected for spreadsheet browsing is separate. Use the Prepare file shortcut.' }
            $pathRoot = [IO.Path]::GetPathRoot($full)
            if ($full.Substring($pathRoot.Length).Contains(':')) { Block 'SG-STREAM' 'Alternate data streams are blocked.' }
            if ([IO.Path]::GetExtension($full).ToLowerInvariant() -notin @('.csv','.txt')) { Block 'SG-TYPE' 'Read accepts UTF-8 TXT/CSV. Use Prepare file to convert a supported document; renaming its extension is not conversion.' }
            $script:stage='METADATA'
            try { $item = Get-Item -LiteralPath $full -Force }
            catch [System.Management.Automation.ItemNotFoundException] { Block 'SG-NOT-FOUND' 'No file exists at that exact path. Do not guess other names or extensions. Use Prepare file and its exact prepared filename.' }
            if ($item.PSIsContainer) { Block 'SG-DIRECTORY' 'Choose an individual TXT/CSV file, not a folder.' }
            if ($item.Length -gt 1048576) { Block 'SG-SIZE' 'The text file exceeds 1 MiB. Prepare a smaller reviewed extract.' }
            $node = $item
            while ($null -ne $node) {
                if (($node.Attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0) { Block 'SG-LINK' 'Linked or cloud-placeholder paths are blocked. Prepare a downloaded copy in an ordinary local folder.' }
                if ($node -is [IO.FileInfo]) { $node = $node.Directory } else { $node = $node.Parent }
            }
            $script:stage='READ'
            # Bounded read even if the file grows after metadata inspection.
            $stream = $null
            try {
                $stream = [IO.File]::Open($full,[IO.FileMode]::Open,[IO.FileAccess]::Read,[IO.FileShare]::Read)
                $bytes = New-Object byte[] 1048577
                $count = 0
                while ($count -lt $bytes.Length) {
                    $n = $stream.Read($bytes,$count,$bytes.Length-$count)
                    if ($n -eq 0) { break }
                    $count += $n
                }
            } finally { if ($null -ne $stream) { $stream.Dispose() } }
            if ($count -gt 1048576) { Block 'SG-SIZE' 'The text file exceeds 1 MiB. Prepare a smaller reviewed extract.' }
            $script:stage='ENCODING'
            $encoding = New-Object System.Text.UTF8Encoding($false,$true)
            $content = $encoding.GetString($bytes,0,$count)
            if ($content.Contains([char]0)) { Block 'SG-ENCODING' 'Binary data or an unsupported text encoding was found. Use Prepare file or export as UTF-8 TXT/CSV.' }
            $script:stage='SCAN'
            CheckText $content
            if ($script:localCheck) {
                @{status='passed';code='SG-PASS';reason='Local file checks passed. Nothing was uploaded. Review and approval are still required; pattern checks can miss sensitive data.'} | ConvertTo-Json -Compress
            }
            exit 0
        }
        default { Block 'SG-EVENT' 'Unsupported or missing hook event. Close Claude and report this code to your data lead.' }
    }
} catch {
    # Only fixed error categories leave this process, never exception text or paths.
    $exception = $_.Exception
    $types = @()
    while ($null -ne $exception) { $types += $exception.GetType().FullName; $exception = $exception.InnerException }
    if ($types -contains 'System.Text.DecoderFallbackException') { Block 'SG-ENCODING' 'The file is not valid UTF-8. Use Prepare file or export as UTF-8 TXT/CSV.' }
    if ($types -contains 'System.IO.FileNotFoundException' -or $types -contains 'System.IO.DirectoryNotFoundException') { Block 'SG-NOT-FOUND' 'The requested file or folder is missing. Use the exact prepared filename; do not guess paths.' }
    if ($types -contains 'System.UnauthorizedAccessException') { Block 'SG-ACCESS' 'Windows denied local file access. Check permissions locally or contact IT; do not disable the safeguard.' }
    if ($script:stage -eq 'READ' -and $types -contains 'System.IO.IOException') { Block 'SG-READ' 'The file could not be opened. Close its editor or sync operation and check that the local copy is available.' }
    Block ('SG-CHECK-' + $script:stage) 'The local check could not complete. Close Claude and report this code to your data lead. No file content is included in this diagnostic.'
}
