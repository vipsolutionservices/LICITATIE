param(
    [Parameter(Mandatory=$true)][string]$Alias
)
# Inregistreaza alias-ul agentului curent pentru status line.
# Deriva session_id din cel mai recent .jsonl din ~\.claude\projects\<cwd-encoded>\
# Updateaza .claude\_aliases.json (session_id -> alias).

$ErrorActionPreference = 'Stop'

# Folderul proiect curent encodat (Claude Code foloseste replace '\','-' si ':' -> '-')
$cwd = (Get-Location).Path
$encoded = ($cwd -replace '[\\: ]', '-')
$projectsDir = Join-Path $env:USERPROFILE ".claude\projects\$encoded"
if (-not (Test-Path $projectsDir)) {
    Write-Error "Nu gasesc folderul de proiect Claude: $projectsDir"
    exit 1
}

$latest = Get-ChildItem -Path $projectsDir -Filter '*.jsonl' -File |
          Sort-Object LastWriteTime -Descending |
          Select-Object -First 1
if (-not $latest) {
    Write-Error "Nu gasesc niciun .jsonl in $projectsDir"
    exit 1
}
$sessionId = $latest.BaseName

# Update _aliases.json
$mapFile = Join-Path $PSScriptRoot '_aliases.json'
$map = @{}
if (Test-Path $mapFile) {
    try {
        $existing = Get-Content $mapFile -Raw | ConvertFrom-Json
        foreach ($p in $existing.PSObject.Properties) { $map[$p.Name] = $p.Value }
    } catch {}
}
$map[$sessionId] = $Alias
$map | ConvertTo-Json -Depth 5 | Set-Content -Path $mapFile -Encoding UTF8

Write-Output "Registered: session=$sessionId -> alias=$Alias"
Write-Output "Map file: $mapFile"
