# Installs or updates outlook-classic-mcp for the current user.
#
#   irm https://raw.githubusercontent.com/jli-software/outlook-classic-mcp/main/install.ps1 | iex
#
# Downloads the latest release, verifies its SHA-256 and registers the server
# in Claude Desktop, Claude Code and Codex. No admin rights needed.

$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'

$repo  = 'jli-software/outlook-classic-mcp'
$asset = 'outlook-classic-mcp-windows-x64.exe'
$dir   = Join-Path $env:LOCALAPPDATA 'Programs\outlook-classic-mcp'
$exe   = Join-Path $dir 'outlook-classic-mcp.exe'

[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

Write-Host "outlook-classic-mcp installer"
$release = Invoke-RestMethod "https://api.github.com/repos/$repo/releases/latest" -Headers @{ 'User-Agent' = 'outlook-classic-mcp-installer' }
$exeUrl  = ($release.assets | Where-Object name -eq $asset).browser_download_url
$sumUrl  = ($release.assets | Where-Object name -eq 'SHA256SUMS').browser_download_url
if (-not $exeUrl -or -not $sumUrl) { throw "Release $($release.tag_name) has no $asset or SHA256SUMS." }
Write-Host "  version  $($release.tag_name)"

New-Item -ItemType Directory -Force -Path $dir | Out-Null
$tmp = Join-Path $dir "$asset.download"
Invoke-WebRequest $exeUrl -OutFile $tmp -UseBasicParsing

# Checksum
$sums = (Invoke-WebRequest $sumUrl -UseBasicParsing).Content
# PowerShell 7 returns application/octet-stream as bytes, 5.1 as text.
if ($sums -is [byte[]]) { $sums = [Text.Encoding]::UTF8.GetString($sums) }
$line = ($sums -split "`n") | Where-Object { $_ -match [regex]::Escape($asset) } | Select-Object -First 1
if (-not $line) { Remove-Item $tmp; throw "No checksum for $asset in SHA256SUMS." }
$want = ($line -split '\s+')[0].ToLower()
$got  = (Get-FileHash $tmp -Algorithm SHA256).Hash.ToLower()
if ($want -ne $got) { Remove-Item $tmp; throw "Checksum mismatch for $asset (expected $want, got $got)." }

# A running MCP client keeps the old EXE open; Windows allows renaming it.
if (Test-Path $exe) {
    $old = "$exe.old"
    Remove-Item $old -Force -ErrorAction SilentlyContinue
    try { Remove-Item $exe -Force } catch { Rename-Item $exe $old -Force }
}
Move-Item $tmp $exe -Force
Unblock-File $exe
Write-Host "  installed $exe"
Write-Host ""

& $exe install
if ($LASTEXITCODE -ne 0) { throw "Registration failed (exit code $LASTEXITCODE)." }
