# Read-only: open a JSON with .NET UTF-8 APIs (avoids OEM codepage issues).
# Usage (from repo root):
#   powershell -NoProfile -File tmp/test_paths_readonly.ps1
# With explicit path:
#   powershell -NoProfile -File tmp/test_paths_readonly.ps1 -Path 'pdf_parsing_v2_engine\templates\agcc_287\od_page1.json'

param(
    [string] $Path = 'pdf_parsing_v2_engine\templates\agcc_287\od_page1.json'
)

$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path)
Set-Location $root

$p = Join-Path $root $Path
Write-Host "Testing:" $p
if (-not (Test-Path -LiteralPath $p)) {
    Write-Host "NOT FOUND (LiteralPath). Exit 2."
    exit 2
}
$txt = [System.IO.File]::ReadAllText($p, [System.Text.UTF8Encoding]::new($false))
$null = $txt | ConvertFrom-Json
Write-Host "OK: UTF-8 JSON parses, length chars:" $txt.Length
