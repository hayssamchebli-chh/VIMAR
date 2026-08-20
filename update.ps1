$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $root

Write-Host "============================================"
Write-Host "  Vimar Datasheet Tool - Update"
Write-Host "============================================"
Write-Host ""
Write-Host "Downloading the latest version..."

$zipUrl = "https://github.com/hayssamchebli-chh/VIMAR/archive/refs/heads/verified-browser-download.zip"
$tmpZip = Join-Path $env:TEMP "vimar_update.zip"
$tmpDir = Join-Path $env:TEMP "vimar_update_extract"

try {
    Invoke-WebRequest -Uri $zipUrl -OutFile $tmpZip -UseBasicParsing
} catch {
    Write-Host ""
    Write-Host "Could not download the update. Check your internet connection and try again."
    Write-Host "Error: $_"
    Read-Host "Press Enter to close"
    exit 1
}

if (Test-Path $tmpDir) { Remove-Item $tmpDir -Recurse -Force }
Expand-Archive -Path $tmpZip -DestinationPath $tmpDir -Force

$extracted = Get-ChildItem $tmpDir -Directory | Select-Object -First 1
if (-not $extracted) {
    Write-Host "The downloaded update looked empty - nothing was changed."
    Read-Host "Press Enter to close"
    exit 1
}

Write-Host "Copying updated files..."
# Only the app's own files are touched. manual_datasheets/ (your saved PDFs)
# and anything else you added are never removed or overwritten.
$filesToUpdate = @(
    "app.py", "vimar.py", "merge.py",
    "requirements.txt", "README.md",
    "item_type_template.pdf", "toc_logo.png",
    "setup.bat", "start.bat", "update.bat", "update.ps1"
)
$updated = 0
foreach ($f in $filesToUpdate) {
    $src = Join-Path $extracted.FullName $f
    if (Test-Path $src) {
        Copy-Item $src -Destination $root -Force
        $updated++
    }
}
Write-Host "  $updated file(s) updated."

Write-Host ""
Write-Host "Checking requirements..."
python -m pip install -r (Join-Path $root "requirements.txt") --quiet

Write-Host ""
Write-Host "Cleaning up..."
Remove-Item $tmpZip -Force -ErrorAction SilentlyContinue
Remove-Item $tmpDir -Recurse -Force -ErrorAction SilentlyContinue

Write-Host ""
Write-Host "============================================"
Write-Host "  Update complete! You can close this window"
Write-Host "  and use the tool as normal."
Write-Host "============================================"
Read-Host "Press Enter to close"
