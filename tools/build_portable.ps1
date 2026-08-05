param(
    [string]$LibreOfficeVersion = "26.2.5",
    [switch]$SkipTests
)

$ErrorActionPreference = "Stop"
$repoRoot = Split-Path -Parent $PSScriptRoot
Set-Location $repoRoot
$python = if (Test-Path -LiteralPath (Join-Path $repoRoot ".venv\Scripts\python.exe")) {
    Join-Path $repoRoot ".venv\Scripts\python.exe"
} else {
    "python"
}

& (Join-Path $PSScriptRoot "prepare_libreoffice.ps1") -Version $LibreOfficeVersion
if (-not $SkipTests) {
    & $python -m unittest discover -s tests -v
    if ($LASTEXITCODE -ne 0) { throw "Unit tests failed." }
}

& $python -m PyInstaller --noconfirm --clean VoteDocxApp.spec
if ($LASTEXITCODE -ne 0) { throw "PyInstaller build failed." }

$portableRoot = Join-Path $repoRoot "dist\QunzhongVote-v0.3.5-Portable"
$runtimeTarget = Join-Path $portableRoot "runtime\libreoffice"
if (Test-Path -LiteralPath $portableRoot) {
    Remove-Item -LiteralPath $portableRoot -Recurse -Force
}
New-Item -ItemType Directory -Force -Path $runtimeTarget | Out-Null

Copy-Item -LiteralPath (Join-Path $repoRoot "dist\QunzhongVote.exe") -Destination (Join-Path $portableRoot "QunzhongVote.exe") -Force
Copy-Item -Path (Join-Path $repoRoot "vendor\libreoffice\*") -Destination $runtimeTarget -Recurse -Force
Copy-Item -LiteralPath (Join-Path $repoRoot "PORTABLE_README.txt") -Destination $portableRoot -Force

$zipPath = Join-Path $repoRoot "dist\QunzhongVote-v0.3.5-Windows-Portable.zip"
if (Test-Path -LiteralPath $zipPath) {
    Remove-Item -LiteralPath $zipPath -Force
}
$sevenZip = @(
    "C:\Program Files\7-Zip\7z.exe",
    "C:\Program Files (x86)\7-Zip\7z.exe"
) | Where-Object { Test-Path -LiteralPath $_ } | Select-Object -First 1
if ($sevenZip) {
    & $sevenZip a -tzip -mx=9 $zipPath (Join-Path $portableRoot "*") | Out-Host
    if ($LASTEXITCODE -ne 0) { throw "7-Zip packaging failed." }
} else {
    Compress-Archive -Path (Join-Path $portableRoot "*") -DestinationPath $zipPath -CompressionLevel Optimal
}

Write-Host "Portable directory: $portableRoot"
Write-Host "Portable ZIP: $zipPath"
