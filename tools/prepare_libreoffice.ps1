param(
    [string]$Version = "26.2.5",
    [switch]$Force
)

$ErrorActionPreference = "Stop"
$repoRoot = Split-Path -Parent $PSScriptRoot
$downloadsDir = Join-Path $repoRoot "downloads"
$vendorDir = Join-Path $repoRoot "vendor"
$runtimeDir = Join-Path $vendorDir "libreoffice"
$sofficePath = Join-Path $runtimeDir "program\soffice.exe"

if ((Test-Path -LiteralPath $sofficePath) -and -not $Force) {
    & (Join-Path $PSScriptRoot "prune_libreoffice.ps1") -RuntimeDir $runtimeDir
    Write-Host "LibreOffice runtime already prepared: $sofficePath"
    exit 0
}

New-Item -ItemType Directory -Force -Path $downloadsDir, $vendorDir | Out-Null
$fileName = "LibreOffice_${Version}_Win_x86-64.msi"
$downloadPath = Join-Path $downloadsDir $fileName
$downloadUrl = "https://download.documentfoundation.org/libreoffice/stable/$Version/win/x86_64/$fileName"

if (-not (Test-Path -LiteralPath $downloadPath)) {
    Write-Host "Downloading official LibreOffice $Version runtime..."
    Invoke-WebRequest -UseBasicParsing -Uri $downloadUrl -OutFile $downloadPath
}

$downloadInfo = Get-Item -LiteralPath $downloadPath
if ($downloadInfo.Length -lt 100MB) {
    throw "LibreOffice download is incomplete: $downloadPath"
}

$extractRoot = Join-Path $vendorDir "libreoffice-image"
if (Test-Path -LiteralPath $extractRoot) {
    Remove-Item -LiteralPath $extractRoot -Recurse -Force
}
New-Item -ItemType Directory -Force -Path $extractRoot | Out-Null

Write-Host "Extracting LibreOffice without installing it..."
$process = Start-Process -FilePath "msiexec.exe" -ArgumentList @(
    "/a",
    ('"' + $downloadPath + '"'),
    "/qn",
    ('TARGETDIR="' + $extractRoot + '"')
) -Wait -PassThru -WindowStyle Hidden
if ($process.ExitCode -ne 0) {
    throw "LibreOffice administrative extraction failed with exit code $($process.ExitCode)."
}

$locatedSoffice = Get-ChildItem -LiteralPath $extractRoot -Recurse -Filter "soffice.exe" -File |
    Where-Object { $_.DirectoryName -match '[\\/]program$' } |
    Select-Object -First 1
if (-not $locatedSoffice) {
    throw "The extracted package does not contain program\soffice.exe."
}

$officeRoot = Split-Path -Parent $locatedSoffice.DirectoryName
if (Test-Path -LiteralPath $runtimeDir) {
    Remove-Item -LiteralPath $runtimeDir -Recurse -Force
}
New-Item -ItemType Directory -Force -Path $runtimeDir | Out-Null
Copy-Item -Path (Join-Path $officeRoot "*") -Destination $runtimeDir -Recurse -Force

if (-not (Test-Path -LiteralPath $sofficePath)) {
    throw "Prepared runtime is invalid: $sofficePath"
}

& (Join-Path $PSScriptRoot "prune_libreoffice.ps1") -RuntimeDir $runtimeDir
Remove-Item -LiteralPath $extractRoot -Recurse -Force
Write-Host "Prepared LibreOffice runtime: $runtimeDir"
