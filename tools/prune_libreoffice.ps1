param(
    [Parameter(Mandatory = $true)]
    [string]$RuntimeDir
)

$ErrorActionPreference = "Stop"
$runtime = (Resolve-Path -LiteralPath $RuntimeDir).Path
$soffice = Join-Path $runtime "program\soffice.exe"
if (-not (Test-Path -LiteralPath $soffice)) {
    throw "Refusing to prune invalid LibreOffice runtime: $runtime"
}

function Remove-RuntimePath {
    param([string]$RelativePath)
    $target = Join-Path $runtime $RelativePath
    if (-not (Test-Path -LiteralPath $target)) { return }
    $resolved = (Resolve-Path -LiteralPath $target).Path
    if (-not $resolved.StartsWith($runtime + [IO.Path]::DirectorySeparatorChar, [StringComparison]::OrdinalIgnoreCase)) {
        throw "Refusing to remove path outside runtime: $resolved"
    }
    Remove-Item -LiteralPath $resolved -Recurse -Force
}

# Headless Writer conversion does not use these optional modules or user-facing assets.
@(
    "help",
    "readmes",
    "System",
    "System64",
    "share\extensions",
    "share\gallery",
    "share\template",
    "share\wizards",
    "share\Scripts",
    "share\basic",
    "share\autotext",
    "share\autocorr",
    "share\xpdfimport",
    "share\firebird",
    "share\fingerprint",
    "share\labels",
    "share\tipoftheday",
    "program\python.exe",
    "program\python312.dll",
    "program\python312.zip",
    "program\pythonloaderlo.dll",
    "program\pythonscript.py",
    "program\wizards",
    "program\classes",
    "program\shlxthdl",
    "program\shell"
) | ForEach-Object { Remove-RuntimePath $_ }

Get-ChildItem -LiteralPath (Join-Path $runtime "program") -Directory -Filter "python-core-*" -ErrorAction SilentlyContinue |
    ForEach-Object { Remove-RuntimePath ("program\" + $_.Name) }

$resourceRoot = Join-Path $runtime "program\resource"
if (Test-Path -LiteralPath $resourceRoot) {
    Get-ChildItem -LiteralPath $resourceRoot -Directory |
        Where-Object { $_.Name -notin @("common", "zh_CN") } |
        ForEach-Object { Remove-RuntimePath ("program\resource\" + $_.Name) }
}

$registryRoot = Join-Path $runtime "share\registry"
if (Test-Path -LiteralPath $registryRoot) {
    Get-ChildItem -LiteralPath $registryRoot -File -Filter "Langpack-*.xcd" |
        Where-Object { $_.Name -ne "Langpack-zh-CN.xcd" } |
        ForEach-Object { Remove-RuntimePath ("share\registry\" + $_.Name) }
}

Get-ChildItem -LiteralPath $runtime -Recurse -File -Filter "*.sample" -ErrorAction SilentlyContinue |
    ForEach-Object {
        $relative = $_.FullName.Substring($runtime.Length + 1)
        Remove-RuntimePath $relative
    }

$registryResources = Join-Path $registryRoot "res"
if (Test-Path -LiteralPath $registryResources) {
    Get-ChildItem -LiteralPath $registryResources -File |
        Where-Object { $_.Name -notmatch '(_zh-CN|_en-US)\.xcd$' } |
        ForEach-Object { Remove-RuntimePath ("share\registry\res\" + $_.Name) }
}

$configRoot = Join-Path $runtime "share\config"
if (Test-Path -LiteralPath $configRoot) {
    Get-ChildItem -LiteralPath $configRoot -File -Filter "images_*.zip" |
        Where-Object { $_.Name -ne "images_colibre.zip" } |
        ForEach-Object { Remove-RuntimePath ("share\config\" + $_.Name) }
}

Write-Host "Pruned LibreOffice runtime for headless Writer/PDF conversion: $runtime"
