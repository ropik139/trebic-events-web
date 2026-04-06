[CmdletBinding()]
param(
    [string]$ConfigPath = "..\\trebic-events.settings.json",
    [string]$SitePath = "..\\docs",
    [switch]$SkipUpdate
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

function Resolve-ProjectPath {
    param([string]$Path)
    $scriptRoot = Split-Path -Parent $PSCommandPath
    [System.IO.Path]::GetFullPath((Join-Path $scriptRoot $Path))
}

function Resolve-PathFromBase {
    param(
        [string]$BasePath,
        [string]$RelativePath
    )

    if ([System.IO.Path]::IsPathRooted($RelativePath)) {
        return $RelativePath
    }

    [System.IO.Path]::GetFullPath((Join-Path $BasePath $RelativePath))
}

function Ensure-Directory {
    param([string]$Path)
    if (-not (Test-Path -LiteralPath $Path)) {
        New-Item -ItemType Directory -Path $Path | Out-Null
    }
}

function Read-JsonFile {
    param([string]$Path)
    if (-not (Test-Path -LiteralPath $Path)) {
        return $null
    }

    $raw = Get-Content -LiteralPath $Path -Raw -Encoding UTF8
    if ([string]::IsNullOrWhiteSpace($raw)) {
        return $null
    }

    $raw | ConvertFrom-Json
}

$configFullPath = Resolve-ProjectPath -Path $ConfigPath
$config = Read-JsonFile -Path $configFullPath
if ($null -eq $config) {
    throw "Nepodarilo se nacist konfiguraci '$configFullPath'."
}

$scriptDirectory = Split-Path -Parent $PSCommandPath
if (-not $SkipUpdate) {
    & (Join-Path $scriptDirectory "update-trebic-events.ps1") -ConfigPath $ConfigPath
}

$configDirectory = Split-Path -Parent $configFullPath
$reportPath = Resolve-PathFromBase -BasePath $configDirectory -RelativePath $config.output.reportPath
$itemsPath = Resolve-PathFromBase -BasePath $configDirectory -RelativePath $config.output.itemsPath
$runLogPath = Resolve-PathFromBase -BasePath $configDirectory -RelativePath $config.output.runLogPath
$pdfPath = [System.IO.Path]::Combine((Split-Path -Parent $reportPath), "trebic-events.pdf")
$siteRoot = Resolve-ProjectPath -Path $SitePath
$siteDataPath = Join-Path $siteRoot "data"

Ensure-Directory -Path $siteRoot
Ensure-Directory -Path $siteDataPath

Copy-Item -LiteralPath $reportPath -Destination (Join-Path $siteRoot "index.html") -Force
if (Test-Path -LiteralPath $pdfPath) {
    Copy-Item -LiteralPath $pdfPath -Destination (Join-Path $siteRoot "trebic-events.pdf") -Force
}
Copy-Item -LiteralPath $itemsPath -Destination (Join-Path $siteDataPath "trebic-events-items.json") -Force
Copy-Item -LiteralPath $runLogPath -Destination (Join-Path $siteDataPath "trebic-events-last-run.json") -Force

Set-Content -LiteralPath (Join-Path $siteRoot ".nojekyll") -Value "" -Encoding ASCII

$runLog = Read-JsonFile -Path $runLogPath
$generatedAt = if ($null -ne $runLog) { [string]$runLog.generatedAt } else { (Get-Date).ToString("o") }
$siteManifest = [ordered]@{
    generatedAt = $generatedAt
    sitePath    = $siteRoot
    indexPath   = (Join-Path $siteRoot "index.html")
    pdfPath     = if (Test-Path -LiteralPath (Join-Path $siteRoot "trebic-events.pdf")) { (Join-Path $siteRoot "trebic-events.pdf") } else { "" }
}
$siteManifest | ConvertTo-Json -Depth 10 | Set-Content -LiteralPath (Join-Path $siteDataPath "site-manifest.json") -Encoding UTF8

Write-Host "GitHub Pages web je pripraven."
Write-Host "Publikovany vstup: $(Join-Path $siteRoot 'index.html')"
