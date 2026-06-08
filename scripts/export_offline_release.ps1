param(
  [string]$OutputDir = "",
  [string]$PackageName = "",
  [switch]$IncludeWorkingTree = $true
)

$ErrorActionPreference = "Stop"

$Root = Resolve-Path (Join-Path $PSScriptRoot "..")
if (-not $OutputDir) {
  $OutputDir = Join-Path $Root "deploy_bundle"
}
$OutputDir = [System.IO.Path]::GetFullPath($OutputDir)
New-Item -ItemType Directory -Force -Path $OutputDir | Out-Null

$stamp = Get-Date -Format "yyyyMMdd_HHmmss"
if (-not $PackageName) {
  $PackageName = "bom-tools_offline_$stamp.zip"
}
$packagePath = Join-Path $OutputDir $PackageName
$staging = Join-Path $env:TEMP "bom-tools_offline_$stamp"

if (Test-Path $staging) {
  Remove-Item -Recurse -Force $staging
}
New-Item -ItemType Directory -Force -Path $staging | Out-Null

$dataExcludes = @(
  "web_app2/auth_data",
  "web_app2/cache",
  "web_app2/uploads",
  "web_app2/outputs",
  "web_app2/logs",
  "web_app2/bug_reports",
  "web_app2/feature_requests",
  "web_app2/manufacturer_aliases",
  "web_app2/__pycache__"
)

function Test-IsExcluded {
  param([string]$RelativePath)
  $normalized = $RelativePath -replace "\\", "/"
  if ($normalized -eq ".git" -or $normalized.StartsWith(".git/")) { return $true }
  if ($normalized -eq "deploy_bundle") { return $true }
  if ($normalized.StartsWith("deploy_bundle/offline_")) { return $true }
  if ($normalized -eq "deploy_bundle/web_app2" -or $normalized.StartsWith("deploy_bundle/web_app2/")) { return $true }
  if ($normalized -eq "deploy_bundle/README.txt") { return $true }
  if ($normalized -eq "deploy_bundle/requirements.txt") { return $true }
  if ($normalized -eq "reference_LiXinYu" -or $normalized.StartsWith("reference_LiXinYu/")) { return $true }
  if ($normalized -match '(^|/)__pycache__(/|$)') { return $true }
  if ($normalized.EndsWith(".pyc")) { return $true }
  if ($normalized.EndsWith(".pyo")) { return $true }
  foreach ($prefix in $dataExcludes) {
    if ($normalized -eq $prefix -or $normalized.StartsWith("$prefix/")) {
      return $true
    }
  }
  return $false
}

function Copy-ReleaseFile {
  param([string]$RelativePath)
  if (Test-IsExcluded $RelativePath) { return }
  $src = Join-Path $Root $RelativePath
  if (-not (Test-Path $src -PathType Leaf)) { return }
  $dest = Join-Path $staging $RelativePath
  $destDir = Split-Path $dest -Parent
  New-Item -ItemType Directory -Force -Path $destDir | Out-Null
  Copy-Item -LiteralPath $src -Destination $dest -Force
}

function Get-RelativePathCompat {
  param(
    [Parameter(Mandatory=$true)][string]$BasePath,
    [Parameter(Mandatory=$true)][string]$FullPath
  )
  $base = [System.IO.Path]::GetFullPath($BasePath).TrimEnd('\') + '\'
  $baseUri = New-Object System.Uri($base)
  $fullUri = New-Object System.Uri([System.IO.Path]::GetFullPath($FullPath))
  $relUri = $baseUri.MakeRelativeUri($fullUri)
  return [System.Uri]::UnescapeDataString($relUri.ToString()).Replace('/', '\')
}

Push-Location $Root
try {
  if ($IncludeWorkingTree) {
    Get-ChildItem -LiteralPath $Root -Recurse -File -Force | ForEach-Object {
      $rel = Get-RelativePathCompat -BasePath $Root -FullPath $_.FullName
      Copy-ReleaseFile $rel
    }
  } else {
    $tracked = git -c core.quotepath=false ls-files
    foreach ($rel in $tracked) {
      Copy-ReleaseFile $rel
    }
  }
}
finally {
  Pop-Location
}

$manifest = [ordered]@{
  generated_at = (Get-Date).ToString("s")
  package_name = $PackageName
  include_working_tree = [bool]$IncludeWorkingTree
  excluded_runtime_data = $dataExcludes
  one_click_entry = "deploy_one_click.bat"
  install_script = "scripts/install_offline_release.ps1"
}
$manifest | ConvertTo-Json -Depth 5 | Set-Content -Path (Join-Path $staging "offline_release_manifest.json") -Encoding UTF8

$readme = @"
BOM Tools offline release

Install on server:

  powershell -ExecutionPolicy Bypass -File .\scripts\install_offline_release.ps1 -PackagePath .\$PackageName -InstallDir C:\path\to\bom-tools

One-click entry:

  deploy_one_click.bat
  deploy_one_click.bat C:\path\to\bom-tools

The one-click script deploys the package when available, preserves runtime
data and offline assets, installs dependencies, then starts the web service.

Runtime data is intentionally excluded from this package:

  web_app2/auth_data/users.sqlite3
  web_app2/cache
  web_app2/uploads
  web_app2/outputs
  web_app2/logs
  web_app2/bug_reports
  web_app2/feature_requests
  web_app2/manufacturer_aliases

The install script backs up those directories before replacing code, then restores them.
"@
$readme | Set-Content -Path (Join-Path $staging "OFFLINE_DEPLOY_README.txt") -Encoding UTF8

if (Test-Path $packagePath) {
  Remove-Item -Force $packagePath
}
Compress-Archive -Path (Join-Path $staging "*") -DestinationPath $packagePath -Force
Remove-Item -Recurse -Force $staging

Write-Host "Offline release package created:"
Write-Host $packagePath
