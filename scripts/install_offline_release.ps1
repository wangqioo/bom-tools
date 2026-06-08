param(
  [Parameter(Mandatory=$true)][string]$PackagePath,
  [Parameter(Mandatory=$true)][string]$InstallDir,
  [string]$UsersDbPath = "",
  [switch]$InstallRequirements
)

$ErrorActionPreference = "Stop"

$PackagePath = [System.IO.Path]::GetFullPath($PackagePath)
$InstallDir = [System.IO.Path]::GetFullPath($InstallDir)
if (-not (Test-Path $PackagePath -PathType Leaf)) {
  throw "Package not found: $PackagePath"
}

$stamp = Get-Date -Format "yyyyMMdd_HHmmss"
$backupRoot = Join-Path $InstallDir "_runtime_backup_$stamp"
$extractDir = Join-Path $env:TEMP "bom-tools_install_$stamp"

$runtimePaths = @(
  "web_app2/auth_data",
  "web_app2/cache",
  "web_app2/uploads",
  "web_app2/outputs",
  "web_app2/logs",
  "web_app2/bug_reports",
  "web_app2/feature_requests",
  "web_app2/manufacturer_aliases"
)

$deploymentAssetPaths = @(
  "deploy_bundle/wheels",
  "deploy_bundle/ms-playwright",
  "deploy_bundle/offline_final",
  "wheels",
  "ms-playwright"
)

$configPaths = @(
  ".env",
  "web_app2/.env",
  "web_app2/run.ps1"
)

function Copy-DirectoryContents {
  param(
    [Parameter(Mandatory=$true)][string]$Source,
    [Parameter(Mandatory=$true)][string]$Destination
  )
  New-Item -ItemType Directory -Force -Path $Destination | Out-Null
  Copy-Item -Path (Join-Path $Source "*") -Destination $Destination -Recurse -Force
}

if (Test-Path $extractDir) {
  Remove-Item -Recurse -Force $extractDir
}
New-Item -ItemType Directory -Force -Path $extractDir | Out-Null
Expand-Archive -Path $PackagePath -DestinationPath $extractDir -Force

New-Item -ItemType Directory -Force -Path $InstallDir | Out-Null
New-Item -ItemType Directory -Force -Path $backupRoot | Out-Null

foreach ($rel in ($runtimePaths + $deploymentAssetPaths)) {
  $src = Join-Path $InstallDir $rel
  if (Test-Path $src) {
    $dest = Join-Path $backupRoot $rel
    New-Item -ItemType Directory -Force -Path (Split-Path $dest -Parent) | Out-Null
    Copy-Item -LiteralPath $src -Destination $dest -Recurse -Force
  }
}

foreach ($rel in $configPaths) {
  $src = Join-Path $InstallDir $rel
  if (Test-Path $src -PathType Leaf) {
    $dest = Join-Path $backupRoot $rel
    New-Item -ItemType Directory -Force -Path (Split-Path $dest -Parent) | Out-Null
    Copy-Item -LiteralPath $src -Destination $dest -Force
  }
}

$preserveNames = @((Split-Path $backupRoot -Leaf))
Get-ChildItem -Force -LiteralPath $InstallDir | ForEach-Object {
  if ($preserveNames -contains $_.Name) { return }
  Remove-Item -LiteralPath $_.FullName -Recurse -Force
}

Copy-DirectoryContents -Source $extractDir -Destination $InstallDir

foreach ($rel in ($runtimePaths + $deploymentAssetPaths)) {
  $backupPath = Join-Path $backupRoot $rel
  if (Test-Path $backupPath) {
    $dest = Join-Path $InstallDir $rel
    if (Test-Path $dest) {
      Remove-Item -LiteralPath $dest -Recurse -Force
    }
    New-Item -ItemType Directory -Force -Path (Split-Path $dest -Parent) | Out-Null
    Copy-Item -LiteralPath $backupPath -Destination $dest -Recurse -Force
  }
}

foreach ($rel in $configPaths) {
  $backupPath = Join-Path $backupRoot $rel
  if (Test-Path $backupPath -PathType Leaf) {
    $dest = Join-Path $InstallDir $rel
    New-Item -ItemType Directory -Force -Path (Split-Path $dest -Parent) | Out-Null
    Copy-Item -LiteralPath $backupPath -Destination $dest -Force
  }
}

if ($UsersDbPath) {
  $UsersDbPath = [System.IO.Path]::GetFullPath($UsersDbPath)
  if (-not (Test-Path $UsersDbPath -PathType Leaf)) {
    throw "Users database not found: $UsersDbPath"
  }
  $destDb = Join-Path $InstallDir "web_app2\auth_data\users.sqlite3"
  New-Item -ItemType Directory -Force -Path (Split-Path $destDb -Parent) | Out-Null
  Copy-Item -LiteralPath $UsersDbPath -Destination $destDb -Force
  foreach ($suffix in @("-wal", "-shm")) {
    $sidecar = "$destDb$suffix"
    if (Test-Path $sidecar) {
      Remove-Item -LiteralPath $sidecar -Force
    }
  }
}

$finalDb = Join-Path $InstallDir "web_app2\auth_data\users.sqlite3"
if (-not (Test-Path $finalDb -PathType Leaf)) {
  throw "users.sqlite3 is missing after install: $finalDb"
}

Remove-Item -Recurse -Force $extractDir

if ($InstallRequirements) {
  $requirements = Join-Path $InstallDir "web_app2\requirements.txt"
  if (Test-Path $requirements) {
    python -m pip install -r $requirements
  }
}

Write-Host "Offline release installed:"
Write-Host $InstallDir
Write-Host "Runtime backup:"
Write-Host $backupRoot
Write-Host "Restart the BOM Tools service after verifying the backup."
