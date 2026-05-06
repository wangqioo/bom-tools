param(
  [string]$ProjectRoot = (Resolve-Path (Join-Path $PSScriptRoot '..')).Path,
  [string]$OutDir = '',
  [string]$Name = 'dehdl-b',
  [string]$PythonDir = '',
  [string]$PythonArchive = '',
  [string]$PythonVersion = '3.10.11',
  [ValidateSet('official', 'tuna', 'npmmirror')]
  [string]$PythonMirror = 'tuna',
  [string]$MineruVenv = '',
  [string]$MineruModelDir = '',
  [string]$MineruConfig = '',
  [ValidateSet('huggingface', 'modelscope')]
  [string]$MineruModelSource = 'huggingface',
  [ValidateSet('pipeline', 'vlm', 'all')]
  [string]$MineruModelType = 'pipeline',
  [string]$HuggingFaceEndpoint = 'https://hf-mirror.com',
  [string]$MineruModelDownloader = '',
  [string]$PipIndexUrl = 'https://pypi.tuna.tsinghua.edu.cn/simple',
  [string]$PipExtraIndexUrl = 'https://download.pytorch.org/whl/cu121',
  [string]$MineruWheelSpec = 'mineru[pipeline]',
  [string]$AssetCacheDir = '',
  [switch]$NoDownloadMineruModels,
  [switch]$NoDownloadWheels,
  [switch]$NoIncludeMineruWheels,
  [switch]$StrictMineruWheels,
  [switch]$NoReuseAssets,
  [switch]$IncludeDatasheetSource,
  [switch]$NoZip,
  [switch]$MakeZip
)

$ErrorActionPreference = 'Stop'

if (-not $OutDir) {
  $OutDir = Join-Path $ProjectRoot 'output\offline_migration'
}
if (-not $AssetCacheDir) {
  $AssetCacheDir = Join-Path $OutDir '_asset_cache'
}
if (-not $MineruVenv) {
  $candidate = Join-Path $ProjectRoot '.venv-mineru'
  if (Test-Path $candidate) { $MineruVenv = $candidate }
}
if (-not $MineruModelDir -and $env:PSTX_MINERU_MODEL_DIR) {
  $MineruModelDir = $env:PSTX_MINERU_MODEL_DIR
}
if (-not $MineruConfig -and $env:MINERU_TOOLS_CONFIG_JSON) {
  $MineruConfig = $env:MINERU_TOOLS_CONFIG_JSON
}
if (-not $MineruConfig) {
  $defaultConfig = Join-Path $env:USERPROFILE '.mineru\mineru.json'
  if (Test-Path $defaultConfig) { $MineruConfig = $defaultConfig }
}

$pythonLauncher = 'python'
if (Get-Command py -ErrorAction SilentlyContinue) {
  $pythonLauncher = 'py'
}

$argsList = @(
  'pstx_cli.py',
  'offline-migration',
  'prepare',
  '--project-root', $ProjectRoot,
  '--out-dir', $OutDir,
  '--name', $Name,
  '--target-platform', 'windows-amd64',
  '--target-profile', 'windows-rtx4060-cuda',
  '--python-mirror', $PythonMirror,
  '--pretty'
)

if ($PythonDir) {
  $argsList += @('--python-dir', $PythonDir)
} elseif ($PythonArchive) {
  $argsList += @('--python-archive', $PythonArchive)
} else {
  $argsList += @('--python-version', $PythonVersion)
}

if ($MineruVenv) {
  $argsList += @('--mineru-venv', $MineruVenv)
} else {
  Write-Warning 'No MinerU venv was found. If MinerU models/config are prepared, offline-migration will try to create .venv-mineru with mineru[pipeline]. Pass -MineruVenv to use a tested runtime.'
}

if ($MineruModelDir) {
  $argsList += @('--mineru-model-dir', $MineruModelDir)
} else {
  if (-not $NoDownloadMineruModels) {
    $argsList += @('--download-mineru-models', '--mineru-model-source', $MineruModelSource, '--mineru-model-type', $MineruModelType)
    if ($HuggingFaceEndpoint -and $MineruModelSource -eq 'huggingface') {
      $argsList += @('--huggingface-endpoint', $HuggingFaceEndpoint)
    }
    if ($MineruModelDownloader) {
      $argsList += @('--mineru-model-downloader', $MineruModelDownloader)
    }
  } else {
    Write-Warning 'No MinerU model directory was provided. Pass -MineruModelDir or remove -NoDownloadMineruModels to include offline models for computer B.'
  }
}

if ($MineruConfig) {
  $argsList += @('--mineru-config', $MineruConfig)
} else {
  if ($MineruModelDir -or $NoDownloadMineruModels) {
    Write-Warning 'No mineru.json was provided. Pass -MineruConfig or set MINERU_TOOLS_CONFIG_JSON to patch local model paths on computer B.'
  }
}

if (-not $NoDownloadWheels) {
  $argsList += @('--download-wheels', '--pip-index-url', $PipIndexUrl)
  if ($PipExtraIndexUrl) {
    $argsList += @('--pip-extra-index-url', $PipExtraIndexUrl)
  }
}
if ($AssetCacheDir) {
  $argsList += @('--asset-cache-dir', $AssetCacheDir)
}
if ($NoReuseAssets) {
  $argsList += '--no-reuse-assets'
}
if (-not $NoIncludeMineruWheels) {
  $argsList += @('--include-mineru-wheels', '--mineru-wheel-spec', $MineruWheelSpec)
  if ($StrictMineruWheels) {
    $argsList += '--strict-mineru-wheels'
  }
}
if ($IncludeDatasheetSource) {
  $argsList += '--include-datasheet-source'
}
if (-not $MakeZip -or $NoZip) {
  $argsList += '--no-zip'
}

Push-Location $ProjectRoot
try {
  Write-Host "Preparing PSTX offline migration bundle for Windows RTX 4060/CUDA..."
  & $pythonLauncher @argsList
  if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }
  Write-Host ''
  if ($MakeZip -and -not $NoZip) {
    Write-Host 'Done. Move the generated zip/folder from output\offline_migration to computer B and run RUN_SETUP_B.bat.'
  } else {
    Write-Host 'Done. Manually compress the generated bundle folder if needed, then move it to computer B and run RUN_SETUP_B.bat.'
  }
} finally {
  Pop-Location
}
