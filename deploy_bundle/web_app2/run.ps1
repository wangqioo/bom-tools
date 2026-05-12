$ErrorActionPreference = "Stop"
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$OutputEncoding = [System.Text.Encoding]::UTF8
$env:PYTHONIOENCODING = "utf-8"

$port = if ($env:PORT) { $env:PORT } else { "5000" }
Write-Host "Starting BOM Tools Web on http://127.0.0.1:$port"
python app.py
