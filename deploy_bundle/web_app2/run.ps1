$ErrorActionPreference = "Stop"
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$OutputEncoding = [System.Text.Encoding]::UTF8
$env:PYTHONIOENCODING = "utf-8"

$port = if ($env:PORT) { $env:PORT } else { "5000" }
Write-Host "Starting BOM Tools Web on http://127.0.0.1:$port"

$waitress = python -c "import importlib.util,sys; sys.exit(0 if importlib.util.find_spec('waitress') else 1)" 2>$null
if ($LASTEXITCODE -eq 0) {
    Write-Host "Using waitress WSGI server"
    python -m waitress --host=0.0.0.0 --port=$port app:app
} else {
    Write-Host "waitress not installed; using Flask development server"
    python app.py
}
