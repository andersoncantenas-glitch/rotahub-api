$ErrorActionPreference = "Stop"

$projectRoot = Split-Path -Parent $PSScriptRoot
Set-Location $projectRoot

$env:DATABASE_URL = "sqlite+aiosqlite:///C:/pdc_rota/banco.db"
$env:ROTA_DB = "C:/pdc_rota/banco.db"
$env:APP_ENV = "development"
$env:ROTA_SECRET = "rota-secreta"
$env:ALLOWED_HOSTS = "localhost,127.0.0.1,0.0.0.0,10.0.2.2"
$env:CORS_ORIGINS = "http://localhost:8000,http://127.0.0.1:8000,http://10.0.2.2:8000"
$env:ROTA_ENABLE_LEGACY_MOBILE_API = "1"
$env:HOST = "0.0.0.0"
$env:PORT = "8000"

$pythonExe = Join-Path $projectRoot ".venv\Scripts\python.exe"
if (-not (Test-Path $pythonExe)) {
    $pythonExe = "python"
}

& $pythonExe -m uvicorn backend.main:app --host $env:HOST --port ([int]$env:PORT)
