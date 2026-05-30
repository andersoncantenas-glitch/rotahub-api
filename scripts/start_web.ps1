$ErrorActionPreference = "Stop"

$projectRoot = Split-Path -Parent $PSScriptRoot
Set-Location $projectRoot

$useExternalEnv = $env:ROTA_USE_EXTERNAL_ENV -eq "1"

if ((-not $env:ROTA_DB) -or (-not $useExternalEnv)) {
    $env:ROTA_DB = Join-Path $projectRoot "rotadb.db"
}
if ((-not $env:DATABASE_URL) -or (-not $useExternalEnv)) {
    $dbPath = $env:ROTA_DB.Replace("\", "/")
    $env:DATABASE_URL = "sqlite+aiosqlite:///$dbPath"
}
if ((-not $env:ROTA_MOBILE_PHOTOS_DIR) -or (-not $useExternalEnv)) {
    $env:ROTA_MOBILE_PHOTOS_DIR = Join-Path $projectRoot ".rotahub_runtime\fotos_rotas"
}
if (-not $env:ROTA_ENABLE_LEGACY_MOBILE_API) {
    $env:ROTA_ENABLE_LEGACY_MOBILE_API = "1"
}
if (-not $env:HOST) {
    $env:HOST = "0.0.0.0"
}
if (-not $env:PORT) {
    $env:PORT = "8000"
}

$python = $env:PYTHON_EXE
if (-not $python) {
    $python = "python"
}
& $python -m uvicorn backend.main:app --host $env:HOST --port ([int]$env:PORT)
