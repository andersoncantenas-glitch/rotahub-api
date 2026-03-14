$ErrorActionPreference = "Stop"

$projectRoot = Split-Path -Parent $PSScriptRoot
$configPath = Join-Path $projectRoot "config\server.runtime.integration.local.json"

if (-not (Test-Path $configPath)) {
    throw "Config da API local de integracao nao encontrado: $configPath"
}

$config = Get-Content -Path $configPath -Raw | ConvertFrom-Json
$runtime = $config.runtime
$tenant = $config.tenant
$api = $config.api
$logging = $config.logging

$dbPath = [string]$runtime.db_path
$apiBaseUrl = ([string]$api.base_url).TrimEnd("/")
$desktopSecret = [string]$runtime.desktop_secret
$appEnv = [string]$config.app_env
$tenantId = [string]$tenant.tenant_id
$companyId = [string]$tenant.company_id
$logLevel = [string]$logging.level

if (-not $dbPath) {
    throw "db_path nao configurado em $configPath"
}
if (-not $desktopSecret) {
    throw "desktop_secret nao configurado em $configPath"
}

$dbDir = Split-Path -Parent $dbPath
if ($dbDir) {
    New-Item -ItemType Directory -Force -Path $dbDir | Out-Null
}

$env:ROTA_CONFIG_FILE = $configPath
$env:ROTA_APP_ENV = $appEnv
$env:ROTA_TENANT_ID = $tenantId
$env:ROTA_COMPANY_ID = $companyId
$env:ROTA_DB = $dbPath
$env:ROTA_SERVER_URL = $apiBaseUrl
$env:ROTA_SECRET = $desktopSecret
$env:ROTA_LOG_LEVEL = $logLevel
$env:ROTA_DESKTOP_SYNC_API = "0"
$env:ROTA_ALLOW_REMOTE_READ = "0"
$env:ROTA_ALLOW_REMOTE_WRITE = "0"

Write-Host "Iniciando API local de integracao..." -ForegroundColor Cyan
Write-Host "Config: $configPath"
Write-Host "URL local: $apiBaseUrl"
Write-Host "Banco servidor local: $dbPath"
Write-Host "Tenant: $tenantId"
Write-Host ""
Write-Host "Esse ambiente eh separado do desktop LOCAL ONLY e separado do Render." -ForegroundColor Yellow
Write-Host "Para Android Emulator, use 10.0.2.2:8000 no app mobile." -ForegroundColor Yellow
Write-Host ""

python -m uvicorn api_server:app --host 0.0.0.0 --port 8000

