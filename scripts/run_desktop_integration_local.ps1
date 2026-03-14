$ErrorActionPreference = "Stop"

$projectRoot = Split-Path -Parent $PSScriptRoot
$configPath = Join-Path $projectRoot "config\desktop.runtime.integration.local.json"

if (-not (Test-Path $configPath)) {
    throw "Config de integracao local nao encontrado: $configPath"
}

$env:ROTA_CONFIG_FILE = $configPath
$env:ROTA_SECRET = "rota-secreta-local"

Write-Host "Iniciando desktop em modo de integracao local..." -ForegroundColor Cyan
Write-Host "Config: $configPath"
Write-Host "API local: http://127.0.0.1:8000"
Write-Host "Banco desktop de integracao: .rotahub_runtime\\desktop\\staging\\local-integration\\rotahub_integration.db"
Write-Host ""
Write-Host "O python main.py padrao continua em LOCAL ONLY e nao foi alterado." -ForegroundColor Yellow
Write-Host ""

$apiReady = $false
for ($attempt = 1; $attempt -le 5; $attempt++) {
    try {
        $resp = Invoke-WebRequest -Uri "http://127.0.0.1:8000/openapi.json" -UseBasicParsing -TimeoutSec 2
        if ($resp.StatusCode -eq 200) {
            $apiReady = $true
            break
        }
    } catch {
        Start-Sleep -Milliseconds 600
    }
}

if (-not $apiReady) {
    throw "API local http://127.0.0.1:8000 nao respondeu. Inicie primeiro .\\scripts\\run_api_integration_local.ps1 e depois execute este script novamente."
}

python (Join-Path $projectRoot "main.py")
