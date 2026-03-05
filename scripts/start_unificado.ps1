Param(
    [string]$ApiHost = "http://127.0.0.1:8000"
)

$ErrorActionPreference = "Stop"

$root = Split-Path -Parent $PSScriptRoot
Set-Location $root

Write-Host "== RotaHub / Modo Unificado ==" -ForegroundColor Cyan

# Carrega variaveis do .env (se existir)
$envFile = Join-Path $root ".env"
if (Test-Path $envFile) {
    Get-Content $envFile | ForEach-Object {
        $line = $_.Trim()
        if (-not $line -or $line.StartsWith("#")) { return }
        $parts = $line -split "=", 2
        if ($parts.Count -eq 2) {
            [System.Environment]::SetEnvironmentVariable($parts[0].Trim(), $parts[1].Trim(), "Process")
        }
    }
    Write-Host "ENV carregado de .env"
} else {
    Write-Warning ".env nao encontrado. Use .env.unificado.exemplo como base."
}

if (-not $env:ROTA_SECRET) { throw "ROTA_SECRET nao definido." }
if (-not $env:ROTA_DB) { throw "ROTA_DB nao definido." }

# Forca modo sincronizado
$env:ROTA_DESKTOP_SYNC_API = "1"
$env:ROTA_SERVER_URL = $ApiHost

Write-Host "ROTA_DB=$env:ROTA_DB"
Write-Host "ROTA_SERVER_URL=$env:ROTA_SERVER_URL"
Write-Host "ROTA_DESKTOP_SYNC_API=$env:ROTA_DESKTOP_SYNC_API"

Write-Host ""
Write-Host "1) Iniciando API local..." -ForegroundColor Yellow
Start-Process powershell -ArgumentList "-NoExit", "-Command", "cd `"$root`"; python api_server.py"

Start-Sleep -Seconds 1

Write-Host "2) Iniciando Desktop..." -ForegroundColor Yellow
python main.py

