Param(
    [string]$ApiHost = "http://127.0.0.1:8000"
)

$ErrorActionPreference = "Stop"

$root = Split-Path -Parent $PSScriptRoot
Set-Location $root

Write-Host "== RotaHub / Modo Unificado ==" -ForegroundColor Cyan
if ($ApiHost -ne "http://127.0.0.1:8000") {
    Write-Warning "Este script foi padronizado para a API local de integracao em http://127.0.0.1:8000."
}

Write-Host ""
Write-Host "1) Iniciando API local..." -ForegroundColor Yellow
Start-Process powershell -ArgumentList "-NoExit", "-ExecutionPolicy", "Bypass", "-File", (Join-Path $root "scripts\run_api_integration_local.ps1")

Start-Sleep -Seconds 2

Write-Host "2) Iniciando Desktop integrado..." -ForegroundColor Yellow
powershell -ExecutionPolicy Bypass -File (Join-Path $root "scripts\run_desktop_integration_local.ps1")
