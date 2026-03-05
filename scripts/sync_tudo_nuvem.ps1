param(
    [string]$RotaDb = "C:\rotahub\rota_granja.db",
    [string]$ServerUrl = "https://rotahub-api.onrender.com",
    [string]$Secret = "",
    [switch]$All
)

$ErrorActionPreference = "Stop"

Write-Host "=== SYNC TUDO NUVEM ===" -ForegroundColor Cyan
Write-Host "DB: $RotaDb"
Write-Host "API: $ServerUrl"

if ([string]::IsNullOrWhiteSpace($Secret)) {
    $Secret = $Env:ROTA_SECRET
}
if ([string]::IsNullOrWhiteSpace($Secret)) {
    throw "ROTA_SECRET não informado. Use -Secret ou defina a variável de ambiente ROTA_SECRET."
}

if (-not (Test-Path $RotaDb)) {
    throw "Banco não encontrado: $RotaDb"
}

$Env:ROTA_DB = $RotaDb
$Env:ROTA_SERVER_URL = $ServerUrl.TrimEnd("/")
$Env:ROTA_SECRET = $Secret

$scriptPath = Join-Path $PSScriptRoot "sync_tudo_nuvem.py"
if (-not (Test-Path $scriptPath)) {
    throw "Script Python não encontrado: $scriptPath"
}

$args = @($scriptPath)
if ($All.IsPresent) {
    $args += "--all"
}

Write-Host ""
Write-Host "Executando: python $($args -join ' ')" -ForegroundColor Yellow
& python @args
$exitCode = $LASTEXITCODE

Write-Host ""
if ($exitCode -eq 0) {
    Write-Host "Sync concluído com sucesso." -ForegroundColor Green
} else {
    Write-Host "Sync concluído com falhas (exit code: $exitCode)." -ForegroundColor Red
}

exit $exitCode

