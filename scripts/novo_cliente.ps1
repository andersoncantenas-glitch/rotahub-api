param(
    [Parameter(Mandatory = $true)]
    [string]$ClienteSlug,

    [Parameter(Mandatory = $true)]
    [string]$ApiUrl,

    [string]$DbPath = "",
    [string]$AppVersion = "1.0.0"
)

$ErrorActionPreference = "Stop"

function New-StrongSecret {
    return [Convert]::ToBase64String((1..48 | ForEach-Object { Get-Random -Maximum 256 }))
}

$slug = ($ClienteSlug -replace "[^a-zA-Z0-9_-]", "").ToLower()
if ([string]::IsNullOrWhiteSpace($slug)) {
    throw "ClienteSlug invalido."
}

$api = $ApiUrl.Trim().TrimEnd("/")
if (-not ($api.StartsWith("http://") -or $api.StartsWith("https://"))) {
    throw "ApiUrl deve iniciar com http:// ou https://"
}

if ([string]::IsNullOrWhiteSpace($DbPath)) {
    $DbPath = "/var/data/rota_granja_$slug.db"
}

$secret = New-StrongSecret

Write-Host ""
Write-Host "=== KIT CLIENTE NOVO ===" -ForegroundColor Cyan
Write-Host "ClienteSlug : $slug"
Write-Host "ApiUrl      : $api"
Write-Host "ROTA_DB      : $DbPath"
Write-Host "ROTA_SECRET  : $secret"
Write-Host ""

Write-Host "1) Render - Environment Variables" -ForegroundColor Yellow
Write-Host "ROTA_DB=$DbPath"
Write-Host "ROTA_SECRET=$secret"
Write-Host "ROTA_CORS_ORIGINS=$api"
Write-Host ""

Write-Host "2) Render - Start Command" -ForegroundColor Yellow
Write-Host "python -c `"import os,sqlite3; p=os.environ.get('ROTA_DB','rota_granja.db'); os.makedirs(os.path.dirname(p) or '.', exist_ok=True); sqlite3.connect(p).close()`" && python -m uvicorn api_server:app --host 0.0.0.0 --port `$PORT"
Write-Host ""

Write-Host "3) Desktop (estacao cliente) - PowerShell" -ForegroundColor Yellow
Write-Host "setx ROTA_SERVER_URL `"$api`""
Write-Host "setx ROTA_SECRET `"$secret`""
Write-Host "setx ROTA_DESKTOP_SYNC_API `"1`""
Write-Host ""

Write-Host "4) APK (build) - Terminal VSCode" -ForegroundColor Yellow
Write-Host "cd C:\flutter_application_1"
Write-Host "flutter clean"
Write-Host "flutter pub get"
Write-Host "flutter build apk --release --dart-define=API_BASE_URL=$api"
Write-Host ""

Write-Host "5) Desktop build (opcional) - Terminal VSCode" -ForegroundColor Yellow
Write-Host "cd C:\pdc_rota"
Write-Host "powershell -ExecutionPolicy Bypass -File .\scripts\build_desktop.ps1 -AppVersion $AppVersion"
Write-Host ""

Write-Host "6) Validacao API" -ForegroundColor Yellow
Write-Host "$api/ping"
Write-Host ""
