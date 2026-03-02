param(
    [string]$AppVersion = "1.0.0"
)

$ErrorActionPreference = "Stop"

$root = Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path)
Set-Location $root

Write-Host "==> Projeto: $root"

if (-not (Test-Path ".venv")) {
    Write-Host "==> Criando .venv"
    python -m venv .venv
}

$py = Join-Path $root ".venv\Scripts\python.exe"
if (-not (Test-Path $py)) {
    throw "Python da venv nao encontrado em $py"
}

Write-Host "==> Instalando dependencias de desktop"
& $py -m pip install --upgrade pip
& $py -m pip install -r requirements_desktop.txt

function Remove-DirWithRetry {
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [int]$Attempts = 6,
        [int]$DelayMs = 700
    )
    if (-not (Test-Path $Path)) { return }

    for ($i = 1; $i -le $Attempts; $i++) {
        try {
            Remove-Item $Path -Recurse -Force -ErrorAction Stop
            if (-not (Test-Path $Path)) { return }
        }
        catch {
            if ($i -eq $Attempts) {
                throw "Nao foi possivel remover '$Path'. Feche o app/Explorer/antivirus que esteja usando arquivos desta pasta e rode novamente."
            }
            Start-Sleep -Milliseconds $DelayMs
        }
    }
}

Write-Host "==> Limpando builds antigos"
# Garante que o executavel antigo nao esteja travando arquivos em dist/
Get-Process -Name "RotaHubDesktop" -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
Start-Sleep -Milliseconds 500

Remove-DirWithRetry -Path "build"
Remove-DirWithRetry -Path "dist"

Write-Host "==> Gerando EXE (PyInstaller)"
& $py -m PyInstaller `
  --noconfirm `
  --clean `
  --windowed `
  --name "RotaHubDesktop" `
  --icon "assets\app_icon.ico" `
  --hidden-import "pandas" `
  --hidden-import "openpyxl" `
  --hidden-import "xlrd" `
  --collect-all "pandas" `
  --collect-all "openpyxl" `
  --collect-all "xlrd" `
  --add-data "assets;assets" `
  --add-data "certificados;certificados" `
  --add-data "rota_granja.db;." `
  main.py

Write-Host "==> Build concluido em: dist\RotaHubDesktop\RotaHubDesktop.exe"
Write-Host "==> Versao informada: $AppVersion"
Write-Host "==> Proximo passo: gerar instalador com Inno Setup (installer\rotahub.iss)"
