param(
    [string]$AppVersion = ""
)

$ErrorActionPreference = "Stop"

$root = Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path)
Set-Location $root

Write-Host "==> Projeto: $root"

if ([string]::IsNullOrWhiteSpace($AppVersion)) {
    $versionPy = Join-Path $root "version.py"
    if (-not (Test-Path $versionPy)) {
        throw "Arquivo version.py nao encontrado em $versionPy"
    }
    $match = Select-String -Path $versionPy -Pattern 'APP_VERSION\s*=\s*"(\d+\.\d+\.\d+)"' | Select-Object -First 1
    if (-not $match) {
        throw "Nao foi possivel identificar APP_VERSION em version.py"
    }
    $AppVersion = $match.Matches[0].Groups[1].Value
}

if (-not (Test-Path ".venv")) {
    Write-Host "==> Criando .venv"
    python -m venv .venv
}

$py = Join-Path $root ".venv\Scripts\python.exe"
if (-not (Test-Path $py)) {
    throw "Python da venv nao encontrado em $py"
}

$pythonBase = & $py -c "import sys; print(sys.base_prefix)"
if ([string]::IsNullOrWhiteSpace($pythonBase)) {
    throw "Nao foi possivel identificar o Python base usado pela venv."
}
$stdlibTkinter = Join-Path $pythonBase "Lib\tkinter"
if (-not (Test-Path $stdlibTkinter)) {
    throw "Pacote tkinter nao encontrado em $stdlibTkinter"
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

$tkInfo = & $py -c "import os, sys; base = sys.base_prefix; tcl_root = os.path.join(base, 'tcl'); pairs = [];`nfor name in sorted(os.listdir(tcl_root)) if os.path.isdir(tcl_root) else []:`n    full = os.path.join(tcl_root, name)`n    if os.path.isdir(full) and (name.startswith('tcl8.') or name.startswith('tk8.')):`n        pairs.append(f'{name}={full}')`nprint('\n'.join(pairs))"
$tkMap = @{}
foreach ($line in $tkInfo) {
    if ([string]::IsNullOrWhiteSpace($line)) { continue }
    $parts = $line -split '=', 2
    if ($parts.Count -eq 2) {
        $tkMap[$parts[0]] = $parts[1]
    }
}

if (-not $tkMap.ContainsKey("tcl8.6") -or -not $tkMap.ContainsKey("tk8.6")) {
    throw "Nao foi possivel localizar as bibliotecas Tcl/Tk do Python em $($py)."
}

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
  --hidden-import "tkinter" `
  --hidden-import "_tkinter" `
  --collect-all "pandas" `
  --collect-all "openpyxl" `
  --collect-all "xlrd" `
  --runtime-hook "scripts\pyi_rth_tkinter.py" `
  --add-data "$stdlibTkinter;tkinter" `
  --add-data "assets;assets" `
  --add-data "certificados;certificados" `
  --add-data "config;config" `
  --add-data "$($tkMap['tcl8.6']);tcl\tcl8.6" `
  --add-data "$($tkMap['tk8.6']);tcl\tk8.6" `
  main.py

Write-Host "==> Build concluido em: dist\RotaHubDesktop\RotaHubDesktop.exe"
Write-Host "==> Versao do codigo-fonte: $AppVersion"
Write-Host "==> Proximo passo: gerar instalador com Inno Setup (installer\rotahub.iss)"
Write-Host "==> Gere o setup somente apos este build, para evitar empacotar um dist antigo."
