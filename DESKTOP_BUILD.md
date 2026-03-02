# Build e Instalacao do RotaHub Desktop (Windows)

## Objetivo
Gerar um executavel instalavel do `main.py` sem perder dados em atualizacoes.

## O que foi ajustado
- O `main.py` agora usa:
  - **Dev** (`python main.py`): banco local da pasta do projeto.
  - **Instalado** (PyInstaller): banco em `%LOCALAPPDATA%\RotaHubDesktop\rota_granja.db`.
- Isso evita sobrescrever banco ao atualizar o app instalado.

## 1) Gerar EXE
No PowerShell, dentro de `C:\pdc_rota`:

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\build_desktop.ps1 -AppVersion 1.0.0
```

Saida esperada:
- `dist\RotaHubDesktop\RotaHubDesktop.exe`

## 2) Gerar instalador (Inno Setup)
1. Instale o Inno Setup (se ainda nao tiver).
2. Abra `installer\rotahub.iss`.
3. Ajuste `MyAppVersion`.
4. Compile.

Saida esperada:
- `dist_installer\RotaHubDesktop_Setup_1.0.0.exe`

## 3) Atualizacao sem perder dados
- Gere nova versao com `-AppVersion`.
- Reinstale por cima com novo setup.
- O banco do usuario permanece em `%LOCALAPPDATA%\RotaHubDesktop`.

## 4) URL da API online no desktop
Para usar backend online:

```powershell
setx ROTA_SERVER_URL "https://rotahub-api.onrender.com" /M
```

Reabra o app apos definir.
