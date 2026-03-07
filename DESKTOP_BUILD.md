# Build e Instalacao do RotaHub Desktop (Windows)

## Objetivo
Gerar um executavel instalavel do `main.py` sem perder dados em atualizacoes.

## O que foi ajustado
- O `main.py` agora usa configuracao por ambiente:
  - **Development** (`python main.py`): banco isolado em `.rotahub_runtime\desktop\development\dev-local\`.
  - **Staging/externo** (`.exe` com `config.json` ajustado): banco isolado por tenant em `%LOCALAPPDATA%\RotaHubDesktop\desktop\staging\<tenant>\`.
  - **Production** (`.exe`/servidor): banco isolado por tenant em `%LOCALAPPDATA%\RotaHubDesktop\desktop\production\<tenant>\`.
- O executavel nao leva mais banco de desenvolvimento embutido.
- Atualizacao do app nao substitui os bancos persistentes.

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

## 4) Configuracao externa do `.exe`
Na primeira execucao, o desktop gera `%LOCALAPPDATA%\RotaHubDesktop\config.json`.

Ajuste esse arquivo para:
- `app_env`: `staging` ou `production`
- `tenant.tenant_id` / `tenant.company_id`
- `runtime.db_path` ou `runtime.data_root`
- `api.base_url`
- `runtime.desktop_secret`

Use `config\desktop.runtime.example.json` como referencia.

## 5) URL da API online via variavel de ambiente
Para override rapido:

```powershell
setx ROTA_SERVER_URL "https://rotahub-api.onrender.com" /M
```

Reabra o app apos definir.
