# Implantacao Cliente Novo (Zero Dados)

Objetivo: instalar Desktop + APK para uma nova empresa sem herdar dados de clientes anteriores.

## 1) Padrao de isolamento

Cada empresa deve ter:
- Banco proprio (`ROTA_DB` exclusivo)
- Segredo proprio (`ROTA_SECRET` exclusivo)
- URL propria da API (idealmente um servico Render por empresa)

Nunca reutilize o mesmo banco entre empresas.

## 2) Servidor (Render) - empresa nova

### 2.1 Environment Variables

Configure no servico Render da empresa:

- `ROTA_DB=/var/data/empresa_nome.db`
- `ROTA_SECRET=<segredo_forte_unico>`
- `ROTA_CORS_ORIGINS=https://app-da-empresa.com` (ou vazio se nao usar web)

Observacao:
- Se nao houver dominio web, voce pode manter `ROTA_CORS_ORIGINS` vazio.
- Para persistencia real, o servico deve ter Disk no Render.

### 2.2 Start Command

Use este comando para criar o banco automaticamente se nao existir:

```bash
python -c "import os,sqlite3; p=os.environ.get('ROTA_DB','rota_granja.db'); os.makedirs(os.path.dirname(p) or '.', exist_ok=True); sqlite3.connect(p).close()" && python -m uvicorn api_server:app --host 0.0.0.0 --port $PORT
```

### 2.3 Validacao da API

Abra:

```text
https://<url-da-api>/ping
```

Resultado esperado:
- `"ok": true`
- `"db": "/var/data/empresa_nome.db"` (ou caminho configurado no `ROTA_DB`)

## 3) Gerar Desktop (.exe + setup)

Terminal VSCode em `C:\pdc_rota`:

```powershell
cd C:\pdc_rota
powershell -ExecutionPolicy Bypass -File .\scripts\build_desktop.ps1 -AppVersion 1.0.0
```

Saida:
- `dist\RotaHubDesktop\RotaHubDesktop.exe`

Para gerar instalador:
- Abrir `installer\rotahub.iss` no Inno Setup
- Compilar

Saida:
- `dist_installer\RotaHubDesktop_Setup_<versao>.exe`

## 4) Instalar Desktop na empresa

No computador da empresa, instale o setup e configure:

Terminal PowerShell:

```powershell
setx ROTA_SERVER_URL "https://<url-da-api-da-empresa>"
setx ROTA_SECRET "<mesmo-segredo-do-render>"
setx ROTA_DESKTOP_SYNC_API "1"
```

Feche e abra o Desktop novamente.

## 5) Gerar APK (sem emulador)

Terminal VSCode em `C:\flutter_application_1`:

```powershell
cd C:\flutter_application_1
flutter clean
flutter pub get
flutter build apk --release --dart-define=API_BASE_URL=https://<url-da-api-da-empresa>
```

Saida:
- `build\app\outputs\flutter-apk\app-release.apk`

Instale esse APK no celular da empresa.

## 6) Limpeza de cache para troca de empresa

Se o mesmo computador/celular foi usado em testes anteriores:

- Desktop: apagar `%LOCALAPPDATA%\RotaHubDesktop\rota_granja.db` (opcional, recomendado)
- Celular: limpar dados do app ou reinstalar APK

## 7) Checklist de homologacao (empresa nova)

1. Login admin no Desktop.
2. Cadastrar motorista, ajudantes, veiculo e clientes.
3. Importar vendas.
4. Criar programacao nova.
5. Login do motorista no celular.
6. Confirmar rota no app.
7. Iniciar rota, carregar, alterar status operacional.
8. Finalizar rota no app.
9. Confirmar liberacao em Recebimentos no Desktop.
10. Fechar prestacao sem erros.

## 8) Comandos de diagnostico rapido

### API ativa e banco em uso

```text
https://<url-da-api>/ping
```

### Se aparecer dado antigo no Desktop

1. Confirmar `API: ONLINE` e URL correta.
2. Confirmar `ROTA_SERVER_URL` e `ROTA_SECRET` da estacao.
3. Confirmar `/ping` aponta para o banco da empresa atual.
4. Limpar cache local do Desktop/celular e reabrir.

