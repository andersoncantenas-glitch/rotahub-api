`config/environments/` define os defaults de `development`, `staging` e `production`.

`python main.py`
- usa `development` por padrão
- grava dados em `.rotahub_runtime/desktop/development/dev-local/`
- não sincroniza com servidor por padrão

`.exe` instalado
- usa `production` por padrão
- grava dados em `%LOCALAPPDATA%\RotaHubDesktop\desktop\<env>\<tenant>\`
- lê configuração externa de `%LOCALAPPDATA%\RotaHubDesktop\config.json`

Arquivos de exemplo
- `config/desktop.runtime.example.json`: estação externa/homologação
- `config/server.runtime.example.json`: servidor por tenant/empresa
- `config/server.runtime.json`: runtime publicado isolado do development

Regras de isolamento:
- `development` (`python main.py`) usa somente `.rotahub_runtime/desktop/development/...`
- `server` usa somente `.rotahub_runtime/server/<env>/.../rotahub_server.db`
- `desktop` publicado usa somente `%LOCALAPPDATA%\\RotaHubDesktop\\...\\rotahub_desktop.db` e a API do ambiente publicado
- `development` nao pode sincronizar dados nem publicar inserts/updates/deletes para staging/producao
- bancos locais, runtime local e snapshots nao devem subir para Git/Render

Bootstrap do servidor:
- `python init_server_db.py --reset --admin-pass SUA_SENHA`
- cria/atualiza o schema do banco publicado
- remove dados operacionais existentes
- preserva somente o usuario `ADMIN`
