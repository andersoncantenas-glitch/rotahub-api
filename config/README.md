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
