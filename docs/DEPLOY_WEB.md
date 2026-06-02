# Deploy do RotaHub Web

Este guia deixa o sistema pronto para migrar para uma VPS ou hospedagem Python. Para Hostinger, prefira VPS. Hospedagem compartilhada normalmente nao e ideal para FastAPI/uvicorn rodando continuamente.

## O que precisa ser persistente

Mantenha estes caminhos fora da pasta que pode ser substituida a cada deploy:

- Banco SQLite: `ROTA_DB=/var/rotahub/data/rotadb.db`
- Fotos do app motorista: `ROTA_MOBILE_PHOTOS_DIR=/var/rotahub/data/fotos_rotas`
- Backups: `BACKUP_DIR=/var/rotahub/data/backups`
- Pacotes de migracao: `ROTA_EXPORT_DIR=/var/rotahub/data/exports`

Se usar PostgreSQL, mantenha `DATABASE_URL=postgresql://...` e ainda preserve `ROTA_MOBILE_PHOTOS_DIR`.

## Variaveis obrigatorias

Use `.env.production.example` como base:

```bash
cp .env.production.example .env
```

Ajuste no servidor:

- `ENVIRONMENT=production`
- `DEBUG=false`
- `ALLOWED_HOSTS=seudominio.com.br,www.seudominio.com.br,127.0.0.1,localhost`
- `CORS_ORIGINS=https://seudominio.com.br,https://www.seudominio.com.br`
- `SECRET_KEY` com valor longo
- `JWT_SECRET_KEY` com outro valor longo
- `ROTA_SECRET` com segredo do app motorista
- `ROTA_ENABLE_LEGACY_MOBILE_API=1`
- `ROTA_DB` ou `DATABASE_URL`
- `ROTA_MOBILE_PHOTOS_DIR`
- `BACKUP_DIR`
- `ROTA_EXPORT_DIR`

## Instalar em uma VPS Linux

Exemplo base:

```bash
sudo mkdir -p /opt/rotahub /var/rotahub/data/fotos_rotas
sudo chown -R $USER:$USER /opt/rotahub /var/rotahub
cd /opt/rotahub

python3 -m venv .venv
source .venv/bin/activate
pip install --upgrade pip
pip install -r backend/requirements.txt
```

Antes de subir:

```bash
python scripts/check_deploy_ready.py
```

Subida manual para teste:

```bash
python scripts/server_start.py
```

Acesse:

- Sistema web: `http://IP_DO_SERVIDOR:8000/app/index.html`
- Healthcheck: `http://IP_DO_SERVIDOR:8000/health`
- Prontidao de banco/fotos: `http://IP_DO_SERVIDOR:8000/ready`

## systemd

Crie `/etc/systemd/system/rotahub.service`:

```ini
[Unit]
Description=RotaHub Web
After=network.target

[Service]
WorkingDirectory=/opt/rotahub
EnvironmentFile=/opt/rotahub/.env
ExecStart=/opt/rotahub/.venv/bin/python scripts/server_start.py
Restart=always
RestartSec=5
User=rotahub

[Install]
WantedBy=multi-user.target
```

Depois:

```bash
sudo systemctl daemon-reload
sudo systemctl enable rotahub
sudo systemctl start rotahub
sudo systemctl status rotahub
```

## Nginx com HTTPS

Use Nginx como proxy para a porta `8000`:

```nginx
server {
    server_name seudominio.com.br www.seudominio.com.br;

    client_max_body_size 20M;

    location / {
        proxy_pass http://127.0.0.1:8000;
        proxy_set_header Host $host;
        proxy_set_header X-Real-IP $remote_addr;
        proxy_set_header X-Forwarded-For $proxy_add_x_forwarded_for;
        proxy_set_header X-Forwarded-Proto $scheme;
    }
}
```

Depois instale certificado HTTPS com Certbot.

## App motorista

No Flutter, a URL da API deve apontar para o dominio HTTPS do servidor. Mantenha o mesmo `ROTA_SECRET` no servidor e no app quando o endpoint legado exigir segredo.

Rotas importantes:

- Web: `/app/index.html`
- API web: `/api/v1`
- Compatibilidade app motorista: endpoints montados na raiz quando `ROTA_ENABLE_LEGACY_MOBILE_API=1`

## Backup portatil antes de migrar

Com o ambiente configurado, gere um pacote unico:

```bash
python scripts/export_runtime_backup.py
```

O arquivo `rotahub_migration_DATA_HORA.zip` sera criado em `ROTA_EXPORT_DIR`. Ele inclui:

- copia consistente do SQLite criada pela API de backup do proprio SQLite;
- pasta `fotos_rotas`;
- `manifest.json` com os caminhos esperados para restauracao.

Baixe esse arquivo para outro computador antes de desligar ou publicar novamente no Render free. O disco `/tmp` do Render free e temporario.

Sem acesso ao terminal, entre no sistema web como administrador, abra `Backup / Exportar` e use `BAIXAR PACOTE DE MIGRACAO`. O download usa o mesmo formato ZIP.

Para escolher os caminhos manualmente:

```bash
python scripts/export_runtime_backup.py \
  --db /var/rotahub/data/rotadb.db \
  --photos-dir /var/rotahub/data/fotos_rotas \
  --output-dir /var/rotahub/data/exports
```

## Backup manual alternativo

SQLite:

```bash
sqlite3 /var/rotahub/data/rotadb.db ".backup '/var/rotahub/data/backup_rotadb.db'"
tar -czf fotos_rotas_backup.tar.gz -C /var/rotahub/data fotos_rotas
```

Na migracao inicial, copie:

- o conteudo `database/rotadb.db` do pacote para `/var/rotahub/data/rotadb.db`
- a pasta `fotos_rotas` do pacote para `/var/rotahub/data/fotos_rotas`
- `.env` ajustado

## Render free durante os testes

O `render.yaml` usa `/tmp/rotahub` porque o plano gratuito nao oferece disco persistente. Isso serve para testes, mas os dados podem desaparecer ao reiniciar ou publicar novamente.

Na futura VPS/Hostinger, nao altere o codigo. Configure `.env` usando `.env.production.example`, aponte `ROTA_DATA_ROOT`, `ROTA_DB`, `BACKUP_DIR`, `ROTA_EXPORT_DIR` e `ROTA_MOBILE_PHOTOS_DIR` para `/var/rotahub/data`, e suba com:

```bash
python scripts/server_start.py
```

## Checklist final

Rode:

```bash
python scripts/check_deploy_ready.py
```

Confirme:

- `/health` responde `healthy`
- `/ready` responde `ready`
- login web funciona
- app motorista sincroniza uma rota de teste
- foto enviada pelo app aparece em Despesas/Mortalidade
- PDF de prestacao abre com os dados da rota
