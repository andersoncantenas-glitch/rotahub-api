# Deploy do RotaHub Web

Este guia deixa o sistema pronto para migrar para uma VPS ou hospedagem Python. Para Hostinger, prefira VPS. Hospedagem compartilhada normalmente nao e ideal para FastAPI/uvicorn rodando continuamente.

## O que precisa ser persistente

Mantenha estes caminhos fora da pasta que pode ser substituida a cada deploy:

- Banco SQLite: `ROTA_DB=/var/rotahub/data/rotadb.db`
- Fotos do app motorista: `ROTA_MOBILE_PHOTOS_DIR=/var/rotahub/data/fotos_rotas`

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
uvicorn backend.main:app --host 0.0.0.0 --port 8000
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
ExecStart=/opt/rotahub/.venv/bin/uvicorn backend.main:app --host 0.0.0.0 --port 8000
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

## Backup antes de migrar

SQLite:

```bash
sqlite3 /var/rotahub/data/rotadb.db ".backup '/var/rotahub/data/backup_rotadb.db'"
tar -czf fotos_rotas_backup.tar.gz -C /var/rotahub/data fotos_rotas
```

Na migracao inicial, copie:

- `rotadb.db`
- pasta `fotos_rotas`
- `.env` ajustado

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
