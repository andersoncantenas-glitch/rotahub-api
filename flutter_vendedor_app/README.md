# RotaHub Vendedor

Projeto Flutter separado para vendedores operarem no mesmo ecossistema do
desktop `main.py`, usando a mesma API central deste repositorio.

## Arquitetura de comunicacao

- Fonte de verdade compartilhada: `api_server.py`
- Desktop e Flutter precisam apontar para a mesma `baseUrl`
- Endpoints `desktop/*` usam o mesmo `X-Desktop-Secret`
- Endpoints `auth/vendedor/login` e `vendedor/*` usam token `Bearer`
- Cache offline existe apenas como resiliencia temporaria

Observacoes:

- O nome antigo `flutter_application_1` ainda aparece em alguns artefatos
  gerados de plataforma, mas o projeto funcional aqui e `flutter_vendedor_app`.
- A arquitetura unificada esta documentada em
  `ARQUITETURA_COMUNICACAO_UNIFICADA.md`.

## Escopo atual

- Configuracao local de `baseUrl`, `X-Desktop-Secret`, vendedor e cidade base
- Primeira sincronizacao obrigatoria para baixar cadastros
- Cadastro de avulsa com motorista, veiculo, equipe, local e clientes
- Lista e detalhe de avulsas
- Fila offline para criacao e conciliacao futura
- Rascunho compartilhado, pre-programacoes e programacao oficial

## Como usar

1. Entre na pasta `flutter_vendedor_app`
2. Rode `flutter pub get`
3. Rode `flutter run --dart-define=ROTA_SERVER_URL=... --dart-define=ROTA_SECRET=...`

Tambem sao aceitos os aliases legados `VENDOR_API_BASE_URL` e
`VENDOR_DESKTOP_SECRET`, mas o preferido agora e usar os mesmos nomes do
ambiente compartilhado do desktop/API.

## Regras de integracao

- O app Flutter deve usar a mesma API do desktop `main.py`
- O service de referencia deve seguir o contrato real de
  `lib/services/vendedor_api_service.dart`
- Nenhum payload compartilhado deve existir apenas localmente no app
- Ao reconectar, a fila offline deve ser sincronizada com a API central

## Observacoes operacionais

- O app exige o primeiro acesso online para habilitar o cache local
- Se `baseUrl`, `secret`, vendedor ou cidade mudarem, a primeira sincronizacao
  e exigida novamente
- Neste ambiente o SDK Flutter nao estava no `PATH`, entao a estrutura foi
  revisada sem executar build local aqui
