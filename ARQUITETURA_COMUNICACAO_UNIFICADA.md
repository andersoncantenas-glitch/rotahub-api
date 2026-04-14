# Arquitetura de Comunicacao Unificada

## Objetivo

`main.py`, `flutter_application_1` (projeto atual `flutter_vendedor_app`) e `APP_VENDEDOR_API`
precisam operar sobre a mesma comunicacao e os mesmos dados compartilhados.

Neste repositorio, isso significa:

- Uma unica API central: `api_server.py`
- Uma unica base persistida por ambiente/tenant, resolvida pelo `runtime_config.py`
- Um unico contrato de endpoints para leitura/escrita compartilhada
- Cache local apenas como resiliencia, nunca como fonte de verdade de dados compartilhados

## Componentes

### 1. Desktop `main.py`

- Carrega configuracao de runtime com `load_app_config("desktop")`
- Usa `APP_CONFIG.api_base_url` e `ROTA_SECRET`
- Chama a API central pelos endpoints `desktop/*`
- Continua podendo ter banco local/runtime, mas os fluxos compartilhados devem refletir a API central

### 2. API central `api_server.py`

- E o hub de comunicacao entre desktop e Flutter
- Expone:
  - `desktop/*` para operacoes de cadastro, programacoes, avulsas e consultas operacionais
  - `auth/vendedor/login` para sessao do vendedor
  - `vendedor/*` para rascunho e pre-programacoes
- E a fonte de verdade para os dados compartilhados entre desktop e apps

### 3. Flutter `flutter_application_1`

- No repositorio, o projeto real e `flutter_vendedor_app`
- Ainda existem artefatos gerados com o nome antigo `flutter_application_1`
- Deve apontar para a mesma `baseUrl` da API central usada pelo desktop
- Usa:
  - `X-Desktop-Secret` para endpoints `desktop/*`
  - `Bearer token` para endpoints `auth/vendedor/login` e `vendedor/*`

### 4. `APP_VENDEDOR_API_SERVICE_EXEMPLO.dart`

- E um arquivo de referencia para a integracao Flutter
- Deve seguir o mesmo contrato do service real em:
  - `flutter_vendedor_app/lib/services/vendedor_api_service.dart`
- Nao deve inventar endpoints ou payloads paralelos

## Fonte de verdade

Para manter os tres componentes sincronizados, a regra e:

- Cadastros compartilhados: API central
- Programacoes oficiais: API central
- Avulsas: API central
- Rascunho/pre-programacoes do vendedor: API central

Somente os caches offline locais podem divergir temporariamente.
Depois da reconexao, eles devem sincronizar de volta com a API central.

## Contrato de comunicacao

### Desktop -> API

- `main.py` le e grava via `desktop/*`
- Header obrigatorio:
  - `X-Desktop-Secret: <ROTA_SECRET>`

### Flutter vendedor -> API

- Bootstrap/config usa a mesma `baseUrl` do backend central
- Endpoints `desktop/*`:
  - usam `X-Desktop-Secret`
- Endpoints de sessao/rascunho:
  - `POST /auth/vendedor/login`
  - `GET/POST/PATCH/DELETE /vendedor/*`
  - usam token `Bearer`

### Service de exemplo -> mesma API

- O arquivo `APP_VENDEDOR_API_SERVICE_EXEMPLO.dart` deve replicar essa mesma logica
- Ele nao pode apontar para outra API ou usar um healthcheck diferente do app real sem necessidade

## Regras obrigatorias de alinhamento

1. `main.py` e `flutter_vendedor_app` precisam usar a mesma `api_base_url`
2. O `desktop_secret` do desktop e o do Flutter precisam ser do mesmo ambiente
3. Toda escrita compartilhada precisa passar pela API central
4. Nenhum fluxo compartilhado deve manter uma "versao local oficial" separada da API
5. Cache offline so pode ser temporario e precisa ter sincronizacao/flush
6. Contratos de payload/response devem ser unificados entre:
   - `api_server.py`
   - `flutter_vendedor_app/lib/services/vendedor_api_service.dart`
   - `APP_VENDEDOR_API_SERVICE_EXEMPLO.dart`

## Fluxo recomendado

1. Desktop configura ambiente, `api_base_url` e `desktop_secret`
2. API central sobe com o schema compartilhado
3. Flutter vendedor aponta para a mesma API
4. Flutter autentica vendedor e consome/escreve os mesmos dados do desktop
5. Offline no Flutter gera apenas fila temporaria
6. Ao reconectar, a fila volta para a API central e fica visivel para o desktop

## Arquivos de referencia

- Desktop runtime/config:
  - `main.py`
  - `runtime_config.py`
  - `config/desktop.runtime.json`
- API central:
  - `api_server.py`
- Flutter vendedor:
  - `flutter_vendedor_app/lib/core/default_app_config.dart`
  - `flutter_vendedor_app/lib/services/vendedor_api_service.dart`
- Referencia de integracao:
  - `APP_VENDEDOR_API_SERVICE_EXEMPLO.dart`

## Decisao arquitetural

Se os tres componentes precisarem "compartilhar as mesmas informacoes", a arquitetura correta e:

- `api_server.py` como hub e fonte de verdade
- `main.py` e Flutter como clientes da mesma API
- Nenhum contrato paralelo fora dessa API

