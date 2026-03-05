# Plano De Migracao 100% Online

## Objetivo
Migrar o sistema para modelo `server-first`, onde:
- API + banco central são a unica fonte de verdade.
- Desktop e APK motorista atuam como clientes.
- Nenhuma operacao critica depende de escrita local como verdade final.

## Estado Atual (Resumo)
- Desktop ainda grava SQLite local em varios fluxos de negocio (`programacoes`, `recebimentos`, `despesas`, `clientes`).
- API ja cobre parte importante (cadastros, rotas, monitoramento, avulsas, eventos do app motorista).
- Arquitetura atual e hibrida (local + sync), nao 100% online.

## Meta Arquitetural
1. Banco central (PostgreSQL recomendado) como unico banco oficial.
2. API executa 100% das regras de negocio e validacoes.
3. Desktop remove `INSERT/UPDATE/DELETE` de negocio local.
4. APK motorista opera offline-first com fila e idempotencia.
5. Sincronizacao entre estacoes em near real-time.

## Fases De Execucao
1. Fase 0 - Contratos e seguranca
2. Fase 1 - Cadastros 100% API
3. Fase 2 - Programacao 100% API
4. Fase 3 - Recebimentos/Despesas/Prestacao 100% API
5. Fase 4 - Relatorios e monitoramento centralizados
6. Fase 5 - Corte final do SQLite como fonte de verdade

## Fase 0 - Contratos E Seguranca
- Definir schema canonicamente no servidor: `motoristas`, `veiculos`, `ajudantes`, `clientes`, `programacoes`, `programacao_itens`, `recebimentos`, `despesas`, `transferencias`, `substituicoes`, `rota_gps_pings`, `event_log`.
- Definir padrao de erro da API (`code`, `message`, `details`, `trace_id`).
- Definir idempotencia para escrita critica (`X-Idempotency-Key`).
- Definir controle de concorrencia (`version` ou `updated_at`).
- Definir autenticacao e autorizacao por perfil.

## Fase 1 - Cadastros 100% API
- Desktop:
1. Substituir salvar/editar/excluir local de `motoristas`, `veiculos`, `ajudantes`, `clientes`, `usuarios` por endpoints API.
2. Manter cache local apenas para leitura opcional.
- API:
1. Confirmar CRUD completo para todos os cadastros.
2. Garantir regras de unicidade e validacoes no backend.

## Fase 2 - Programacao 100% API
- Desktop:
1. `salvar_programacao` nao escreve mais local como fonte final.
2. Criacao/edicao/reabertura/finalizacao de programacao via API.
3. Marcacao de vendas usadas e vinculo de clientes via API transacional.
- API:
1. Endpoint transacional para upsert de programacao + itens.
2. Endpoint de reconciliacao e consistencia de vinculos.
3. Endpoint de status da programacao com versionamento.

## Fase 3 - Recebimentos/Despesas/Prestacao 100% API
- Desktop:
1. Inserir/editar/excluir recebimentos e despesas via API.
2. Fechamento de prestacao somente via API.
3. Bloqueios de negocio baseados no estado retornado pela API.
- API:
1. CRUD de recebimentos.
2. CRUD de despesas.
3. Endpoint de fechamento de prestacao com validacao logistica.
4. Endpoint de reabertura com trilha de auditoria.

## Fase 4 - Relatorios E Monitoramento
- Relatorios gerados do banco central.
- Monitoramento de rotas e status operacional 100% server-side.
- Opcional: endpoint para exportacao pronta de relatorios (CSV/XLSX/PDF).

## Fase 5 - Corte Final Do Local
- Desligar escrita de negocio local no Desktop.
- SQLite local vira somente:
1. cache de leitura
2. fila tecnica (se necessario)
- Remover rotinas de reconciliacao que tratam local como fonte de verdade.

## Mapa De Gaps Principais (Hoje)
- Ja existe na API:
1. Cadastros base desktop (`/desktop/cadastros/*`)
2. Upsert de rota (`/desktop/rotas/upsert`)
3. Avulsas (`/desktop/avulsas*`)
4. Fluxo motorista (login, iniciar, finalizar, status, gps, carregamento, substituicao, transferencia)
- Falta/insuficiente para 100% online:
1. CRUD completo de `recebimentos` no servidor
2. CRUD completo de `despesas` no servidor
3. Fechamento/reabertura de prestacao server-side completo
4. CRUD de `programacoes` com regras completas do Desktop
5. Endpoints de relatorios administrativos centrais
6. Endpoint para `usuarios`/permissoes administrativos

## Criterios De Aceite Por Fase
- Fase 1 pronta quando:
1. Nenhuma tela de cadastro usa `INSERT/UPDATE/DELETE` local como verdade final.
2. API recusa dados invalidos mesmo com cliente adulterado.
- Fase 2 pronta quando:
1. Programacao completa e itens vivem no servidor.
2. Conflitos de edicao entre estacoes sao detectados.
- Fase 3 pronta quando:
1. Recebimentos/despesas/prestacao persistem apenas no servidor.
2. Fechamento e reabertura geram auditoria.
- Fase 5 pronta quando:
1. Queda de uma estacao nao causa divergencia estrutural de dados.
2. Nova estacao em qualquer rede autenticada enxerga o mesmo estado.

## Plano Tecnico Imediato (Execucao)
1. Criar contrato OpenAPI para os recursos faltantes (`programacoes`, `recebimentos`, `despesas`, `prestacoes`, `usuarios`).
2. Implementar no `api_server.py` os endpoints faltantes de `recebimentos` e `despesas`.
3. Refatorar `main.py` para trocar as escritas locais desses dois modulos por chamadas `_call_api`.
4. Implementar idempotencia nas rotas de escrita.
5. Adicionar testes de integracao API para:
1. salvar programacao
2. salvar recebimento
3. salvar despesa
4. fechar prestacao
6. Repetir a mesma estrategia modulo a modulo ate zerar escrita local de negocio.

## Riscos E Mitigacoes
- Risco: divergencia durante migracao.
- Mitigacao: feature flag por modulo (`USE_API_<MODULO>=1`) e rollout gradual.
- Risco: duplicidade por retransmissao.
- Mitigacao: idempotencia obrigatoria em escrita.
- Risco: regressao de regras antigas.
- Mitigacao: testes de regressao com cenarios reais.

## Proximo Passo Recomendado
Implementar agora a Fase 3 inicial:
1. API: CRUD de `recebimentos` e `despesas`.
2. Desktop: tela `RecebimentosPage` e `DespesasPage` sem `INSERT/UPDATE/DELETE` local de negocio.
